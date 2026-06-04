Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports NAudio.CoreAudioApi
Imports NAudio.Wave

Namespace Transcription

    ''' <summary>
    ''' Captures microphone and / or system-audio (WASAPI loopback), resamples both to 16 kHz mono 16-bit PCM,
    ''' mixes them with a soft AGC/limiter, and raises Frame events. Optionally produces interleaved stereo
    ''' (mic = L, system = R) when MultiChannelStereo = True for engines that support multi-channel diarization.
    ''' Never persists audio unless AudioDebugDump = True (writes a single rotating WAV in %TEMP%).
    ''' </summary>
    Public Class AudioCaptureService
        Implements IDisposable

        Public Class FrameEventArgs
            Inherits EventArgs
            Public Property Pcm As Byte()
            Public Property BytesValid As Integer
            Public Sub New(pcm As Byte(), bytes As Integer)
                Me.Pcm = pcm
                Me.BytesValid = bytes
            End Sub
        End Class

        Public Event Frame As EventHandler(Of FrameEventArgs)
        Public Event CaptureError As EventHandler(Of TranscriptionErrorEventArgs)

        Public Property MicDeviceIndex As Integer = 0
        Public Property SourceMode As AudioSourceMode = AudioSourceMode.MicrophoneOnly
        Public Property SystemAudioRenderDeviceId As String = ""
        Public Property MultiChannelStereo As Boolean = False
        Public Property AudioDebugDump As Boolean = False

        Private Const TARGET_SAMPLE_RATE As Integer = 16000

        Private _waveIn As WaveInEvent
        Private _loopbackCapture As WasapiLoopbackCapture
        Private _loopbackRawProvider As BufferedWaveProvider
        Private _loopbackResampler As MediaFoundationResampler
        Private _debugWriter As WaveFileWriter
        Private _running As Boolean

        Public Sub Start()
            If _running Then Return

            Dim targetFormat As New WaveFormat(TARGET_SAMPLE_RATE, 16, 1)

            If SourceMode <> AudioSourceMode.SystemAudioOnly Then
                _waveIn = New WaveInEvent() With {
                    .DeviceNumber = MicDeviceIndex,
                    .WaveFormat = targetFormat,
                    .BufferMilliseconds = 50
                }
                AddHandler _waveIn.DataAvailable, AddressOf OnMicData
            End If

            If SourceMode <> AudioSourceMode.MicrophoneOnly Then
                Dim device As MMDevice = Nothing
                If Not String.IsNullOrEmpty(SystemAudioRenderDeviceId) Then
                    Try
                        device = New MMDeviceEnumerator().GetDevice(SystemAudioRenderDeviceId)
                    Catch
                        device = Nothing
                    End Try
                End If
                _loopbackCapture = If(device IsNot Nothing, New WasapiLoopbackCapture(device), New WasapiLoopbackCapture())
                _loopbackRawProvider = New BufferedWaveProvider(_loopbackCapture.WaveFormat) With {.DiscardOnBufferOverflow = True}
                AddHandler _loopbackCapture.DataAvailable, Sub(s, e) _loopbackRawProvider.AddSamples(e.Buffer, 0, e.BytesRecorded)
                _loopbackResampler = New MediaFoundationResampler(_loopbackRawProvider, targetFormat) With {.ResamplerQuality = 60}
                Try
                    _loopbackCapture.StartRecording()
                Catch ex As Exception
                    RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs("Cannot start loopback: " & ex.Message, ex, False))
                    _loopbackCapture?.Dispose() : _loopbackCapture = Nothing
                    _loopbackResampler?.Dispose() : _loopbackResampler = Nothing
                End Try
            End If

            ' System-audio-only path: we need a pump because there's no mic-driven OnMicData.
            If SourceMode = AudioSourceMode.SystemAudioOnly Then StartSystemOnlyPump()

            If AudioDebugDump Then
                Try
                    Dim debugPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInk_AudioDebug.wav")
                    Dim writerFormat As WaveFormat = If(MultiChannelStereo, New WaveFormat(TARGET_SAMPLE_RATE, 16, 2), targetFormat)
                    _debugWriter = New WaveFileWriter(debugPath, writerFormat)
                Catch
                End Try
            End If

            If _waveIn IsNot Nothing Then _waveIn.StartRecording()
            _running = True
        End Sub

        Private _pumpCts As CancellationTokenSource
        Private Sub StartSystemOnlyPump()
            _pumpCts = New CancellationTokenSource()
            Dim ct = _pumpCts.Token
            Task.Run(Sub()
                         Const frameBytes As Integer = TARGET_SAMPLE_RATE * 2 \ 20 ' 50 ms
                         Dim buf(frameBytes - 1) As Byte
                         While Not ct.IsCancellationRequested
                             Try
                                 Dim n = If(_loopbackResampler Is Nothing, 0, _loopbackResampler.Read(buf, 0, frameBytes))
                                 If n > 0 Then EmitFrame(Nothing, 0, buf, n) Else Thread.Sleep(20)
                             Catch
                                 Thread.Sleep(20)
                             End Try
                         End While
                     End Sub)
        End Sub

        Private Sub OnMicData(sender As Object, e As WaveInEventArgs)
            If Not _running Then Return
            Try
                Dim sysBuf As Byte() = Nothing
                Dim sysBytes As Integer = 0
                If _loopbackResampler IsNot Nothing Then
                    sysBuf = New Byte(e.BytesRecorded - 1) {}
                    sysBytes = _loopbackResampler.Read(sysBuf, 0, e.BytesRecorded)
                End If
                EmitFrame(e.Buffer, e.BytesRecorded, sysBuf, sysBytes)
            Catch ex As Exception
                RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs("OnMicData: " & ex.Message, ex, False))
            End Try
        End Sub

        ''' <summary>Mix or interleave mic+system into the output frame and raise.</summary>
        Private _agcGain As Single = 1.0F
        Private Sub EmitFrame(mic As Byte(), micBytes As Integer, sys As Byte(), sysBytes As Integer)
            Dim outBuf As Byte()
            Dim outBytes As Integer

            If MultiChannelStereo AndAlso SourceMode = AudioSourceMode.MicrophoneAndSystem Then
                ' Interleave mic = L, sys = R
                Dim samples = Math.Max(micBytes, sysBytes) \ 2
                outBuf = New Byte(samples * 4 - 1) {}
                outBytes = samples * 4
                For i = 0 To samples - 1
                    Dim m As Short = If(mic IsNot Nothing AndAlso i * 2 + 1 < micBytes, BitConverter.ToInt16(mic, i * 2), CShort(0))
                    Dim s As Short = If(sys IsNot Nothing AndAlso i * 2 + 1 < sysBytes, BitConverter.ToInt16(sys, i * 2), CShort(0))
                    Buffer.BlockCopy(BitConverter.GetBytes(m), 0, outBuf, i * 4, 2)
                    Buffer.BlockCopy(BitConverter.GetBytes(s), 0, outBuf, i * 4 + 2, 2)
                Next
            Else
                ' Mono mix with soft AGC/limiter
                Dim primary = If(mic, sys)
                Dim primaryBytes = If(mic IsNot Nothing, micBytes, sysBytes)
                outBuf = New Byte(primaryBytes - 1) {}
                outBytes = primaryBytes
                Dim peak As Integer = 1
                For i = 0 To primaryBytes - 2 Step 2
                    Dim a As Integer = If(mic IsNot Nothing AndAlso i + 1 < micBytes, BitConverter.ToInt16(mic, i), 0)
                    Dim b As Integer = If(sys IsNot Nothing AndAlso i + 1 < sysBytes, BitConverter.ToInt16(sys, i), 0)
                    Dim sum As Integer = CInt((a + b) * _agcGain)
                    If sum > Short.MaxValue Then sum = Short.MaxValue
                    If sum < Short.MinValue Then sum = Short.MinValue
                    If Math.Abs(sum) > peak Then peak = Math.Abs(sum)
                    Dim bb = BitConverter.GetBytes(CShort(sum))
                    outBuf(i) = bb(0) : outBuf(i + 1) = bb(1)
                Next
                ' Slow AGC: target peak ~ 22000
                Dim desired As Single = 22000.0F / peak
                If desired < 1.0F Then
                    _agcGain = _agcGain * 0.9F + desired * 0.1F   ' attack
                Else
                    _agcGain = _agcGain * 0.995F + desired * 0.005F ' release
                End If
                If _agcGain < 0.25F Then _agcGain = 0.25F
                If _agcGain > 4.0F Then _agcGain = 4.0F
            End If

            If _debugWriter IsNot Nothing Then
                Try : _debugWriter.Write(outBuf, 0, outBytes) : Catch : End Try
            End If

            RaiseEvent Frame(Me, New FrameEventArgs(outBuf, outBytes))
        End Sub

        Public Sub [Stop]()
            If Not _running Then Return
            _running = False
            Try : _pumpCts?.Cancel() : Catch : End Try
            Try
                If _waveIn IsNot Nothing Then
                    RemoveHandler _waveIn.DataAvailable, AddressOf OnMicData
                    _waveIn.StopRecording() : _waveIn.Dispose() : _waveIn = Nothing
                End If
            Catch : End Try
            Try
                If _loopbackCapture IsNot Nothing Then
                    _loopbackCapture.StopRecording() : _loopbackCapture.Dispose() : _loopbackCapture = Nothing
                End If
            Catch : End Try
            Try : _loopbackResampler?.Dispose() : _loopbackResampler = Nothing : Catch : End Try
            Try : _debugWriter?.Dispose() : _debugWriter = Nothing : Catch : End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            [Stop]()
        End Sub

    End Class

End Namespace