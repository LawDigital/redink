' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: AudioCaptureService.vb
' Purpose: Provides a service for capturing audio from a specified input
'          device. It manages the audio stream and provides data to the
'          transcription engines.
'
' Architecture:
'  - Device Management: Enumerates available audio input devices and allows
'    for the selection of a specific device.
'  - Audio Streaming: Uses NAudio to capture microphone input and/or WASAPI
'    loopback from a selected render device.
'  - Format Normalization: Converts all capture paths to 16 kHz mono 16-bit PCM
'    before handing frames to the transcription engines.
'  - Lifecycle Management: Controls the starting and stopping of the audio
'    capture process.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.IO
Imports NAudio.CoreAudioApi
Imports NAudio.Wave

Namespace Transcription

    ''' <summary>
    ''' Captures microphone and / or system-audio (WASAPI loopback), converts both to
    ''' 16 kHz mono 16-bit PCM, mixes them with a soft AGC/limiter, and raises Frame
    ''' events. Optionally produces interleaved stereo (mic = L, system = R) when
    ''' MultiChannelStereo = True for engines that support multi-channel diarization.
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
        Private Const TARGET_BITS_PER_SAMPLE As Integer = 16
        Private Const TARGET_CHANNELS As Integer = 1
        Private Const TARGET_FRAME_MILLISECONDS As Integer = 50
        Private Const TARGET_FRAME_BYTES As Integer = TARGET_SAMPLE_RATE * 2 * TARGET_FRAME_MILLISECONDS \ 1000

        Private Shared ReadOnly PcmSubFormat As New Guid("00000001-0000-0010-8000-00aa00389b71")
        Private Shared ReadOnly FloatSubFormat As New Guid("00000003-0000-0010-8000-00aa00389b71")

        Private _waveIn As WaveInEvent
        Private _loopbackCapture As WasapiLoopbackCapture
        Private _debugWriter As WaveFileWriter
        Private _running As Boolean
        Private _agcGain As Single = 1.0F

        Private ReadOnly _systemPcmSyncRoot As New Object()
        Private ReadOnly _debugSyncRoot As New Object()
        Private _systemPcmQueue As New Queue(Of Byte)()

        Public Sub Start()
            If _running Then
                Return
            End If

            Dim targetFormat As New WaveFormat(TARGET_SAMPLE_RATE, TARGET_BITS_PER_SAMPLE, TARGET_CHANNELS)

            If SourceMode <> AudioSourceMode.SystemAudioOnly Then
                _waveIn = New WaveInEvent() With {
                    .DeviceNumber = MicDeviceIndex,
                    .WaveFormat = targetFormat,
                    .BufferMilliseconds = TARGET_FRAME_MILLISECONDS
                }

                AddHandler _waveIn.DataAvailable, AddressOf OnMicData
            End If

            If SourceMode <> AudioSourceMode.MicrophoneOnly Then
                Dim device As MMDevice = Nothing

                If Not String.IsNullOrEmpty(SystemAudioRenderDeviceId) Then
                    Try
                        device = New MMDeviceEnumerator().GetDevice(SystemAudioRenderDeviceId)
                    Catch ex As Exception
                        RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs(
                            "Selected output device is no longer available. Falling back to the default output device.",
                            ex,
                            False))
                        device = Nothing
                    End Try
                End If

                _loopbackCapture = If(device IsNot Nothing, New WasapiLoopbackCapture(device), New WasapiLoopbackCapture())
                AddHandler _loopbackCapture.DataAvailable, AddressOf OnLoopbackData
            End If

            If AudioDebugDump Then
                Try
                    Dim debugPath As String = Path.Combine(Path.GetTempPath(), "RedInk_AudioDebug.wav")
                    Dim writerFormat As WaveFormat = If(
                        MultiChannelStereo AndAlso SourceMode = AudioSourceMode.MicrophoneAndSystem,
                        New WaveFormat(TARGET_SAMPLE_RATE, TARGET_BITS_PER_SAMPLE, 2),
                        targetFormat)

                    _debugWriter = New WaveFileWriter(debugPath, writerFormat)
                Catch
                End Try
            End If

            _running = True

            If _loopbackCapture IsNot Nothing Then
                Try
                    _loopbackCapture.StartRecording()
                Catch ex As Exception
                    RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs("Cannot start loopback: " & ex.Message, ex, False))
                    Try
                        RemoveHandler _loopbackCapture.DataAvailable, AddressOf OnLoopbackData
                    Catch
                    End Try
                    Try
                        _loopbackCapture.Dispose()
                    Catch
                    End Try
                    _loopbackCapture = Nothing
                End Try
            End If

            If _waveIn IsNot Nothing Then
                _waveIn.StartRecording()
            End If
        End Sub

        Private Sub OnMicData(sender As Object, e As WaveInEventArgs)
            If Not _running Then
                Return
            End If

            Try
                Dim micBytes As Integer = e.BytesRecorded - (e.BytesRecorded Mod 2)
                If micBytes <= 0 Then
                    Return
                End If

                Dim micBuf(micBytes - 1) As Byte
                Buffer.BlockCopy(e.Buffer, 0, micBuf, 0, micBytes)

                If SourceMode = AudioSourceMode.MicrophoneOnly Then
                    EmitFrame(micBuf, micBytes, Nothing, 0)
                    Return
                End If

                Dim sysBuf As Byte() = DequeueSystemPcm(micBytes)
                Dim sysBytes As Integer = If(sysBuf IsNot Nothing, sysBuf.Length, 0)

                EmitFrame(micBuf, micBytes, sysBuf, sysBytes)
            Catch ex As Exception
                RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs("OnMicData: " & ex.Message, ex, False))
            End Try
        End Sub

        Private Sub OnLoopbackData(sender As Object, e As WaveInEventArgs)
            If Not _running OrElse _loopbackCapture Is Nothing Then
                Return
            End If

            Try
                Dim pcm As Byte() = ConvertToTargetPcm(e.Buffer, e.BytesRecorded, _loopbackCapture.WaveFormat)
                If pcm Is Nothing OrElse pcm.Length <= 0 Then
                    Return
                End If

                EnqueueSystemPcm(pcm, pcm.Length)

                If SourceMode = AudioSourceMode.SystemAudioOnly Then
                    Dim frames As List(Of Byte()) = DequeueSystemFrames(TARGET_FRAME_BYTES)

                    For Each frame As Byte() In frames
                        EmitFrame(Nothing, 0, frame, frame.Length)
                    Next
                End If
            Catch ex As Exception
                RaiseEvent CaptureError(Me, New TranscriptionErrorEventArgs("OnLoopbackData: " & ex.Message, ex, False))
            End Try
        End Sub

        Private Sub EnqueueSystemPcm(buffer As Byte(), bytesValid As Integer)
            If buffer Is Nothing OrElse bytesValid <= 0 Then
                Return
            End If

            SyncLock _systemPcmSyncRoot
                For i As Integer = 0 To bytesValid - 1
                    _systemPcmQueue.Enqueue(buffer(i))
                Next
            End SyncLock
        End Sub

        Private Function DequeueSystemPcm(maxBytes As Integer) As Byte()
            If maxBytes <= 0 Then
                Return Nothing
            End If

            SyncLock _systemPcmSyncRoot
                Dim bytesToRead As Integer = Math.Min(maxBytes, _systemPcmQueue.Count)
                bytesToRead -= (bytesToRead Mod 2)

                If bytesToRead <= 0 Then
                    Return Nothing
                End If

                Dim result(bytesToRead - 1) As Byte

                For i As Integer = 0 To bytesToRead - 1
                    result(i) = _systemPcmQueue.Dequeue()
                Next

                Return result
            End SyncLock
        End Function

        Private Function DequeueSystemFrames(frameBytes As Integer) As List(Of Byte())
            Dim frames As New List(Of Byte())()

            If frameBytes <= 0 Then
                Return frames
            End If

            SyncLock _systemPcmSyncRoot
                While _systemPcmQueue.Count >= frameBytes
                    Dim frame(frameBytes - 1) As Byte

                    For i As Integer = 0 To frameBytes - 1
                        frame(i) = _systemPcmQueue.Dequeue()
                    Next

                    frames.Add(frame)
                End While
            End SyncLock

            Return frames
        End Function

        ''' <summary>Mix or interleave mic+system into the output frame and raise.</summary>
        Private Sub EmitFrame(mic As Byte(), micBytes As Integer, sys As Byte(), sysBytes As Integer)
            Dim outBuf As Byte()
            Dim outBytes As Integer

            If MultiChannelStereo AndAlso SourceMode = AudioSourceMode.MicrophoneAndSystem Then
                Dim samples As Integer = Math.Max(micBytes, sysBytes) \ 2
                If samples <= 0 Then
                    Return
                End If

                outBuf = New Byte(samples * 4 - 1) {}
                outBytes = samples * 4

                For i As Integer = 0 To samples - 1
                    Dim m As Short = If(mic IsNot Nothing AndAlso i * 2 + 1 < micBytes, BitConverter.ToInt16(mic, i * 2), CShort(0))
                    Dim s As Short = If(sys IsNot Nothing AndAlso i * 2 + 1 < sysBytes, BitConverter.ToInt16(sys, i * 2), CShort(0))

                    Buffer.BlockCopy(BitConverter.GetBytes(m), 0, outBuf, i * 4, 2)
                    Buffer.BlockCopy(BitConverter.GetBytes(s), 0, outBuf, i * 4 + 2, 2)
                Next
            Else
                Dim primary As Byte() = If(mic, sys)
                Dim primaryBytes As Integer = If(mic IsNot Nothing, micBytes, sysBytes)

                If primary Is Nothing OrElse primaryBytes <= 0 Then
                    Return
                End If

                outBuf = New Byte(primaryBytes - 1) {}
                outBytes = primaryBytes

                If mic Is Nothing AndAlso sys IsNot Nothing Then
                    Buffer.BlockCopy(sys, 0, outBuf, 0, primaryBytes)
                Else
                    Dim peak As Integer = 1

                    For i As Integer = 0 To primaryBytes - 2 Step 2
                        Dim a As Integer = If(mic IsNot Nothing AndAlso i + 1 < micBytes, BitConverter.ToInt16(mic, i), 0)
                        Dim b As Integer = If(sys IsNot Nothing AndAlso i + 1 < sysBytes, BitConverter.ToInt16(sys, i), 0)
                        Dim sum As Integer = CInt((a + b) * _agcGain)

                        If sum > Short.MaxValue Then
                            sum = Short.MaxValue
                        End If

                        If sum < Short.MinValue Then
                            sum = Short.MinValue
                        End If

                        If Math.Abs(sum) > peak Then
                            peak = Math.Abs(sum)
                        End If

                        Dim bb As Byte() = BitConverter.GetBytes(CShort(sum))
                        outBuf(i) = bb(0)
                        outBuf(i + 1) = bb(1)
                    Next

                    Dim desired As Single = 22000.0F / peak

                    If desired < 1.0F Then
                        _agcGain = _agcGain * 0.9F + desired * 0.1F
                    Else
                        _agcGain = _agcGain * 0.995F + desired * 0.005F
                    End If

                    If _agcGain < 0.25F Then
                        _agcGain = 0.25F
                    End If

                    If _agcGain > 4.0F Then
                        _agcGain = 4.0F
                    End If
                End If
            End If

            If _debugWriter IsNot Nothing Then
                SyncLock _debugSyncRoot
                    Try
                        _debugWriter.Write(outBuf, 0, outBytes)
                    Catch
                    End Try
                End SyncLock
            End If

            RaiseEvent Frame(Me, New FrameEventArgs(outBuf, outBytes))
        End Sub

        Private Shared Function ConvertToTargetPcm(source As Byte(), bytesValid As Integer, format As WaveFormat) As Byte()
            If source Is Nothing OrElse format Is Nothing OrElse bytesValid <= 0 Then
                Return New Byte() {}
            End If

            Dim blockAlign As Integer = Math.Max(1, format.BlockAlign)
            Dim frameCount As Integer = bytesValid \ blockAlign
            Dim channelCount As Integer = Math.Max(1, format.Channels)

            If frameCount <= 0 OrElse format.SampleRate <= 0 Then
                Return New Byte() {}
            End If

            Dim mono(frameCount - 1) As Single

            For frameIndex As Integer = 0 To frameCount - 1
                Dim frameOffset As Integer = frameIndex * blockAlign
                Dim sum As Single = 0.0F

                For channelIndex As Integer = 0 To channelCount - 1
                    Dim sampleOffset As Integer = frameOffset + channelIndex * Math.Max(1, format.BitsPerSample \ 8)
                    sum += ReadSampleAsSingle(source, sampleOffset, format)
                Next

                mono(frameIndex) = sum / channelCount
            Next

            Dim resampled As Single()

            If format.SampleRate = TARGET_SAMPLE_RATE Then
                resampled = mono
            Else
                Dim outputSampleCount As Integer =
                    CInt(Math.Floor(frameCount * (CDbl(TARGET_SAMPLE_RATE) / CDbl(format.SampleRate))))

                If outputSampleCount <= 0 Then
                    Return New Byte() {}
                End If

                ReDim resampled(outputSampleCount - 1)

                For i As Integer = 0 To outputSampleCount - 1
                    Dim sourcePosition As Double = CDbl(i) * CDbl(format.SampleRate) / CDbl(TARGET_SAMPLE_RATE)
                    Dim sourceIndex As Integer = CInt(Math.Floor(sourcePosition))
                    Dim fraction As Double = sourcePosition - CDbl(sourceIndex)

                    If sourceIndex >= mono.Length - 1 Then
                        resampled(i) = mono(mono.Length - 1)
                    Else
                        Dim a As Double = mono(sourceIndex)
                        Dim b As Double = mono(sourceIndex + 1)
                        resampled(i) = CSng(a + (b - a) * fraction)
                    End If
                Next
            End If

            Dim outputBytes(resampled.Length * 2 - 1) As Byte

            For i As Integer = 0 To resampled.Length - 1
                Dim sampleValue As Single = resampled(i)

                If sampleValue > 1.0F Then
                    sampleValue = 1.0F
                ElseIf sampleValue < -1.0F Then
                    sampleValue = -1.0F
                End If

                Dim pcmValue As Integer = CInt(Math.Round(sampleValue * 32767.0F))

                If pcmValue > Short.MaxValue Then
                    pcmValue = Short.MaxValue
                ElseIf pcmValue < Short.MinValue Then
                    pcmValue = Short.MinValue
                End If

                Dim bytes As Byte() = BitConverter.GetBytes(CShort(pcmValue))
                outputBytes(i * 2) = bytes(0)
                outputBytes(i * 2 + 1) = bytes(1)
            Next

            Return outputBytes
        End Function

        Private Shared Function ReadSampleAsSingle(source As Byte(), offset As Integer, format As WaveFormat) As Single
            If source Is Nothing OrElse format Is Nothing OrElse offset < 0 OrElse offset >= source.Length Then
                Return 0.0F
            End If

            If IsFloatEncoding(format) Then
                If format.BitsPerSample = 32 AndAlso offset + 3 < source.Length Then
                    Return ClampToUnit(BitConverter.ToSingle(source, offset))
                End If

                Return 0.0F
            End If

            If IsPcmEncoding(format) Then
                Select Case format.BitsPerSample
                    Case 8
                        Return (CSng(source(offset)) - 128.0F) / 128.0F

                    Case 16
                        If offset + 1 < source.Length Then
                            Return CSng(BitConverter.ToInt16(source, offset) / 32768.0R)
                        End If

                    Case 24
                        If offset + 2 < source.Length Then
                            Dim value As Integer =
                                source(offset) Or
                                (source(offset + 1) << 8) Or
                                (source(offset + 2) << 16)

                            If (value And &H800000) <> 0 Then
                                value = value Or &HFF000000
                            End If

                            Return CSng(value / 8388608.0R)
                        End If

                    Case 32
                        If offset + 3 < source.Length Then
                            Return CSng(BitConverter.ToInt32(source, offset) / 2147483648.0R)
                        End If
                End Select
            End If

            Return 0.0F
        End Function

        Private Shared Function IsPcmEncoding(format As WaveFormat) As Boolean
            If format Is Nothing Then
                Return False
            End If

            If format.Encoding = WaveFormatEncoding.Pcm Then
                Return True
            End If

            Dim extensible As WaveFormatExtensible = TryCast(format, WaveFormatExtensible)
            If extensible IsNot Nothing Then
                Return extensible.SubFormat = PcmSubFormat
            End If

            Return False
        End Function

        Private Shared Function IsFloatEncoding(format As WaveFormat) As Boolean
            If format Is Nothing Then
                Return False
            End If

            If format.Encoding = WaveFormatEncoding.IeeeFloat Then
                Return True
            End If

            Dim extensible As WaveFormatExtensible = TryCast(format, WaveFormatExtensible)
            If extensible IsNot Nothing Then
                Return extensible.SubFormat = FloatSubFormat
            End If

            Return False
        End Function

        Private Shared Function ClampToUnit(value As Single) As Single
            If value > 1.0F Then
                Return 1.0F
            End If

            If value < -1.0F Then
                Return -1.0F
            End If

            Return value
        End Function

        Public Sub [Stop]()
            If Not _running Then
                Return
            End If

            _running = False

            Try
                If _waveIn IsNot Nothing Then
                    RemoveHandler _waveIn.DataAvailable, AddressOf OnMicData
                    _waveIn.StopRecording()
                    _waveIn.Dispose()
                    _waveIn = Nothing
                End If
            Catch
            End Try

            Try
                If _loopbackCapture IsNot Nothing Then
                    RemoveHandler _loopbackCapture.DataAvailable, AddressOf OnLoopbackData
                    _loopbackCapture.StopRecording()
                    _loopbackCapture.Dispose()
                    _loopbackCapture = Nothing
                End If
            Catch
            End Try

            SyncLock _systemPcmSyncRoot
                _systemPcmQueue.Clear()
            End SyncLock

            Try
                If _debugWriter IsNot Nothing Then
                    SyncLock _debugSyncRoot
                        _debugWriter.Dispose()
                        _debugWriter = Nothing
                    End SyncLock
                End If
            Catch
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            [Stop]()
        End Sub

    End Class

End Namespace