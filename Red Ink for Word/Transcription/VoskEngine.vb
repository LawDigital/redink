Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq
Imports Vosk

Namespace Transcription

    Public Class VoskEngine
        Implements ITranscriptionEngine

        Public Event PartialResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.PartialResult
        Public Event FinalResult As EventHandler(Of TranscriptionEventArgs) Implements ITranscriptionEngine.FinalResult
        Public Event EngineError As EventHandler(Of TranscriptionErrorEventArgs) Implements ITranscriptionEngine.EngineError
        Public Event Status As EventHandler(Of TranscriptionStatusEventArgs) Implements ITranscriptionEngine.Status

        Public ReadOnly Property Name As String Implements ITranscriptionEngine.Name
            Get
                Return _modelName
            End Get
        End Property

        Public ReadOnly Property Kind As EngineKind Implements ITranscriptionEngine.Kind
            Get
                Return EngineKind.Vosk
            End Get
        End Property

        Public ReadOnly Property SupportsLiveStreaming As Boolean Implements ITranscriptionEngine.SupportsLiveStreaming
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsFileTranscription As Boolean Implements ITranscriptionEngine.SupportsFileTranscription
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsDiarization As Boolean Implements ITranscriptionEngine.SupportsDiarization
            Get
                Return True
            End Get
        End Property

        Public ReadOnly Property SupportsMultiChannelDiarization As Boolean Implements ITranscriptionEngine.SupportsMultiChannelDiarization
            Get
                Return False
            End Get
        End Property

        Private ReadOnly _modelRoot As String
        Private ReadOnly _modelName As String
        Private _model As Model
        Private _spkModel As SpkModel
        Private _rec As VoskRecognizer

        Private _knownSpeakers As New Dictionary(Of String, List(Of List(Of Double)))()
        Private _similarityThreshold As Double = 1.0

        Public Sub New(modelRoot As String, modelDirName As String)
            _modelRoot = modelRoot
            _modelName = modelDirName
        End Sub

        Private Sub Init(opts As TranscriptionOptions)
            Dim modelPath As String = Path.Combine(_modelRoot, _modelName)

            _model = New Model(modelPath)
            _rec = New VoskRecognizer(_model, 16000.0F)

            If opts.EnableDiarization Then
                Dim spkRoot As String = Path.Combine(_modelRoot, "Speaker")
                Dim spk As String = Nothing

                If Directory.Exists(spkRoot) Then
                    spk = Directory.GetDirectories(spkRoot, "vosk-model*").FirstOrDefault()
                End If

                If String.IsNullOrEmpty(spk) Then
                    RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("No speaker model found in " & spkRoot, Nothing, False))
                Else
                    _spkModel = New SpkModel(spk)
                    _rec.SetSpkModel(_spkModel)
                End If
            End If

            _rec.SetMaxAlternatives(0)
            _rec.SetWords(True)
            _rec.SetPartialWords(True)

            _similarityThreshold = opts.VoskSimilarityThreshold
            If _similarityThreshold < 0.2 Then _similarityThreshold = 0.2
            If _similarityThreshold > 2.5 Then _similarityThreshold = 2.5
        End Sub

        Public Function StartLiveAsync(opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.StartLiveAsync
            Init(opts)
            Return Task.CompletedTask
        End Function

        Public Function PushAudioAsync(pcm As Byte(), bytesValid As Integer, ct As CancellationToken) As Task Implements ITranscriptionEngine.PushAudioAsync
            If _rec Is Nothing Then
                Return Task.CompletedTask
            End If

            Try
                If _rec.AcceptWaveform(pcm, bytesValid) Then
                    ProcessJson(_rec.Result(), True)
                Else
                    ProcessJson(_rec.PartialResult(), False)
                End If
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Vosk: " & ex.Message, ex, False))
            End Try

            Return Task.CompletedTask
        End Function

        Public Function StopLiveAsync() As Task Implements ITranscriptionEngine.StopLiveAsync
            Try
                If _rec IsNot Nothing Then
                    ProcessJson(_rec.FinalResult(), True)
                End If
            Catch
            End Try

            Return Task.CompletedTask
        End Function

        Public Async Function TranscribeFileAsync(filePath As String, opts As TranscriptionOptions, ct As CancellationToken) As Task Implements ITranscriptionEngine.TranscribeFileAsync
            Init(opts)

            Dim pcm As Byte() = LoadAudioToPcm16Mono16k(filePath)
            Dim offset As Integer = 0
            Const chunkSize As Integer = 4096

            While offset < pcm.Length AndAlso Not ct.IsCancellationRequested
                Dim len As Integer = Math.Min(chunkSize, pcm.Length - offset)
                Dim slice(len - 1) As Byte
                Buffer.BlockCopy(pcm, offset, slice, 0, len)
                Await PushAudioAsync(slice, len, ct)
                offset += len
            End While

            Try
                ProcessJson(_rec.FinalResult(), True)
            Catch
            End Try
        End Function

        Private Sub ProcessJson(json As String, isFinal As Boolean)
            If String.IsNullOrEmpty(json) Then
                Return
            End If

            Try
                Dim jo As JObject = JObject.Parse(json)

                If isFinal Then
                    Dim t As String = If(jo("text")?.ToString(), String.Empty)
                    If String.IsNullOrWhiteSpace(t) Then
                        Return
                    End If

                    Dim speaker As String = String.Empty
                    If jo("spk") IsNot Nothing AndAlso jo("spk").Type = JTokenType.Array Then
                        Dim emb As List(Of Double) = CType(jo("spk"), JArray).Select(Function(x) CDbl(x)).ToList()
                        speaker = IdentifySpeaker(emb)
                    End If

                    RaiseEvent FinalResult(Me, New TranscriptionEventArgs(t, True, speaker))
                Else
                    Dim p As String = If(jo("partial")?.ToString(), String.Empty)
                    If Not String.IsNullOrWhiteSpace(p) Then
                        RaiseEvent PartialResult(Me, New TranscriptionEventArgs(p, False))
                    End If
                End If
            Catch ex As Exception
                RaiseEvent EngineError(Me, New TranscriptionErrorEventArgs("Vosk JSON: " & ex.Message, ex, False))
            End Try
        End Sub

        Private Function IdentifySpeaker(newEmb As List(Of Double)) As String
            newEmb = Normalize(newEmb)

            Dim best As String = "Unknown"
            Dim bestDist As Double = Double.MaxValue

            For Each kvp In _knownSpeakers
                Dim avg As List(Of Double) = Average(kvp.Value)
                Dim d As Double = EuclideanDistance(avg, newEmb)

                If d < bestDist AndAlso d < _similarityThreshold Then
                    best = kvp.Key
                    bestDist = d
                End If
            Next

            If best = "Unknown" Then
                Dim id As String = "Speaker " & (_knownSpeakers.Count + 1).ToString()
                _knownSpeakers(id) = New List(Of List(Of Double)) From {newEmb}
                Return id
            Else
                _knownSpeakers(best).Add(newEmb)
                If _knownSpeakers(best).Count > 5 Then
                    _knownSpeakers(best).RemoveAt(0)
                End If
                Return best
            End If
        End Function

        Private Shared Function Normalize(v As List(Of Double)) As List(Of Double)
            Dim n As Double = Math.Sqrt(v.Sum(Function(x) x * x))
            If n = 0 Then
                Return v
            End If
            Return v.Select(Function(x) x / n).ToList()
        End Function

        Private Shared Function Average(list As List(Of List(Of Double))) As List(Of Double)
            Dim size As Integer = list(0).Count
            Dim acc As New List(Of Double)(New Double(size - 1) {})

            For Each e In list
                For i As Integer = 0 To size - 1
                    acc(i) += e(i)
                Next
            Next

            For i As Integer = 0 To size - 1
                acc(i) /= list.Count
            Next

            Return acc
        End Function

        Private Shared Function EuclideanDistance(a As List(Of Double), b As List(Of Double)) As Double
            Dim s As Double = 0
            For i As Integer = 0 To a.Count - 1
                s += (a(i) - b(i)) ^ 2
            Next
            Return Math.Sqrt(s)
        End Function

        Friend Shared Function LoadAudioToPcm16Mono16k(filePath As String) As Byte()
            Using r As New MediaFoundationReader(filePath)
                Dim fmt As New WaveFormat(16000, 16, 1)

                Using rs As New MediaFoundationResampler(r, fmt)
                    rs.ResamplerQuality = 60

                    Using ms As New MemoryStream()
                        Using w As New WaveFileWriter(ms, fmt)
                            Dim buf(4095) As Byte
                            Dim n As Integer

                            Do
                                n = rs.Read(buf, 0, buf.Length)
                                If n > 0 Then
                                    w.Write(buf, 0, n)
                                End If
                            Loop While n > 0

                            w.Flush()
                        End Using

                        Dim all As Byte() = ms.ToArray()
                        If all.Length > 44 Then
                            Dim raw(all.Length - 45) As Byte
                            Buffer.BlockCopy(all, 44, raw, 0, all.Length - 44)
                            Return raw
                        End If

                        Return all
                    End Using
                End Using
            End Using
        End Function

        Public Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
            Try
                If _rec IsNot Nothing Then
                    _rec.Dispose()
                End If
            Catch
            End Try
            _rec = Nothing

            Try
                If _model IsNot Nothing Then
                    _model.Dispose()
                End If
            Catch
            End Try
            _model = Nothing

            Return New ValueTask()
        End Function
    End Class

End Namespace