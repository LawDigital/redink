' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Opt-in production-code tests. Invoke from a Windows debug session after deployment:
' SharedLibrary.Agents.TextFileSnapshotSelfTests.RunAll(<allowed writable test root>)
' Creates/deletes only one unique child directory. No OCR, no LLM, no UI is needed.
Option Strict On
Option Explicit On
Option Infer On

Namespace Agents
    Public NotInheritable Class TextFileSnapshotSelfTests
        Private Sub New()
        End Sub

        Public Shared Function RunAll(testRoot As System.String) As System.String
            Dim root As System.String = PathPolicy.Resolve(System.IO.Path.Combine(testRoot,
                "RI-P1-Tests-" & System.Guid.NewGuid().ToString("N")), PathAccess.Write)
            System.IO.Directory.CreateDirectory(root)
            Dim log As New System.Text.StringBuilder()
            Dim passed As System.Int32 = 0
            Dim failed As System.Int32 = 0
            Try
                Dim source As System.String = System.IO.Path.Combine(root, "source.md")
                Dim content As System.String = "# Test" & System.Environment.NewLine & "Grüsse - vollständig." & System.Environment.NewLine
                Dim utf8 As New System.Text.UTF8Encoding(True, True)
                System.IO.File.WriteAllText(source, content, utf8)

                RunCase("strict read preserves content and raw-byte SHA-256", log, passed, failed,
                    Sub()
                        Dim snapshot As TextFileSnapshot = TextFileSnapshot.Read(source, True)
                        Require(snapshot.Content = content, "Content changed")
                        Require(snapshot.SizeBytes = New System.IO.FileInfo(source).Length, "Byte size mismatch")
                        Require(snapshot.Sha256 = TextFileSnapshot.ComputeHash(System.IO.File.ReadAllBytes(source)), "Hash mismatch")
                    End Sub)

                For Each encoding As System.Text.Encoding In New System.Text.Encoding() {
                    New System.Text.UnicodeEncoding(False, True, True), New System.Text.UnicodeEncoding(True, True, True),
                    New System.Text.UTF32Encoding(False, True, True), New System.Text.UTF32Encoding(True, True, True)}
                    Dim selectedEncoding As System.Text.Encoding = encoding
                    RunCase("BOM decoding " & selectedEncoding.WebName, log, passed, failed,
                        Sub()
                            Dim p As System.String = System.IO.Path.Combine(root, selectedEncoding.WebName & ".txt")
                            System.IO.File.WriteAllText(p, content, selectedEncoding)
                            Require(TextFileSnapshot.Read(p, True).Content = content, "BOM decoding changed text")
                        End Sub)
                Next

                RunCase("invalid UTF-8 rejected by document file input", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "invalid.txt")
                        System.IO.File.WriteAllBytes(p, New System.Byte() {&HC3, &H28})
                        ExpectInputError("invalid_text_encoding", Sub() TextFileSnapshot.Read(p, True))
                    End Sub)

                RunCase("expected hash rejects changed source", log, passed, failed,
                    Sub() ExpectInputError("source_hash_mismatch", Sub() TextFileSnapshot.Read(source, True, New System.String("0"c, 64))))
                RunCase("malformed expected hash rejected", log, passed, failed,
                    Sub() ExpectInputError("invalid_source_hash", Sub() TextFileSnapshot.Read(source, True, "bad")))
                RunCase("missing source rejected", log, passed, failed,
                    Sub() ExpectInputError("not_found", Sub() TextFileSnapshot.Read(System.IO.Path.Combine(root, "missing.txt"), True)))

                RunCase("legacy prefix, byte size and nonpositive cap preserved", log, passed, failed,
                    Sub()
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                        args("max_chars") = 4
                        Dim r As Newtonsoft.Json.Linq.JObject = ReadResult(args)
                        Require(r.Value(Of System.String)("text") = content.Substring(0, 4), "Legacy prefix changed")
                        Require(r.Value(Of System.Int64)("size") = New System.IO.FileInfo(source).Length, "size is no longer bytes")
                        Require(r.Value(Of System.Boolean)("truncated"), "Legacy truncated missing")
                        args("max_chars") = -1
                        Require(ReadResult(args).Value(Of System.String)("text") = content, "Legacy negative cap changed")
                    End Sub)

                RunCase("paging reassembles exact source", log, passed, failed,
                    Sub()
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                        args("start_char") = 0 : args("max_chars") = 3
                        Dim joined As New System.Text.StringBuilder()
                        Do
                            Dim r As Newtonsoft.Json.Linq.JObject = ReadResult(args)
                            Require(r("error") Is Nothing, "Window failed")
                            joined.Append(r.Value(Of System.String)("text"))
                            args("expected_snapshot_sha256") = r.Value(Of System.String)("snapshot_sha256")
                            If Not r.Value(Of System.Boolean)("has_more") Then
                                Require(r("next_offset").Type = Newtonsoft.Json.Linq.JTokenType.Null, "EOF next_offset is not null")
                                Exit Do
                            End If
                            args("start_char") = r.Value(Of System.Int32)("next_offset")
                        Loop
                        Require(joined.ToString() = content, "Paged content differs")
                    End Sub)

                RunCase("offset alias and numeric string normalized", log, passed, failed,
                    Sub()
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                        args("offset") = "2" : args("max_chars") = 4
                        Require(ReadResult(args).Value(Of System.String)("text") = content.Substring(2, 4), "Alias failed")
                    End Sub)
                RunCase("conflicting offsets rejected", log, passed, failed,
                    Sub()
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                        args("offset") = 2 : args("start_char") = 1
                        Require(ReadResult(args).Value(Of System.String)("error") = "conflicting_offsets", "Conflict ignored")
                    End Sub)
                For Each bad As System.Object In New System.Object() {-1, 1.5D, True, "abc", Nothing}
                    Dim badOffset As System.Object = bad
                    RunCase("invalid offset " & System.Convert.ToString(badOffset, System.Globalization.CultureInfo.InvariantCulture), log, passed, failed,
                        Sub()
                            Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                            args("start_char") = badOffset
                            Require(ReadResult(args).Value(Of System.String)("error") = "invalid_offset", "Invalid offset accepted")
                        End Sub)
                Next
                RunCase("EOF and beyond EOF distinguished", log, passed, failed,
                    Sub()
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(source)
                        args("start_char") = content.Length
                        Require(ReadResult(args).Value(Of System.String)("text") = System.String.Empty, "EOF failed")
                        args("start_char") = content.Length + 1
                        Require(ReadResult(args).Value(Of System.String)("error") = "offset_out_of_range", "Past EOF accepted")
                    End Sub)
                RunCase("explicit windows preserve surrogate pairs", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "unicode.txt")
                        Dim emoji As System.String = System.Char.ConvertFromUtf32(&H1F600)
                        System.IO.File.WriteAllText(p, "A" & emoji & "B", utf8)
                        Dim args As System.Collections.Generic.Dictionary(Of System.String, System.Object) = ArgsFor(p)
                        args("start_char") = 1 : args("max_chars") = 1
                        Dim r As Newtonsoft.Json.Linq.JObject = ReadResult(args)
                        Require(r.Value(Of System.String)("text") = emoji AndAlso r.Value(Of System.Int32)("next_offset") = 3, "Surrogate split")
                        args("start_char") = 2
                        Require(ReadResult(args).Value(Of System.String)("error") = "invalid_offset", "Low-surrogate offset accepted")
                    End Sub)

                RunCase("configured text size limit enforced before allocation", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "too-large.txt")
                        Using stream As New System.IO.FileStream(p, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write)
                            stream.SetLength(CLng(PathPolicy.MaxFileSizeBytes) + 1L)
                        End Using
                        ExpectInputError("file_too_large", Sub() TextFileSnapshot.Read(p, True))
                    End Sub)

                RunCase("atomic no-overwrite and replacement", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "atomic.txt")
                        TextFileSnapshot.WriteUtf8Atomic(p, "first", False)
                        Dim refused As System.Boolean = False
                        Try
                            TextFileSnapshot.WriteUtf8Atomic(p, "second", False)
                        Catch ex As System.IO.IOException
                            refused = True
                        End Try
                        Require(refused AndAlso TextFileSnapshot.Read(p).Content = "first", "No-overwrite violated")
                        TextFileSnapshot.WriteUtf8Atomic(p, "second", True)
                        Require(TextFileSnapshot.Read(p).Content = "second", "Replacement failed")
                        Require(System.IO.Directory.GetFiles(root, ".ri-text-*.tmp").Length = 0, "Temporary file leaked")
                    End Sub)

                RunCase("extraction resource reuses identical source and invalidates changed source/config", log, passed, failed,
                    Sub()
                        TextExtractionResourceRegistry.ClearForTests()
                        Dim p As System.String = System.IO.Path.Combine(root, "resource-source.txt")
                        System.IO.File.WriteAllText(p, "alpha", utf8)
                        Dim count As System.Int32 = 0
                        Dim extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)) =
                            Function(snapshotPath As System.String, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionPayload)
                                System.Threading.Interlocked.Increment(count)
                                Return System.Threading.Tasks.Task.FromResult(New TextExtractionPayload With {
                                    .Success = True,
                                    .Content = System.IO.File.ReadAllText(snapshotPath),
                                    .ContentFormat = "plain_text"
                                })
                            End Function
                        Dim first As TextExtractionResource = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg-a", "opt-a", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Dim second As TextExtractionResource = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg-a", "opt-a", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Require(count = 1, "Identical resource was extracted twice")
                        Require(first.ResourceId = second.ResourceId AndAlso second.ReuseStatus = "session_cache_hit", "Resource reuse metadata wrong")

                        System.IO.File.WriteAllText(p, "beta", utf8)
                        Dim changedSource As TextExtractionResource = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg-a", "opt-a", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Require(count = 2 AndAlso changedSource.ResourceId <> first.ResourceId, "Changed source reused stale resource")

                        Dim changedConfig As TextExtractionResource = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg-b", "opt-a", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Require(count = 3 AndAlso changedConfig.ResourceId <> changedSource.ResourceId, "Changed configuration reused stale resource")
                    End Sub)

                RunCase("failed extraction resource is never cached", log, passed, failed,
                    Sub()
                        TextExtractionResourceRegistry.ClearForTests()
                        Dim p As System.String = System.IO.Path.Combine(root, "resource-failure.txt")
                        System.IO.File.WriteAllText(p, "failure-source", utf8)
                        Dim count As System.Int32 = 0
                        Dim extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)) =
                            Function(snapshotPath As System.String, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionPayload)
                                System.Threading.Interlocked.Increment(count)
                                Return System.Threading.Tasks.Task.FromResult(New TextExtractionPayload With {
                                    .Success = False,
                                    .ErrorCode = "expected_failure",
                                    .Message = "test"
                                })
                            End Function
                        TextExtractionResourceRegistry.ResolveAsync(p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        TextExtractionResourceRegistry.ResolveAsync(p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Require(count = 2, "Failed resource was cached")
                    End Sub)

                RunCase("cancelled extraction resource is never cached", log, passed, failed,
                    Sub()
                        TextExtractionResourceRegistry.ClearForTests()
                        Dim p As System.String = System.IO.Path.Combine(root, "resource-cancel.txt")
                        System.IO.File.WriteAllText(p, "cancel-source", utf8)
                        Dim count As System.Int32 = 0
                        Dim extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)) =
                            Function(snapshotPath As System.String, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionPayload)
                                System.Threading.Interlocked.Increment(count)
                                If count = 1 Then
                                    Return System.Threading.Tasks.Task.FromException(Of TextExtractionPayload)(New System.OperationCanceledException("test cancellation"))
                                End If
                                Return System.Threading.Tasks.Task.FromResult(New TextExtractionPayload With {.Success = True, .Content = "ok", .ContentFormat = "plain_text"})
                            End Function
                        Dim cancelled As System.Boolean = False
                        Try
                            TextExtractionResourceRegistry.ResolveAsync(p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Catch ex As System.OperationCanceledException
                            cancelled = True
                        End Try
                        Require(cancelled, "Cancellation was swallowed")
                        Dim second As TextExtractionResource = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None).GetAwaiter().GetResult()
                        Require(count = 2 AndAlso second IsNot Nothing AndAlso second.Payload.Success, "Cancelled resource was cached or blocked retry")
                    End Sub)

                RunCase("concurrent identical extraction uses single flight", log, passed, failed,
                    Sub()
                        TextExtractionResourceRegistry.ClearForTests()
                        Dim p As System.String = System.IO.Path.Combine(root, "resource-flight.txt")
                        System.IO.File.WriteAllText(p, "single-flight", utf8)
                        Dim count As System.Int32 = 0
                        Dim extractor As System.Func(Of System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of TextExtractionPayload)) =
                            Async Function(snapshotPath As System.String, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of TextExtractionPayload)
                                System.Threading.Interlocked.Increment(count)
                                Await System.Threading.Tasks.Task.Delay(150, token).ConfigureAwait(False)
                                Return New TextExtractionPayload With {.Success = True, .Content = System.IO.File.ReadAllText(snapshotPath), .ContentFormat = "plain_text"}
                            End Function
                        Dim firstTask As System.Threading.Tasks.Task(Of TextExtractionResource) = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None)
                        Dim secondTask As System.Threading.Tasks.Task(Of TextExtractionResource) = TextExtractionResourceRegistry.ResolveAsync(
                            p, root, "test", "v1", "cfg", "opt", extractor, System.Threading.CancellationToken.None)
                        System.Threading.Tasks.Task.WhenAll(firstTask, secondTask).GetAwaiter().GetResult()
                        Require(count = 1, "Single-flight executed extractor more than once")
                        Require(firstTask.Result.ResourceId = secondTask.Result.ResourceId, "Single-flight returned different resources")
                    End Sub)

                RunCase("text export and reuse preserve output hash and timestamp", log, passed, failed,
                    Sub()
                        Dim outputDirectory As System.String = System.IO.Path.Combine(root, "exports")
                        Dim args As New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {
                            {"input_path", source}, {"output_directory", outputDirectory}, {"overwrite", False}}
                        Dim first As Newtonsoft.Json.Linq.JObject = ExportResult(args)
                        Dim item As Newtonsoft.Json.Linq.JToken = first("items")(0)
                        Require(item.Value(Of System.String)("status") = "converted", "Export failed: " & first.ToString())
                        Require(item.Value(Of System.Int32)("char_count") = content.Length, "Export char_count wrong")
                        Dim p As System.String = item.Value(Of System.String)("output_path")
                        Dim stamp As System.DateTime = System.IO.File.GetLastWriteTimeUtc(p)
                        Dim second As Newtonsoft.Json.Linq.JToken = ExportResult(args)("items")(0)
                        Require(second.Value(Of System.String)("status") = "skipped_existing", "Not skipped")
                        Require(second.Value(Of System.String)("snapshot_sha256") = item.Value(Of System.String)("snapshot_sha256"), "Snapshot changed on reuse")
                        Require(System.IO.File.GetLastWriteTimeUtc(p) = stamp, "Reuse rewrote file")
                        Require(second.Value(Of System.Boolean)("source_association_verified"), "Known session output was not provenance-verified")
                        Require(second.Value(Of System.String)("reuse_validation") = "session_resource_verified", "Verified reuse status missing")
                        Require(second.Value(Of System.String)("resource_id") = item.Value(Of System.String)("resource_id"), "Resource identity changed")
                    End Sub)

                RunCase("overwrite reuses identical extraction resource without re-extraction", log, passed, failed,
                    Sub()
                        Dim outputDirectory As System.String = System.IO.Path.Combine(root, "resource-reuse")
                        Dim args As New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {
                            {"input_path", source}, {"output_directory", outputDirectory}, {"overwrite", True}}
                        Dim first As Newtonsoft.Json.Linq.JToken = ExportResult(args)("items")(0)
                        Dim second As Newtonsoft.Json.Linq.JToken = ExportResult(args)("items")(0)
                        Require(first.Value(Of System.String)("resource_id") = second.Value(Of System.String)("resource_id"), "Resource id changed")
                        Require(second.Value(Of System.String)("resource_reuse") = "session_cache_hit", "Session cache was not reused")
                        Require(first.Value(Of System.String)("snapshot_sha256") = second.Value(Of System.String)("snapshot_sha256"), "Reused snapshot changed")
                    End Sub)

                RunCase("single-file existing output bypasses even invalid PDF parsing", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "broken.pdf")
                        System.IO.File.WriteAllText(p, "not a PDF", utf8)
                        Dim outputDirectory As System.String = System.IO.Path.Combine(root, "skip")
                        System.IO.Directory.CreateDirectory(outputDirectory)
                        System.IO.File.WriteAllText(System.IO.Path.Combine(outputDirectory, "broken.pdf.txt"), "existing snapshot", utf8)
                        Dim args As New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {
                            {"input_path", p}, {"output_directory", outputDirectory}, {"overwrite", False}, {"ocr_pdf", True}}
                        Dim item As Newtonsoft.Json.Linq.JToken = ExportResult(args)("items")(0)
                        Require(item.Value(Of System.String)("status") = "skipped_existing", "Parser ran before skip")
                        Require(item("ocr_used").Type = Newtonsoft.Json.Linq.JTokenType.Null, "Unknown OCR metadata guessed")
                    End Sub)

                RunCase("invalid PDF failure cannot replace an existing output", log, passed, failed,
                    Sub()
                        Dim p As System.String = System.IO.Path.Combine(root, "invalid.pdf")
                        System.IO.File.WriteAllText(p, "not a PDF", utf8)
                        Dim outputDirectory As System.String = System.IO.Path.Combine(root, "fail-safe")
                        System.IO.Directory.CreateDirectory(outputDirectory)
                        Dim output As System.String = System.IO.Path.Combine(outputDirectory, "invalid.pdf.txt")
                        System.IO.File.WriteAllText(output, "preserve this", utf8)
                        Dim beforeHash As System.String = TextFileSnapshot.ComputeFileHash(output)
                        Dim args As New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {
                            {"input_path", p}, {"output_directory", outputDirectory}, {"overwrite", True}, {"ocr_pdf", False}}
                        Dim result As Newtonsoft.Json.Linq.JObject = ExportResult(args)
                        Require(result.Value(Of System.Int32)("failed_count") = 1, "PDF error reported as conversion success")
                        Require(result("items")(0).Value(Of System.String)("error") = "pdf_read_failed", "Structured PDF error missing")
                        Require(TextFileSnapshot.ComputeFileHash(output) = beforeHash, "Failed extraction overwrote old output")
                    End Sub)

                RunCase("directory existing output uses same pre-extraction gate", log, passed, failed,
                    Sub()
                        Dim inputDirectory As System.String = System.IO.Path.Combine(root, "batch")
                        Dim outputDirectory As System.String = System.IO.Path.Combine(root, "batch-out")
                        System.IO.Directory.CreateDirectory(inputDirectory)
                        System.IO.Directory.CreateDirectory(outputDirectory)
                        System.IO.File.WriteAllText(System.IO.Path.Combine(inputDirectory, "broken.pdf"), "not a PDF", utf8)
                        System.IO.File.WriteAllText(System.IO.Path.Combine(outputDirectory, "broken.pdf.txt"), "existing snapshot", utf8)
                        Dim args As New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {
                            {"input_path", inputDirectory}, {"output_directory", outputDirectory}, {"overwrite", False}}
                        Require(ExportResult(args)("items")(0).Value(Of System.String)("status") = "skipped_existing", "Directory skip failed")
                    End Sub)
            Finally
                Try
                    System.IO.Directory.Delete(root, recursive:=True)
                Catch ex As System.Exception
                    failed += 1
                    log.AppendLine("FAIL cleanup: " & ex.Message)
                End Try
            End Try
            Return "Passed=" & passed.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; Failed=" & failed.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                System.Environment.NewLine & log.ToString()
        End Function

        Private Shared Function ArgsFor(path As System.String) As System.Collections.Generic.Dictionary(Of System.String, System.Object)
            Return New System.Collections.Generic.Dictionary(Of System.String, System.Object) From {{"path", path}}
        End Function
        Private Shared Function ReadResult(args As System.Collections.Generic.Dictionary(Of System.String, System.Object)) As Newtonsoft.Json.Linq.JObject
            Return Newtonsoft.Json.Linq.JObject.Parse(TextTools.Execute(TextTools.ToolRead, args))
        End Function
        Private Shared Function ExportResult(args As System.Collections.Generic.Dictionary(Of System.String, System.Object)) As Newtonsoft.Json.Linq.JObject
            Return Newtonsoft.Json.Linq.JObject.Parse(TextTools.Execute(TextTools.ToolExportToText, args))
        End Function
        Private Shared Sub Require(condition As System.Boolean, message As System.String)
            If Not condition Then Throw New System.InvalidOperationException(message)
        End Sub
        Private Shared Sub ExpectInputError(code As System.String, action As System.Action)
            Try
                action()
            Catch ex As TextFileInputException
                Require(ex.ErrorCode = code, "Expected " & code & ", got " & ex.ErrorCode)
                Return
            End Try
            Throw New System.InvalidOperationException("Expected input error " & code)
        End Sub
        Private Shared Sub RunCase(name As System.String, log As System.Text.StringBuilder,
                                   ByRef passed As System.Int32, ByRef failed As System.Int32, action As System.Action)
            Try
                action()
                passed += 1
                log.AppendLine("PASS " & name)
            Catch ex As System.Exception
                failed += 1
                log.AppendLine("FAIL " & name & ": " & ex.Message)
            End Try
        End Sub
    End Class
End Namespace
