' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Per-sender recommendation outbox. Durable claims are authoritative across a
' crash or a partial job update. Only a successfully prepared Outlook message
' reserves advice; an uncertain submission is never submitted again automatically.

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn

    Private Const AP_PromptCoachingTopicCooldownDays As System.Int32 = 7
    Private Const AP_PromptCoachingPreflightRetryMinutes As System.Int32 = 5

    Private NotInheritable Class AutoPilotPromptCoachingDeliveryItem
        Public Property Job As Newtonsoft.Json.Linq.JObject
        Public Property Result As Global.SharedLibrary.Agents.PromptCoachingAdvice.Result
        Public Property Remaining As New System.Collections.Generic.List(Of Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation)()
    End Class

    Private Shared Function GetAutoPilotPromptCoachingReceiptPath(sender As System.String) As System.String
        Dim canonical As System.String = NormalizeAutoPilotPromptCoachingEmail(sender)
        If canonical.Length = 0 OrElse canonical <> sender Then Throw New System.IO.InvalidDataException("Invalid recommendation receipt owner.")
        Return GetAutoPilotPromptCoachingPath("receipts", Global.SharedLibrary.Agents.ToolingOriginalRequestLog.ComputeRequestSha256(sender), ".json")
    End Function

    Private Shared Function ReadAutoPilotPromptCoachingUtc(value As Newtonsoft.Json.Linq.JObject, name As System.String) As System.DateTime
        If value(name) Is Nothing OrElse value(name).Type <> Newtonsoft.Json.Linq.JTokenType.String Then
            Throw New System.IO.InvalidDataException("Missing recommendation receipt timestamp.")
        End If
        Dim result As System.DateTime
        If Not System.DateTime.TryParse(value.Value(Of System.String)(name), System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.RoundtripKind, result) OrElse result.Kind <> System.DateTimeKind.Utc Then
            Throw New System.IO.InvalidDataException("Invalid recommendation receipt timestamp.")
        End If
        Return result
    End Function

    Private Shared Function ReadAutoPilotPromptCoachingTopicSet(token As Newtonsoft.Json.Linq.JToken,
                                                               requireNonempty As System.Boolean,
                                                               maximum As System.Int32) As System.Collections.Generic.HashSet(Of System.String)
        Dim values As Newtonsoft.Json.Linq.JArray = TryCast(token, Newtonsoft.Json.Linq.JArray)
        If values Is Nothing OrElse values.Count > maximum OrElse (requireNonempty AndAlso values.Count = 0) Then
            Throw New System.IO.InvalidDataException("Invalid recommendation receipt topic set.")
        End If
        Dim result As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each value As Newtonsoft.Json.Linq.JToken In values
            If value.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Invalid recommendation receipt topic.")
            Dim topic As System.String = value.ToObject(Of System.String)()
            ' Submission receipts intentionally retain only canonical topic IDs.
            ' Their validation must not depend on the evolving display-advice schema.
            If Not Global.SharedLibrary.Agents.PromptCoachingAdvice.IsCanonicalTopic(topic) OrElse Not result.Add(topic) Then
                Throw New System.IO.InvalidDataException("Noncanonical or duplicate recommendation receipt topic.")
            End If
        Next
        Return result
    End Function

    Private Shared Function ReadAutoPilotPromptCoachingReceipt(sender As System.String) As Newtonsoft.Json.Linq.JObject
        Dim path As System.String = GetAutoPilotPromptCoachingReceiptPath(sender)
        Dim receipt As Newtonsoft.Json.Linq.JObject
        Try
            receipt = ReadAutoPilotPromptCoachingJson(path)
        Catch ex As System.IO.FileNotFoundException
            Return New Newtonsoft.Json.Linq.JObject From {{"schema", 1}, {"sender", sender}, {"claims", New Newtonsoft.Json.Linq.JArray()}}
        Catch ex As System.IO.DirectoryNotFoundException
            Return New Newtonsoft.Json.Linq.JObject From {{"schema", 1}, {"sender", sender}, {"claims", New Newtonsoft.Json.Linq.JArray()}}
        End Try
        If receipt("schema") Is Nothing OrElse receipt("schema").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
           receipt.Value(Of System.Int32)("schema") <> 1 OrElse
           receipt("sender") Is Nothing OrElse receipt("sender").Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse
           receipt.Value(Of System.String)("sender") <> sender Then Throw New System.IO.InvalidDataException("Invalid recommendation receipt identity or schema.")
        Dim claims As Newtonsoft.Json.Linq.JArray = TryCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
        If claims Is Nothing Then Throw New System.IO.InvalidDataException("Missing recommendation submission claims.")
        If receipt("lastPreflightFailureUtc") IsNot Nothing Then ReadAutoPilotPromptCoachingUtc(receipt, "lastPreflightFailureUtc")
        Dim seenClaims As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each item As Newtonsoft.Json.Linq.JToken In claims
            Dim claim As Newtonsoft.Json.Linq.JObject = TryCast(item, Newtonsoft.Json.Linq.JObject)
            If claim Is Nothing Then Throw New System.IO.InvalidDataException("Invalid recommendation submission claim.")
            For Each name As System.String In New System.String() {"claimId", "state", "mode", "localDate"}
                If claim(name) Is Nothing OrElse claim(name).Type <> Newtonsoft.Json.Linq.JTokenType.String Then
                    Throw New System.IO.InvalidDataException("Incomplete recommendation submission claim.")
                End If
            Next
            Dim claimId As System.String = claim.Value(Of System.String)("claimId")
            If Not System.Text.RegularExpressions.Regex.IsMatch(claimId, "\A[a-f0-9]{32}\z", System.Text.RegularExpressions.RegexOptions.CultureInvariant,
                                                               System.TimeSpan.FromSeconds(1)) OrElse Not seenClaims.Add(claimId) Then
                Throw New System.IO.InvalidDataException("Invalid or duplicate recommendation claim identity.")
            End If
            Select Case claim.Value(Of System.String)("state")
                Case "submitted" : ReadAutoPilotPromptCoachingUtc(claim, "submittedUtc")
                Case "submission_pending", "submission_uncertain"
                Case Else : Throw New System.IO.InvalidDataException("Invalid recommendation submission state.")
            End Select
            Dim mode As System.String = claim.Value(Of System.String)("mode")
            If mode <> "daily" AndAlso mode <> "immediate" Then Throw New System.IO.InvalidDataException("Invalid recommendation delivery mode.")
            Dim localDate As System.DateTime
            If Not System.DateTime.TryParseExact(claim.Value(Of System.String)("localDate"), "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture,
                                                 System.Globalization.DateTimeStyles.None, localDate) Then Throw New System.IO.InvalidDataException("Invalid recommendation delivery date.")
            ReadAutoPilotPromptCoachingUtc(claim, "reservedUtc")
            Dim topics As System.Collections.Generic.HashSet(Of System.String) = ReadAutoPilotPromptCoachingTopicSet(claim("topics"), True, 3)
            Dim runTopics As Newtonsoft.Json.Linq.JArray = TryCast(claim("runTopics"), Newtonsoft.Json.Linq.JArray)
            Dim claimedRunIds As Newtonsoft.Json.Linq.JArray = TryCast(claim("claimedRunIds"), Newtonsoft.Json.Linq.JArray)
            If runTopics Is Nothing OrElse runTopics.Count = 0 OrElse claimedRunIds Is Nothing OrElse claimedRunIds.Count <> runTopics.Count Then
                Throw New System.IO.InvalidDataException("Incomplete recommendation run claims.")
            End If
            Dim runIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim mappedTopics As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each runToken As Newtonsoft.Json.Linq.JToken In runTopics
                Dim run As Newtonsoft.Json.Linq.JObject = TryCast(runToken, Newtonsoft.Json.Linq.JObject)
                If run Is Nothing OrElse run("runId") Is Nothing OrElse run("runId").Type <> Newtonsoft.Json.Linq.JTokenType.String Then
                    Throw New System.IO.InvalidDataException("Invalid recommendation run claim.")
                End If
                Dim runId As System.String = run.Value(Of System.String)("runId")
                If Not System.Text.RegularExpressions.Regex.IsMatch(runId, "\A[a-f0-9]{24}\z", System.Text.RegularExpressions.RegexOptions.CultureInvariant,
                                                                   System.TimeSpan.FromSeconds(1)) OrElse Not runIds.Add(runId) Then Throw New System.IO.InvalidDataException("Invalid or duplicate recommendation run identity.")
                Dim assigned As System.Collections.Generic.HashSet(Of System.String) = ReadAutoPilotPromptCoachingTopicSet(run("topics"), True, 3)
                If Not assigned.IsSubsetOf(topics) Then Throw New System.IO.InvalidDataException("Recommendation run topics disagree with the submission claim.")
                mappedTopics.UnionWith(assigned)
            Next
            Dim listed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each id As Newtonsoft.Json.Linq.JToken In claimedRunIds
                If id.Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse Not listed.Add(id.ToObject(Of System.String)()) Then Throw New System.IO.InvalidDataException("Invalid claimed recommendation run list.")
            Next
            If Not listed.SetEquals(runIds) OrElse Not mappedTopics.SetEquals(topics) Then Throw New System.IO.InvalidDataException("Inconsistent recommendation receipt mappings.")
        Next
        Return receipt
    End Function

    Private Shared Function IsAutoPilotPromptCoachingDailyDeliveryDue(receipt As Newtonsoft.Json.Linq.JObject, nowLocal As System.DateTime) As System.Boolean
        If nowLocal.TimeOfDay < New System.TimeSpan(20, 0, 0) Then Return False
        Dim day As System.String = nowLocal.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
        For Each claim As Newtonsoft.Json.Linq.JObject In DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
            ' Switching from immediate to daily must not trigger an additional mail
            ' on a day which already has a submitted or uncertain recommendation.
            If claim.Value(Of System.String)("localDate") = day Then Return False
        Next
        Return True
    End Function

    Private Shared Function GetAutoPilotPromptCoachingClaimState(receipt As Newtonsoft.Json.Linq.JObject, runId As System.String, topic As System.String) As System.String
        Dim result As System.String = System.String.Empty
        For Each claim As Newtonsoft.Json.Linq.JObject In DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
            For Each run As Newtonsoft.Json.Linq.JObject In DirectCast(claim("runTopics"), Newtonsoft.Json.Linq.JArray)
                If run.Value(Of System.String)("runId") <> runId Then Continue For
                For Each assigned As Newtonsoft.Json.Linq.JToken In DirectCast(run("topics"), Newtonsoft.Json.Linq.JArray)
                    If assigned.ToObject(Of System.String)() <> topic Then Continue For
                    Dim state As System.String = claim.Value(Of System.String)("state")
                    If state <> "submitted" Then Return "submission_uncertain"
                    result = "submitted"
                Next
            Next
        Next
        Return result
    End Function

    Private Shared Function GetAutoPilotPromptCoachingRecentTopics(receipt As Newtonsoft.Json.Linq.JObject, nowUtc As System.DateTime) As System.Collections.Generic.HashSet(Of System.String)
        Dim result As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        Dim cutoff As System.DateTime = nowUtc.AddDays(-AP_PromptCoachingTopicCooldownDays)
        For Each claim As Newtonsoft.Json.Linq.JObject In DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
            If ReadAutoPilotPromptCoachingUtc(claim, "reservedUtc") < cutoff Then Continue For
            result.UnionWith(ReadAutoPilotPromptCoachingTopicSet(claim("topics"), True, 3))
        Next
        Return result
    End Function

    Private Function ReconcileAutoPilotPromptCoachingDeliveryJob(job As Newtonsoft.Json.Linq.JObject,
                                                                receipt As Newtonsoft.Json.Linq.JObject,
                                                                recentTopics As System.Collections.Generic.HashSet(Of System.String)) As AutoPilotPromptCoachingDeliveryItem
        Dim item As New AutoPilotPromptCoachingDeliveryItem With {
            .Job = job, .Result = Global.SharedLibrary.Agents.PromptCoachingAdvice.Result.FromJson(TryCast(job("analysis"), Newtonsoft.Json.Linq.JObject))
        }
        For Each note As System.String In item.Result.ValidationNotes
            LogAutoPilotLogDiagnosticsStatus("coaching-policy", "run " & job.Value(Of System.String)("runId") & ": " & note)
        Next
        If item.Result.Recommendations.Count = 0 Then
            FinishAutoPilotPromptCoachingJob(job, If(item.Result.ValidationNotes.Count > 0, "suppressed", "no_advice"))
            Return item
        End If
        Dim suppressed As System.Collections.Generic.HashSet(Of System.String) =
            If(job("suppressedTopics") Is Nothing,
               New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal),
               ReadAutoPilotPromptCoachingTopicSet(job("suppressedTopics"), False, 3))
        Dim suppressedChanged As System.Boolean = False
        Dim submitted As System.Boolean = False
        Dim uncertain As System.Boolean = False
        For Each advice As Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation In item.Result.Recommendations
            Dim state As System.String = GetAutoPilotPromptCoachingClaimState(receipt, job.Value(Of System.String)("runId"), advice.Topic)
            If state.Length > 0 Then
                submitted = submitted OrElse state = "submitted"
                uncertain = uncertain OrElse state = "submission_uncertain"
            ElseIf suppressed.Contains(advice.Topic) Then
                ' Previously documented cooldown suppression remains terminal.
            ElseIf recentTopics.Contains(advice.Topic) Then
                suppressed.Add(advice.Topic)
                suppressedChanged = True
            Else
                item.Remaining.Add(advice)
            End If
        Next
        If suppressedChanged Then
            Dim ordered As New System.Collections.Generic.List(Of System.String)(suppressed)
            ordered.Sort(System.StringComparer.Ordinal)
            job("suppressedTopics") = New Newtonsoft.Json.Linq.JArray(ordered)
            job("suppressionReason") = "topic_already_submitted_or_reserved_within_7_days"
            LogAutoPilotLogDiagnosticsStatus("coaching-repeat", "run " & job.Value(Of System.String)("runId") & " duplicate advice suppressed by the per-user seven-day topic cooldown.")
        End If
        If item.Remaining.Count = 0 Then
            FinishAutoPilotPromptCoachingJob(job, If(uncertain, "submission_uncertain", If(submitted, "submitted", "suppressed")))
        ElseIf suppressedChanged OrElse item.Remaining.Count < item.Result.Recommendations.Count Then
            ' Retain prose only for still-pending advice, even when other topics in
            ' the same run must wait for a later digest. Receipts keep the audit trail.
            Dim pendingAdvice As New Newtonsoft.Json.Linq.JArray()
            For Each remaining As Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation In item.Remaining
                pendingAdvice.Add(remaining.ToJson())
            Next
            job("analysis") = New Newtonsoft.Json.Linq.JObject From {
                {"schema", Global.SharedLibrary.Agents.PromptCoachingAdvice.CurrentSchema},
                {"language", item.Result.Language}, {"recommendations", pendingAdvice}
            }
            SyncLock _apPromptCoachingFsSync
                WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingPath("jobs", job.Value(Of System.String)("runId"), ".json"), job)
            End SyncLock
        End If
        Return item
    End Function

    Private Async Function MaybeSendAutoPilotPromptCoachingAsync(ct As System.Threading.CancellationToken,
                                                               Optional immediateOnly As System.Boolean = False) As System.Threading.Tasks.Task
        If Not IsAutoPilotPromptCoachingEnabled() Then Return
        ct.ThrowIfCancellationRequested()
        If Not IsAutoPilotLogDiagnosticsMailIdle() Then Return
        Dim bySender As New System.Collections.Generic.SortedDictionary(Of System.String, System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject))(System.StringComparer.Ordinal)
        For Each job As Newtonsoft.Json.Linq.JObject In GetAutoPilotPromptCoachingJobs()
            If job.Value(Of System.String)("status") <> "analysed" Then Continue For
            Dim sender As System.String = job.Value(Of System.String)("sender")
            If Not bySender.ContainsKey(sender) Then bySender(sender) = New System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)()
            bySender(sender).Add(job)
        Next
        For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject)) In bySender
            ct.ThrowIfCancellationRequested()
            If Not IsAutoPilotLogDiagnosticsMailIdle() Then Throw New System.OperationCanceledException(ct)
            Try
                Await DeliverAutoPilotPromptCoachingForUserAsync(pair.Key, pair.Value, ct, immediateOnly).ConfigureAwait(False)
            Catch ex As System.OperationCanceledException When ct.IsCancellationRequested OrElse Not IsAutoPilotLogDiagnosticsMailIdle()
                Throw
            Catch ex As System.Exception
                ' One corrupt receipt or failed account/recipient must not block other users.
                ApDashboardLog("Inky recommendation delivery withheld for user key " & Global.SharedLibrary.Agents.ToolingOriginalRequestLog.ComputeRequestSha256(pair.Key).Substring(0, 12) &
                               ": " & ex.GetType().Name & ". Existing receipts are retained; no state was reset.", "warn")
            End Try
        Next
    End Function

    Private Async Function DeliverAutoPilotPromptCoachingForUserAsync(sender As System.String,
                                                                     jobs As System.Collections.Generic.List(Of Newtonsoft.Json.Linq.JObject),
                                                                     ct As System.Threading.CancellationToken,
                                                                     Optional immediateOnly As System.Boolean = False) As System.Threading.Tasks.Task
        Dim preference As Newtonsoft.Json.Linq.JObject = GetAutoPilotPromptCoachingPreferenceStatus(sender)
        If Not preference.Value(Of System.Boolean)("effectiveEnabled") Then
            If preference.Value(Of System.String)("preference") = "disabled" OrElse preference.Value(Of System.String)("reason") = "opt_in_required" Then
                For Each job As Newtonsoft.Json.Linq.JObject In jobs
                    FinishAutoPilotPromptCoachingJob(job, "suppressed")
                Next
            End If
            Return
        End If
        Dim isDaily As System.Boolean = preference.Value(Of System.String)("deliveryMode") <> "immediate"
        If immediateOnly AndAlso isDaily Then Return
        Dim receipt As Newtonsoft.Json.Linq.JObject
        SyncLock _apPromptCoachingFsSync
            receipt = ReadAutoPilotPromptCoachingReceipt(sender)
            Dim recoveredPending As System.Boolean = False
            For Each previousClaim As Newtonsoft.Json.Linq.JObject In DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
                If previousClaim.Value(Of System.String)("state") = "submission_pending" Then
                    previousClaim("state") = "submission_uncertain"
                    previousClaim("outcomeRecordedUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                    recoveredPending = True
                End If
            Next
            If recoveredPending Then
                WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingReceiptPath(sender), receipt)
                ApDashboardLog("Inky recommendation submission from an earlier interrupted pass is uncertain; its durable claims prevent automatic resubmission.", "warn")
            End If
        End SyncLock
        Dim recentTopics As System.Collections.Generic.HashSet(Of System.String) = GetAutoPilotPromptCoachingRecentTopics(receipt, System.DateTime.UtcNow)
        Dim items As New System.Collections.Generic.List(Of AutoPilotPromptCoachingDeliveryItem)()
        Dim byTopic As New System.Collections.Generic.Dictionary(Of System.String, Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation)(System.StringComparer.Ordinal)
        Dim language As System.String = "en"
        For Each job As Newtonsoft.Json.Linq.JObject In jobs
            ' Immediate mail belongs to one original request, not an arbitrary merged
            ' backlog. Daily delivery can collect several request-specific blocks.
            If Not isDaily AndAlso items.Count > 0 Then Exit For
            Dim item As AutoPilotPromptCoachingDeliveryItem = ReconcileAutoPilotPromptCoachingDeliveryJob(job, receipt, recentTopics)
            If item.Remaining.Count = 0 Then Continue For
            items.Add(item)
            language = item.Result.Language
            For Each advice As Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation In item.Remaining
                If Not byTopic.ContainsKey(advice.Topic) OrElse byTopic(advice.Topic).Confidence < advice.Confidence Then byTopic(advice.Topic) = advice
            Next
        Next
        If byTopic.Count = 0 Then Return
        If isDaily AndAlso Not IsAutoPilotPromptCoachingDailyDeliveryDue(receipt, System.DateTime.Now) Then Return
        If receipt("lastPreflightFailureUtc") IsNot Nothing AndAlso
           (System.DateTime.UtcNow - ReadAutoPilotPromptCoachingUtc(receipt, "lastPreflightFailureUtc")).TotalMinutes < AP_PromptCoachingPreflightRetryMinutes Then Return

        Dim selected As New System.Collections.Generic.List(Of Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation)(byTopic.Values)
        selected.Sort(Function(left, right)
                          Dim rank As System.Int32 = right.Confidence.CompareTo(left.Confidence)
                          Return If(rank <> 0, rank, System.StringComparer.Ordinal.Compare(left.Topic, right.Topic))
                      End Function)
        If selected.Count > 3 Then selected.RemoveRange(3, selected.Count - 3)
        Dim selectedTopics As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each advice As Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation In selected
            selectedTopics.Add(advice.Topic)
        Next
        Dim orderedTopics As New System.Collections.Generic.List(Of System.String)(selectedTopics)
        orderedTopics.Sort(System.StringComparer.Ordinal)
        Dim runTopics As New Newtonsoft.Json.Linq.JArray()
        Dim claimedRunIds As New Newtonsoft.Json.Linq.JArray()
        For Each item As AutoPilotPromptCoachingDeliveryItem In items
            Dim assigned As New System.Collections.Generic.List(Of System.String)()
            For Each advice As Global.SharedLibrary.Agents.PromptCoachingAdvice.Recommendation In item.Remaining
                ' A contextual example represents this exact request only. Do not
                ' claim another request as delivered merely for sharing a topic.
                If selectedTopics.Contains(advice.Topic) AndAlso System.Object.ReferenceEquals(byTopic(advice.Topic), advice) Then assigned.Add(advice.Topic)
            Next
            If assigned.Count = 0 Then Continue For
            assigned.Sort(System.StringComparer.Ordinal)
            Dim runId As System.String = item.Job.Value(Of System.String)("runId")
            runTopics.Add(New Newtonsoft.Json.Linq.JObject From {{"runId", runId}, {"topics", New Newtonsoft.Json.Linq.JArray(assigned)}})
            claimedRunIds.Add(runId)
        Next
        Dim body As System.String = Global.SharedLibrary.Agents.PromptCoachingAdvice.RenderHtml(language, selected, isDaily,
            System.DateTime.Now.ToString("dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture), _apConfig.SubjectTriggerWord)
        If System.String.IsNullOrWhiteSpace(body) OrElse runTopics.Count = 0 Then Return
        Dim claim As New Newtonsoft.Json.Linq.JObject From {
            {"claimId", System.Guid.NewGuid().ToString("N")}, {"state", "submission_pending"},
            {"mode", If(isDaily, "daily", "immediate")}, {"topics", New Newtonsoft.Json.Linq.JArray(orderedTopics)},
            {"claimedRunIds", claimedRunIds}, {"runTopics", runTopics}
        }
        Dim claimed As System.Boolean = False
        Dim submitted As System.Boolean = False
        Try
            Await SwitchToUi(Sub()
                                 submitted = SendAutoPilotPromptCoachingMail(sender, body, language, isDaily, receipt, claim, ct, claimed)
                             End Sub).ConfigureAwait(False)
            If Not claimed Then Return
            claim("state") = "submitted"
            claim("submittedUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
            SyncLock _apPromptCoachingFsSync
                WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingReceiptPath(sender), receipt)
            End SyncLock
            ApDashboardLog("Inky recommendation mail submitted to Outlook for user key " & Global.SharedLibrary.Agents.ToolingOriginalRequestLog.ComputeRequestSha256(sender).Substring(0, 12) &
                           " (" & selected.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " topic(s)).", "info")
        Catch ex As System.OperationCanceledException When Not claimed AndAlso (ct.IsCancellationRequested OrElse Not IsAutoPilotLogDiagnosticsMailIdle())
            Throw
        Catch ex As System.Exception
            If Not claimed Then
                ' A failed atomic write must not leave an in-memory tentative claim
                ' in the later preflight-failure update. Re-read the durable version;
                ' if the claim did reach disk, conservatively retain it as uncertain.
                SyncLock _apPromptCoachingFsSync
                    receipt = ReadAutoPilotPromptCoachingReceipt(sender)
                    For Each storedClaim As Newtonsoft.Json.Linq.JObject In DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray)
                        If storedClaim.Value(Of System.String)("claimId") = claim.Value(Of System.String)("claimId") Then
                            claim = storedClaim
                            claimed = True
                            Exit For
                        End If
                    Next
                End SyncLock
            End If
            If claimed Then
                ' A durable pending claim is enough to prevent a repeat after restart,
                ' even if writing the clearer outcome below also fails.
                claim("state") = If(submitted, "submitted", "submission_uncertain")
                If submitted Then claim("submittedUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                claim("outcomeRecordedUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                Try
                    SyncLock _apPromptCoachingFsSync
                        WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingReceiptPath(sender), receipt)
                    End SyncLock
                Catch persistError As System.Exception
                    ApDashboardLog("Inky recommendation outcome could not be persisted; the durable pending claim still prevents resubmission: " & persistError.GetType().Name, "warn")
                End Try
                ApDashboardLog(If(submitted,
                    "Inky recommendation mail was accepted by Outlook, but its receipt update failed; no automatic resubmission. ",
                    "Inky recommendation submission is uncertain; no automatic resubmission. ") & ex.GetType().Name, "warn")
            Else
                ' No Send was called and no claim was committed: a later retry is safe.
                receipt("lastPreflightFailureUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                SyncLock _apPromptCoachingFsSync
                    WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingReceiptPath(sender), receipt)
                End SyncLock
                ApDashboardLog("Inky recommendation preflight failed before submission; retained for retry after five minutes: " & ex.GetType().Name, "warn")
                Return
            End If
        End Try
        recentTopics = GetAutoPilotPromptCoachingRecentTopics(receipt, System.DateTime.UtcNow)
        For Each item As AutoPilotPromptCoachingDeliveryItem In items
            ReconcileAutoPilotPromptCoachingDeliveryJob(item.Job, receipt, recentTopics)
        Next
        ' Unselected topics remain analysed. Daily mode carries them into a future
        ' digest; immediate mode picks them up in a subsequent idle delivery pass.
    End Function

    ''' <summary>Must execute synchronously on the Outlook STA through SwitchToUi.</summary>
    Private Function SendAutoPilotPromptCoachingMail(sender As System.String,
                                                    body As System.String,
                                                    language As System.String,
                                                    isDaily As System.Boolean,
                                                    ByRef receipt As Newtonsoft.Json.Linq.JObject,
                                                    claim As Newtonsoft.Json.Linq.JObject,
                                                    ct As System.Threading.CancellationToken,
                                                    ByRef claimed As System.Boolean) As System.Boolean
        If NormalizeAutoPilotPromptCoachingEmail(sender) <> sender OrElse sender.Length = 0 Then Throw New System.InvalidOperationException("Recommendation recipient is not a valid original sender.")
        Dim mail As Microsoft.Office.Interop.Outlook.MailItem = Nothing
        Dim accounts As Microsoft.Office.Interop.Outlook.Accounts = Nothing
        Dim sendAccount As Microsoft.Office.Interop.Outlook.Account = Nothing
        Dim recipients As Microsoft.Office.Interop.Outlook.Recipients = Nothing
        Dim resolvedRecipient As Microsoft.Office.Interop.Outlook.Recipient = Nothing
        Try
            ct.ThrowIfCancellationRequested()
            If Not IsAutoPilotLogDiagnosticsMailIdle() Then Throw New System.OperationCanceledException(ct)
            mail = DirectCast(Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem), Microsoft.Office.Interop.Outlook.MailItem)
            mail.To = sender
            mail.CC = System.String.Empty
            mail.BCC = System.String.Empty
            Select Case language
                Case "de" : mail.Subject = "Inky: Tipps für Ihre Aufträge"
                Case "fr" : mail.Subject = "Inky : conseils pour vos demandes"
                Case "it" : mail.Subject = "Inky: suggerimenti per le vostre richieste"
                Case Else : mail.Subject = "Inky: tips for your requests"
            End Select
            mail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML
            mail.HTMLBody = "<html><body>" & body & BuildAutoPilotFooter() & "</body></html>"
            StampAutoPilotReplyHeaders(mail, AP_LoopHeaderValue)
            accounts = Application.Session.Accounts
            If accounts Is Nothing OrElse accounts.Count = 0 Then Throw New System.InvalidOperationException("No Outlook sending account is configured.")
            Dim monitored As System.String = If(_apConfig?.MonitoredMailbox, System.String.Empty)
            If monitored.Length > 0 Then
                For index As System.Int32 = 1 To accounts.Count
                    Dim candidate As Microsoft.Office.Interop.Outlook.Account = Nothing
                    Try
                        candidate = accounts.Item(index)
                        If candidate IsNot Nothing AndAlso System.String.Equals(candidate.SmtpAddress, monitored, System.StringComparison.OrdinalIgnoreCase) Then
                            sendAccount = candidate
                            candidate = Nothing
                            Exit For
                        End If
                    Finally
                        If candidate IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(candidate) : Catch ex As System.Exception : End Try
                        End If
                    End Try
                Next
            End If
            If sendAccount Is Nothing AndAlso accounts.Count = 1 Then sendAccount = accounts.Item(1)
            If sendAccount Is Nothing Then Throw New System.InvalidOperationException("The Outlook account for Inky recommendations could not be determined uniquely.")
            mail.SendUsingAccount = sendAccount
            recipients = mail.Recipients
            If recipients Is Nothing OrElse recipients.Count <> 1 OrElse Not recipients.ResolveAll() Then Throw New System.InvalidOperationException("The original sender could not be resolved uniquely.")
            ' Address-book resolution can replace or expand what was assigned to To.
            ' Validate the resulting envelope, not only the requested display string.
            If recipients.Count <> 1 Then Throw New System.InvalidOperationException("Recommendation recipient resolution changed the envelope cardinality.")
            resolvedRecipient = recipients.Item(1)
            If resolvedRecipient Is Nothing OrElse Not resolvedRecipient.Resolved OrElse
               resolvedRecipient.Type <> System.Convert.ToInt32(Microsoft.Office.Interop.Outlook.OlMailRecipientType.olTo) OrElse
               NormalizeAutoPilotPromptCoachingEmail(GetAutoPilotRecipientSmtpAddress(resolvedRecipient)) <> sender Then
                Throw New System.InvalidOperationException("The resolved recommendation recipient does not match the original sender.")
            End If

            ' Keep the eligibility check and reservation adjacent to Submit. A user/admin
            ' preference write cannot interleave between the final check and Send.
            SyncLock _apPromptCoachingPreferenceGate
                ct.ThrowIfCancellationRequested()
                If Not IsAutoPilotLogDiagnosticsMailIdle() Then Throw New System.OperationCanceledException(ct)
                Dim latestPreference As Newtonsoft.Json.Linq.JObject = GetAutoPilotPromptCoachingPreferenceStatus(sender)
                If Not latestPreference.Value(Of System.Boolean)("effectiveEnabled") Then Return False
                If isDaily <> (latestPreference.Value(Of System.String)("deliveryMode") <> "immediate") Then Return False
                SyncLock _apPromptCoachingFsSync
                    receipt = ReadAutoPilotPromptCoachingReceipt(sender)
                    If isDaily AndAlso Not IsAutoPilotPromptCoachingDailyDeliveryDue(receipt, System.DateTime.Now) Then Return False
                    Dim recent As System.Collections.Generic.HashSet(Of System.String) = GetAutoPilotPromptCoachingRecentTopics(receipt, System.DateTime.UtcNow)
                    If recent.Overlaps(ReadAutoPilotPromptCoachingTopicSet(claim("topics"), True, 3)) Then Return False
                    For Each run As Newtonsoft.Json.Linq.JObject In DirectCast(claim("runTopics"), Newtonsoft.Json.Linq.JArray)
                        For Each topic As System.String In ReadAutoPilotPromptCoachingTopicSet(run("topics"), True, 3)
                            If GetAutoPilotPromptCoachingClaimState(receipt, run.Value(Of System.String)("runId"), topic).Length > 0 Then Return False
                        Next
                    Next
                    claim("reservedUtc") = System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture)
                    claim("localDate") = System.DateTime.Now.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)
                    DirectCast(receipt("claims"), Newtonsoft.Json.Linq.JArray).Add(claim)
                    receipt.Remove("lastPreflightFailureUtc")
                    WriteAutoPilotPromptCoachingJson(GetAutoPilotPromptCoachingReceiptPath(sender), receipt)
                    claimed = True
                End SyncLock
                SubmitAutoPilotMailOnce(mail)
            End SyncLock
            Return True
        Finally
            If resolvedRecipient IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(resolvedRecipient) : Catch ex As System.Exception : End Try
            End If
            If recipients IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients) : Catch ex As System.Exception : End Try
            End If
            If sendAccount IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(sendAccount) : Catch ex As System.Exception : End Try
            End If
            If accounts IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts) : Catch ex As System.Exception : End Try
            End If
            If mail IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.ReleaseComObject(mail) : Catch ex As System.Exception : End Try
            End If
        End Try
    End Function

End Class
