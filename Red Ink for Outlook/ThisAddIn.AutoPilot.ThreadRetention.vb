' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.ThreadRetention.vb
' Purpose:
'   Optional AutoPilot feature (AutoPilot session only — NOT the Local Agent):
'   retains a sender's incoming attachments (e.g. a supplied contract) for a
'   configurable number of days so that FOLLOW-UP mails from the SAME sender in
'   the SAME conversation/topic can continue the discussion with the earlier
'   files still available to the model.
'
' Security model:
'   - Storage key is derived from the AUTHENTICATED inbound sender address and
'     the conversation identity — never from user-supplied path input.
'       {INI dir}/autopilot_users/{canonical_email_identity}/threads/{conversationKey}/
'           files/       — retained inbound attachments
'           meta.json    — DeleteAfterUtc / LastTouchedUtc / ConversationId / Subject
'   - Cross-sender isolation: the folder is a deterministic function of the
'     sender's SMTP address through GetUserDir() (SHA-256-backed canonical identity). A different sender maps
'     to a different folder and physically cannot read another sender's files.
'   - Topic binding: files are only reloaded when the ConversationID (or, as a
'     fallback, the normalized subject) matches, so an unrelated new mail from
'     the same sender will not pull in an old contract.
'   - Trust gate: retention is applied ONLY for whitelisted senders. Because
'     "same sender" trusts the From address, restricting to the operator's
'     whitelist prevents spoofed senders from planting or reading files.
'   - Path traversal: every resolved path is validated with IsPathContained.
'   - Retention bound: a hard DeleteAfterUtc plus per-thread size/file caps.
'   - The per-mail isolated temp directory and its Finally wipe are untouched;
'     retained files are a separate, opt-in copy.
'
' Threading:
'   - AutoPilot processes one mail at a time (processing semaphore), so file I/O
'     here is effectively serialized. Voicemail/scheduler paths do not call in.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Const AP_ThreadRetentionSubdir As String = "threads"
    Private Const AP_ThreadFilesSubdir As String = "files"
    Private Const AP_ThreadMetaFileName As String = "meta.json"

    ''' <summary>Per-thread retention size cap (25 MB).</summary>
    Private Const AP_ThreadRetentionMaxBytesPerThread As Long = 25 * 1024 * 1024

    ''' <summary>Per-thread retention file-count cap.</summary>
    Private Const AP_ThreadRetentionMaxFilesPerThread As Integer = 50

    ''' <summary>
    ''' True when thread retention is enabled for the current session
    ''' (AutoPilot only). ThreadRetentionDays = 0 disables the feature.
    ''' </summary>
    Private Function IsThreadRetentionEnabled() As Boolean
        Return _apActive AndAlso
               _apConfig IsNot Nothing AndAlso
               _apConfig.ThreadRetentionDays > 0
    End Function

    ''' <summary>
    ''' Retention is applied only to whitelisted senders. "Same sender" trusts the
    ''' From address, so restricting to the operator's auto-send whitelist prevents
    ''' spoofed senders from planting or reading retained files.
    ''' </summary>
    Private Function IsThreadRetentionEligible(senderEmail As String, isWhitelisted As Boolean) As Boolean
        If Not IsThreadRetentionEnabled() Then Return False
        If Not isWhitelisted Then Return False
        If String.IsNullOrWhiteSpace(senderEmail) Then Return False
        Return True
    End Function

    ''' <summary>Root retention directory for a sender: {user dir}/threads/.</summary>
    Private Function GetThreadRetentionRootForSender(senderEmail As String) As String
        Dim root = Path.Combine(GetUserDir(senderEmail), AP_ThreadRetentionSubdir)
        If Not Directory.Exists(root) Then Directory.CreateDirectory(root)
        Return root
    End Function

    ''' <summary>
    ''' Strips reply/forward prefixes and normalizes a subject for use as a
    ''' fallback conversation key when no ConversationID is available.
    ''' </summary>
    Private Shared Function NormalizeSubjectForThreadKey(subject As String) As String
        If String.IsNullOrWhiteSpace(subject) Then Return ""
        Dim s = subject.Trim()
        ' Repeatedly strip common reply/forward prefixes (multi-language).
        Dim prefixPattern As String = "^\s*(re|fw|fwd|aw|wg|antw|sv|vs|tr|rif)\s*(\[\d+\])?\s*:\s*"
        Dim guard As Integer = 0
        While guard < 20
            Dim stripped = Regex.Replace(s, prefixPattern, "", RegexOptions.IgnoreCase)
            If stripped.Equals(s, StringComparison.Ordinal) Then Exit While
            s = stripped.Trim()
            guard += 1
        End While
        ' Collapse whitespace, lowercase for stable hashing.
        s = Regex.Replace(s, "\s+", " ").Trim().ToLowerInvariant()
        Return s
    End Function

    ''' <summary>
    ''' Computes a deterministic, filesystem-safe conversation key for a mail.
    ''' Prefers the ConversationID; falls back to the normalized subject.
    ''' Returns Nothing when neither is available (retention cannot be keyed safely).
    ''' </summary>
    Private Shared Function ComputeThreadConversationKey(mailInfo As AutoPilotMailInfo) As String
        If mailInfo Is Nothing Then Return Nothing

        Dim seed As String = Nothing
        If Not String.IsNullOrWhiteSpace(mailInfo.ConversationID) Then
            seed = "cid:" & mailInfo.ConversationID.Trim()
        Else
            Dim normSubject = NormalizeSubjectForThreadKey(mailInfo.Subject)
            If Not String.IsNullOrWhiteSpace(normSubject) Then
                seed = "subj:" & normSubject
            End If
        End If

        If String.IsNullOrWhiteSpace(seed) Then Return Nothing

        Using sha As SHA256 = SHA256.Create()
            Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(seed))
            Dim sb As New StringBuilder(bytes.Length * 2)
            For Each b In bytes
                sb.Append(b.ToString("x2"))
            Next
            ' 32 hex chars is ample and keeps path lengths short.
            Return sb.ToString().Substring(0, 32)
        End Using
    End Function

    ''' <summary>
    ''' Resolves (and optionally creates) the thread directory for a sender+conversation.
    ''' Validates path containment. Returns Nothing on any safety failure.
    ''' </summary>
    Private Function GetThreadDirSafe(senderEmail As String,
                                      conversationKey As String,
                                      createIfMissing As Boolean) As String
        If String.IsNullOrWhiteSpace(conversationKey) Then Return Nothing

        Dim root = GetThreadRetentionRootForSender(senderEmail)
        Dim threadDir = Path.Combine(root, conversationKey)
        If Not IsPathContained(threadDir, root) Then Return Nothing

        If createIfMissing AndAlso Not Directory.Exists(threadDir) Then
            Directory.CreateDirectory(threadDir)
        End If
        Return threadDir
    End Function

    Private Function GetThreadFilesDirSafe(threadDir As String, createIfMissing As Boolean) As String
        If String.IsNullOrWhiteSpace(threadDir) Then Return Nothing
        Dim filesDir = Path.Combine(threadDir, AP_ThreadFilesSubdir)
        If Not IsPathContained(filesDir, threadDir) Then Return Nothing
        If createIfMissing AndAlso Not Directory.Exists(filesDir) Then
            Directory.CreateDirectory(filesDir)
        End If
        Return filesDir
    End Function

    ''' <summary>Reads the DeleteAfterUtc from a thread's meta.json, or Nothing.</summary>
    Private Function GetThreadDeleteAfterUtc(threadDir As String) As DateTime?
        Try
            Dim metaPath = Path.Combine(threadDir, AP_ThreadMetaFileName)
            If Not File.Exists(metaPath) Then Return Nothing
            Dim json = JObject.Parse(File.ReadAllText(metaPath))
            Dim raw = CStr(If(json("DeleteAfterUtc"), Nothing))
            If String.IsNullOrWhiteSpace(raw) Then Return Nothing
            Dim parsed As DateTime
            If DateTime.TryParse(raw, Globalization.CultureInfo.InvariantCulture,
                                 Globalization.DateTimeStyles.RoundtripKind, parsed) Then
                Return parsed.ToUniversalTime()
            End If
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>Writes/refreshes a thread's meta.json.</summary>
    Private Sub WriteThreadMeta(threadDir As String, mailInfo As AutoPilotMailInfo)
        Try
            Dim deleteAfter = DateTime.UtcNow.AddDays(_apConfig.ThreadRetentionDays)
            Dim json As New JObject From {
                {"ConversationId", If(mailInfo?.ConversationID, "")},
                {"Subject", If(mailInfo?.Subject, "")},
                {"LastTouchedUtc", DateTime.UtcNow.ToString("o", Globalization.CultureInfo.InvariantCulture)},
                {"DeleteAfterUtc", deleteAfter.ToString("o", Globalization.CultureInfo.InvariantCulture)}
            }
            File.WriteAllText(Path.Combine(threadDir, AP_ThreadMetaFileName),
                              json.ToString(Newtonsoft.Json.Formatting.Indented))
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] WriteThreadMeta error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Loads previously retained files for the sender's current conversation into the
    ''' per-mail temp directory and registers them as attachments for this session.
    ''' No-op unless retention is enabled AND the sender is whitelisted.
    ''' Expired threads are deleted and ignored.
    ''' </summary>
    Private Function LoadRetainedThreadFiles(mailInfo As AutoPilotMailInfo,
                                             tempDir As String,
                                             attachments As List(Of AutoPilotAttachmentInfo),
                                             isWhitelisted As Boolean) As Integer
        Try
            If mailInfo Is Nothing Then Return 0
            If Not IsThreadRetentionEligible(mailInfo.SenderEmail, isWhitelisted) Then Return 0
            If String.IsNullOrWhiteSpace(tempDir) OrElse Not Directory.Exists(tempDir) Then Return 0
            If attachments Is Nothing Then Return 0

            Dim convKey = ComputeThreadConversationKey(mailInfo)
            If String.IsNullOrWhiteSpace(convKey) Then Return 0

            Dim threadDir = GetThreadDirSafe(mailInfo.SenderEmail, convKey, createIfMissing:=False)
            If threadDir Is Nothing OrElse Not Directory.Exists(threadDir) Then Return 0

            ' Enforce expiry on load.
            Dim deleteAfter = GetThreadDeleteAfterUtc(threadDir)
            If deleteAfter.HasValue AndAlso DateTime.UtcNow > deleteAfter.Value Then
                Try : Directory.Delete(threadDir, recursive:=True) : Catch : End Try
                Return 0
            End If

            Dim filesDir = GetThreadFilesDirSafe(threadDir, createIfMissing:=False)
            If filesDir Is Nothing OrElse Not Directory.Exists(filesDir) Then Return 0

            Dim loaded As Integer = 0
            For Each src In Directory.GetFiles(filesDir)
                Try
                    Dim fileName = Path.GetFileName(src)
                    If String.IsNullOrWhiteSpace(fileName) Then Continue For

                    ' Dedupe against already-present (current inbound) attachments.
                    Dim destPath = Path.Combine(tempDir, fileName)
                    Dim counter = 1
                    While File.Exists(destPath)
                        destPath = Path.Combine(tempDir,
                            Path.GetFileNameWithoutExtension(fileName) & $"_{counter}" & Path.GetExtension(fileName))
                        counter += 1
                    End While

                    File.Copy(src, destPath)

                    attachments.Add(New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = Path.GetFileName(destPath),
                        .TempFilePath = destPath,
                        .Extension = Path.GetExtension(destPath).ToLowerInvariant(),
                        .SizeBytes = New FileInfo(destPath).Length,
                        .IsOverSizeLimit = False,
                        .StatusMessage = "Retained from an earlier message in this conversation",
                        .IsToolOutput = False,
                        .IsRetentionLoaded = True,
                        .OutputFiles = New List(Of String)()
                    })
                    loaded += 1
                Catch ex As System.Exception
                    Debug.WriteLine($"[AutoPilot] LoadRetainedThreadFiles copy error: {ex.Message}")
                End Try
            Next

            If loaded > 0 Then
                ApDashboardLog($"🗂 Loaded {loaded} retained file(s) from an earlier message in this conversation for {mailInfo.SenderEmail}.", "info")
            End If
            Return loaded
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] LoadRetainedThreadFiles error: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Persists this mail's genuinely-inbound attachments into the sender+conversation
    ''' retention store and (re)stamps the retention expiry so an active discussion keeps
    ''' the files alive. Skips retention-loaded, tool-output, and oversized files.
    ''' No-op unless retention is enabled AND the sender is whitelisted.
    ''' </summary>
    Private Sub PersistInboundAttachmentsToThread(mailInfo As AutoPilotMailInfo,
                                                  attachments As List(Of AutoPilotAttachmentInfo),
                                                  isWhitelisted As Boolean)
        Try
            If mailInfo Is Nothing Then Return
            If Not IsThreadRetentionEligible(mailInfo.SenderEmail, isWhitelisted) Then Return

            Dim convKey = ComputeThreadConversationKey(mailInfo)
            If String.IsNullOrWhiteSpace(convKey) Then Return

            Dim threadDir = GetThreadDirSafe(mailInfo.SenderEmail, convKey, createIfMissing:=True)
            If threadDir Is Nothing Then Return
            Dim filesDir = GetThreadFilesDirSafe(threadDir, createIfMissing:=True)
            If filesDir Is Nothing Then Return

            Dim stored As Integer = 0
            If attachments IsNot Nothing Then
                For Each att In attachments
                    Try
                        If att Is Nothing Then Continue For
                        If att.IsRetentionLoaded OrElse att.IsToolOutput OrElse att.IsOverSizeLimit Then Continue For
                        If String.IsNullOrWhiteSpace(att.TempFilePath) OrElse Not File.Exists(att.TempFilePath) Then Continue For

                        ' Bare filename only — never trust a path component.
                        Dim bareName = Path.GetFileName(If(att.OriginalFileName, Path.GetFileName(att.TempFilePath)))
                        If String.IsNullOrWhiteSpace(bareName) Then Continue For

                        Dim destPath = Path.Combine(filesDir, bareName)
                        If Not IsPathContained(destPath, filesDir) Then Continue For

                        File.Copy(att.TempFilePath, destPath, overwrite:=True)
                        stored += 1
                    Catch ex As System.Exception
                        Debug.WriteLine($"[AutoPilot] PersistInboundAttachmentsToThread copy error: {ex.Message}")
                    End Try
                Next
            End If

            Dim hasAnyFiles As Boolean = False
            Try : hasAnyFiles = Directory.GetFiles(filesDir).Length > 0 : Catch : End Try

            ' Only keep a thread alive if it actually holds files.
            If Not hasAnyFiles Then
                Try : Directory.Delete(threadDir, recursive:=True) : Catch : End Try
                Return
            End If

            EnforceThreadRetentionCaps(filesDir)
            WriteThreadMeta(threadDir, mailInfo)

            If stored > 0 Then
                ApDashboardLog($"🗂 Retained {stored} attachment(s) for follow-up discussion (up to {_apConfig.ThreadRetentionDays} day(s)).", "step")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] PersistInboundAttachmentsToThread error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>Prunes oldest files until the per-thread size and count caps are satisfied.</summary>
    Private Shared Sub EnforceThreadRetentionCaps(filesDir As String)
        Try
            If Not Directory.Exists(filesDir) Then Return
            Dim files = New DirectoryInfo(filesDir).GetFiles().OrderBy(Function(f) f.LastWriteTimeUtc).ToList()

            Dim totalBytes As Long = files.Sum(Function(f) f.Length)
            While files.Count > 0 AndAlso
                  (totalBytes > AP_ThreadRetentionMaxBytesPerThread OrElse files.Count > AP_ThreadRetentionMaxFilesPerThread)
                Dim oldest = files(0)
                totalBytes -= oldest.Length
                Try : oldest.Delete() : Catch : End Try
                files.RemoveAt(0)
            End While
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Deletes expired thread retention directories across all senders. Best-effort;
    ''' invoked at AutoPilot start when retention is enabled.
    ''' </summary>
    Private Sub PurgeExpiredThreadRetention()
        Try
            Dim rootDir = GetUserStorageRootDir()
            If Not Directory.Exists(rootDir) Then Return

            Dim purged As Integer = 0
            For Each userDir As String In EnumerateUserStorageDirectoriesForAdminAndMaintenance()
                Dim threadsRoot = Path.Combine(userDir, AP_ThreadRetentionSubdir)
                If Not Directory.Exists(threadsRoot) Then Continue For

                For Each threadDir In Directory.GetDirectories(threadsRoot)
                    Try
                        Dim deleteAfter = GetThreadDeleteAfterUtc(threadDir)
                        Dim filesDir = Path.Combine(threadDir, AP_ThreadFilesSubdir)
                        Dim isEmpty As Boolean = (Not Directory.Exists(filesDir)) OrElse Directory.GetFiles(filesDir).Length = 0

                        If isEmpty OrElse (deleteAfter.HasValue AndAlso DateTime.UtcNow > deleteAfter.Value) Then
                            Directory.Delete(threadDir, recursive:=True)
                            purged += 1
                        End If
                    Catch
                    End Try
                Next
            Next

            If purged > 0 Then
                ApDashboardLog($"🗂 Purged {purged} expired conversation retention folder(s).", "info")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] PurgeExpiredThreadRetention error: {ex.Message}")
        End Try
    End Sub

End Class
