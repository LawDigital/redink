' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365Service.vb
' Purpose: Microsoft 365 / Graph integration library for Red Ink add-ins.
'          Provides interactive sign-in (MSAL), Graph search across all major
'          M365 entity types, and per-element retrieval helpers (mail, files,
'          calendar, Teams chat, OneNote, People). Uses raw HttpClient + JObject
'          for consistency with the rest of the codebase.
'
' Cross-process safety:
'   The MSAL token cache is stored under %APPDATA%\RedInk\m365.msalcache.bin3
'   and protected by Microsoft.Identity.Client.Extensions.Msal which uses a
'   cross-process file lock — multiple Red Ink add-ins running in parallel can
'   safely share the same cached account.
'
' Public surface (high level):
'   Auth:        SignInAsync, SignOutAsync, IsSignedInAsync, GetSignedInUserAsync
'   Search:      SearchAsync (entry point, Or'able sources, IProgress, cancel)
'   Mail:        GetMessageAsync, GetMessageBodyAsync, ListAttachmentsAsync,
'                DownloadAttachmentAsync, DownloadAllAttachmentsAsync,
'                GetMessagesBatchAsync
'   Files:       GetDriveItemAsync, DownloadFileAsync
'   Calendar:    GetEventAsync
'   Teams:       GetChatMessageAsync
'   OneNote:     GetOneNotePageAsync (optionally fetches HTML content)
'   People:      GetPersonAsync
'   Low level:   GetAccessTokenAsync, GraphGetAsync, GraphPostAsync,
'                GraphPageAsync, GraphBatchAsync (exposed Friend for power users)
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Identity.Client
Imports Microsoft.Identity.Client.Extensions.Msal
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Static facade for all Microsoft 365 / Graph operations used by Red Ink add-ins.
    ''' </summary>
    Partial Public Module M365Service

        ' ───────────────────────────────────────────────────────────────────────
        '  CONSTANTS
        ' ───────────────────────────────────────────────────────────────────────

        Private Const GraphV1 As String = "https://graph.microsoft.com/v1.0"
        Private Const GraphBeta As String = "https://graph.microsoft.com/beta"

        Private Const TokenCacheFile As String = "m365.msalcache.bin3"
        Private Const TokenCacheClientIdKey As String = "RedInk.M365"

        Private Const DefaultTenant As String = "common"
        Private Const DefaultScopesIfMissing As String =
            "openid profile offline_access User.Read Mail.Read Files.Read Files.Read.All Sites.Read.All Calendars.Read Chat.Read ChannelMessage.Read.All Notes.Read.All People.Read Contacts.Read"

        ''' <summary>Single shared HttpClient for all Graph calls (recommended pattern).</summary>
        Private ReadOnly _http As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(120)}

        ''' <summary>Per-process MSAL app instance, lazily created.</summary>
        Private _app As IPublicClientApplication
        Private _appClientId As String
        Private _appTenantId As String
        Private ReadOnly _appLock As New Object()
        Private _cacheHelper As MsalCacheHelper

        ' ═══════════════════════════════════════════════════════════════════════
        '  AUTHENTICATION
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Ensures an interactive MSAL sign-in has been performed and a usable
        ''' access token can be acquired. Safe to call repeatedly — uses silent
        ''' acquisition when possible.
        ''' </summary>
        Public Async Function SignInAsync(context As ISharedContext,
                                          Optional ct As CancellationToken = Nothing) As Task(Of String)
            EnsureModernTls()

            Dim app = Await EnsureAppAsync(context).ConfigureAwait(False)
            Dim scopes = ParseScopes(context)

            ' Try silent first.
            Dim accounts = Await app.GetAccountsAsync().ConfigureAwait(False)
            Dim acct = accounts.FirstOrDefault()
            If acct IsNot Nothing Then
                Try
                    Dim silent = Await app.AcquireTokenSilent(scopes, acct).ExecuteAsync(ct).ConfigureAwait(False)
                    Return silent.AccessToken
                Catch ex As MsalUiRequiredException
                    ' fall through to interactive
                End Try
            End If

            Dim result = Await app.AcquireTokenInteractive(scopes) _
                                  .WithPrompt(Prompt.SelectAccount) _
                                  .ExecuteAsync(ct).ConfigureAwait(False)
            Return result.AccessToken
        End Function

        ''' <summary>Removes the cached account(s) and clears the token cache.</summary>
        Public Async Function SignOutAsync(context As ISharedContext) As Task
            Try
                Dim app = Await EnsureAppAsync(context).ConfigureAwait(False)
                Dim accounts = Await app.GetAccountsAsync().ConfigureAwait(False)
                For Each a In accounts.ToList()
                    Await app.RemoveAsync(a).ConfigureAwait(False)
                Next
            Catch ex As Exception
                Debug.WriteLine($"[M365] SignOut error: {ex.Message}")
            End Try
        End Function

        ''' <summary>Returns True when a silent token can be acquired (no UI shown).</summary>
        Public Async Function IsSignedInAsync(context As ISharedContext,
                                              Optional ct As CancellationToken = Nothing) As Task(Of Boolean)
            Try
                Dim token = Await TryGetTokenSilentAsync(context, ct).ConfigureAwait(False)
                Return Not String.IsNullOrEmpty(token)
            Catch
                Return False
            End Try
        End Function

        ''' <summary>Returns the signed-in user's UPN/email, or empty string.</summary>
        Public Async Function GetSignedInUserAsync(context As ISharedContext,
                                                   Optional ct As CancellationToken = Nothing) As Task(Of String)
            Try
                Dim app = Await EnsureAppAsync(context).ConfigureAwait(False)
                Dim accounts = Await app.GetAccountsAsync().ConfigureAwait(False)
                Return If(accounts.FirstOrDefault()?.Username, "")
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Acquires a Graph access token (silent if possible, interactive otherwise).
        ''' All public retrieval methods use this internally.
        ''' </summary>
        Public Async Function GetAccessTokenAsync(context As ISharedContext,
                                                  Optional ct As CancellationToken = Nothing) As Task(Of String)
            Dim silent = Await TryGetTokenSilentAsync(context, ct).ConfigureAwait(False)
            If Not String.IsNullOrEmpty(silent) Then Return silent
            Return Await SignInAsync(context, ct).ConfigureAwait(False)
        End Function

        Private Async Function TryGetTokenSilentAsync(context As ISharedContext,
                                                      ct As CancellationToken) As Task(Of String)
            Dim app = Await EnsureAppAsync(context).ConfigureAwait(False)
            Dim scopes = ParseScopes(context)
            Dim accounts = Await app.GetAccountsAsync().ConfigureAwait(False)
            Dim acct = accounts.FirstOrDefault()
            If acct Is Nothing Then Return ""
            Try
                Dim r = Await app.AcquireTokenSilent(scopes, acct).ExecuteAsync(ct).ConfigureAwait(False)
                Return r.AccessToken
            Catch ex As MsalUiRequiredException
                Return ""
            End Try
        End Function

        Private Async Function EnsureAppAsync(context As ISharedContext) As Task(Of IPublicClientApplication)
            EnsureModernTls()

            If context Is Nothing Then Throw New ArgumentNullException(NameOf(context))

            Dim clientId = If(context.INI_M365ClientId, "").Trim()
            If String.IsNullOrWhiteSpace(clientId) Then
                Throw New InvalidOperationException(
                    "INI_M365ClientId is not configured. Set the AAD application (client) id for the Red Ink add-in.")
            End If
            Dim tenantId = If(context.INI_M365TenantId, "").Trim()
            If String.IsNullOrWhiteSpace(tenantId) Then tenantId = DefaultTenant

            ' Re-build the app if client/tenant changed (defensive — should be stable).
            SyncLock _appLock
                If _app IsNot Nothing AndAlso
                   String.Equals(_appClientId, clientId, StringComparison.OrdinalIgnoreCase) AndAlso
                   String.Equals(_appTenantId, tenantId, StringComparison.OrdinalIgnoreCase) Then
                    Return _app
                End If
            End SyncLock

            Dim builder = PublicClientApplicationBuilder.Create(clientId) _
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}") _
                .WithDefaultRedirectUri()

            Dim app = builder.Build()

            ' Wire cross-process safe token cache.
            Dim cacheDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "RedInk")
            Try
                Directory.CreateDirectory(cacheDir)
                Dim storage = New StorageCreationPropertiesBuilder(TokenCacheFile, cacheDir) _
                    .WithCacheChangedEvent(clientId) _
                    .Build()
                Dim helper = Await MsalCacheHelper.CreateAsync(storage).ConfigureAwait(False)
                helper.RegisterCache(app.UserTokenCache)
                _cacheHelper = helper
            Catch ex As Exception
                ' Cache is best-effort: log and continue with in-memory cache.
                Debug.WriteLine($"[M365] Token cache unavailable, falling back to memory: {ex.Message}")
            End Try

            SyncLock _appLock
                _app = app
                _appClientId = clientId
                _appTenantId = tenantId
            End SyncLock
            Return app
        End Function

        Private Function ParseScopes(context As ISharedContext) As String()
            Dim raw = If(context.INI_M365Scopes, "")
            If String.IsNullOrWhiteSpace(raw) Then raw = DefaultScopesIfMissing
            ' MSAL filters out openid/profile/offline_access automatically; pass them anyway.
            Return raw.Split({" "c, ","c, ";"c, ControlChars.Tab},
                             StringSplitOptions.RemoveEmptyEntries) _
                      .Select(Function(s) s.Trim()) _
                      .Where(Function(s) s.Length > 0) _
                      .Distinct(StringComparer.OrdinalIgnoreCase) _
                      .ToArray()
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  SEARCH (entry point)
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Performs a unified Microsoft 365 search across the requested sources.
        ''' Each source is queried independently so that a single 403/throttle
        ''' does not abort the whole run; per-source errors are reported in
        ''' <see cref="M365SearchResult.ErrorsBySource"/>.
        ''' </summary>
        Public Async Function SearchAsync(context As ISharedContext,
                                          query As String,
                                          sources As M365SearchSources,
                                          Optional options As M365SearchOptions = Nothing,
                                          Optional progress As IProgress(Of M365SearchProgress) = Nothing,
                                          Optional ct As CancellationToken = Nothing) As Task(Of M365SearchResult)

            If String.IsNullOrWhiteSpace(query) Then
                Throw New ArgumentException("query must be non-empty", NameOf(query))
            End If
            If sources = M365SearchSources.None Then sources = M365SearchSources.All
            If options Is Nothing Then options = New M365SearchOptions()

            Dim result As New M365SearchResult() With {
                .Query = query,
                .RequestedSources = sources,
                .StartedUtc = DateTime.UtcNow
            }

            Report(progress, M365ProgressStage.SigningIn, M365SearchSources.None, "Signing in to Microsoft 365…", -1, 0)
            Dim token As String
            Try
                token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Report(progress, M365ProgressStage.Cancelled, M365SearchSources.None, "Cancelled", 0, 0)
                Throw
            Catch ex As Exception
                Report(progress, M365ProgressStage.Failed, M365SearchSources.None, "Sign-in failed: " & ex.Message, 0, 0)
                Throw
            End Try

            ' Build the per-source task list.
            Dim sourceList = EnumerateSources(sources).ToList()
            Dim totalSources = sourceList.Count
            Dim completed As Integer = 0
            Dim totalHits As Integer = 0
            Dim sync As New Object()

            Dim runOne As Func(Of M365SearchSources, Task) =
                Async Function(src As M365SearchSources) As Task
                    ct.ThrowIfCancellationRequested()
                    Report(progress, M365ProgressStage.Searching, src,
                           $"Searching {src}…",
                           CInt(100.0 * completed / Math.Max(totalSources, 1)),
                           totalHits)
                    Try
                        Dim hits = Await SearchOneSourceAsync(token, query, src, options, ct).ConfigureAwait(False)
                        SyncLock sync
                            result.Hits.AddRange(hits)
                            result.CountsBySource(src) = hits.Count
                            totalHits += hits.Count
                        End SyncLock
                    Catch ex As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        SyncLock sync
                            result.ErrorsBySource(src) = ex.Message
                            result.CountsBySource(src) = 0
                        End SyncLock
                        Debug.WriteLine($"[M365] Search '{src}' failed: {ex}")
                    Finally
                        Interlocked.Increment(completed)
                        Report(progress, M365ProgressStage.ParsingResults, src,
                               $"{src}: done",
                               CInt(100.0 * completed / Math.Max(totalSources, 1)),
                               totalHits)
                    End Try
                End Function

            Try
                If options.Parallel Then
                    Await Task.WhenAll(sourceList.Select(Function(s) runOne(s))).ConfigureAwait(False)
                Else
                    For Each s In sourceList
                        Await runOne(s).ConfigureAwait(False)
                    Next
                End If
            Catch ex As OperationCanceledException
                Report(progress, M365ProgressStage.Cancelled, M365SearchSources.None, "Cancelled", 0, totalHits)
                Throw
            End Try

            Try
                Await CanonicalizeSearchHitDatesAsync(context, result.Hits, options, ct).ConfigureAwait(False)
            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                Debug.WriteLine("[M365] CanonicalizeSearchHitDatesAsync failed: " & ex.ToString())
            End Try

            result.TotalEstimated = totalHits
            result.FinishedUtc = DateTime.UtcNow
            Report(progress, M365ProgressStage.Completed, M365SearchSources.None,
       $"Completed: {totalHits} hits", 100, totalHits)
            Return result
        End Function

        Private Iterator Function EnumerateSources(sources As M365SearchSources) As IEnumerable(Of M365SearchSources)
            For Each src In {M365SearchSources.Mail, M365SearchSources.OneDrive, M365SearchSources.SharePoint,
                             M365SearchSources.SharePointSites, M365SearchSources.SharePointListItems,
                             M365SearchSources.Teams, M365SearchSources.Calendar,
                             M365SearchSources.OneNote, M365SearchSources.People}
                If (sources And src) = src Then Yield src
            Next
        End Function

        Private Sub Report(progress As IProgress(Of M365SearchProgress),
                           stage As M365ProgressStage,
                           src As M365SearchSources,
                           msg As String, pct As Integer, hits As Integer)
            If progress Is Nothing Then Return
            Try
                progress.Report(New M365SearchProgress() With {
                    .Stage = stage, .Source = src, .Message = msg,
                    .Percent = pct, .HitsSoFar = hits})
            Catch
            End Try
        End Sub

        ' ───────────────────────────────────────────────────────────────────────
        '  Per-source search dispatchers
        ' ───────────────────────────────────────────────────────────────────────

        Private Async Function SearchOneSourceAsync(token As String, query As String,
                                                    src As M365SearchSources,
                                                    options As M365SearchOptions,
                                                    ct As CancellationToken) As Task(Of List(Of M365SearchHit))
            Select Case src
                Case M365SearchSources.OneNote
                    Return Await SearchOneNoteAsync(token, query, options, ct).ConfigureAwait(False)

                Case M365SearchSources.People
                    Return Await SearchPeopleAsync(token, query, options, ct).ConfigureAwait(False)

                Case M365SearchSources.Calendar
                    If ShouldUseCalendarView(options) Then
                        Return Await SearchCalendarViewAsync(token, query, options, ct).ConfigureAwait(False)
                    End If

                    Return Await SearchViaQueryEndpointAsync(token, query, src, options, ct).ConfigureAwait(False)

                Case Else
                    Return Await SearchViaQueryEndpointAsync(token, query, src, options, ct).ConfigureAwait(False)
            End Select
        End Function

        Private Function ShouldUseCalendarView(options As M365SearchOptions) As Boolean
            If options Is Nothing Then Return False
            Return options.From.HasValue OrElse options.To.HasValue
        End Function

        Private Async Function SearchCalendarViewAsync(token As String,
                                                       query As String,
                                                       options As M365SearchOptions,
                                                       ct As CancellationToken) As Task(Of List(Of M365SearchHit))
            Dim effectiveFromDate As Date =
                If(options.From.HasValue,
                   options.From.Value.Date,
                   If(options.To.HasValue, options.To.Value.Date, Date.Today))

            Dim effectiveToDate As Date =
                If(options.To.HasValue,
                   options.To.Value.Date,
                   effectiveFromDate)

            If effectiveToDate < effectiveFromDate Then
                Dim swapDate As Date = effectiveFromDate
                effectiveFromDate = effectiveToDate
                effectiveToDate = swapDate
            End If

            Dim localTimeZone As System.TimeZoneInfo = System.TimeZoneInfo.Local
            Dim startBoundary As DateTimeOffset = BuildDateBoundaryOffset(effectiveFromDate, localTimeZone)
            Dim endBoundary As DateTimeOffset = BuildDateBoundaryOffset(effectiveToDate.AddDays(1), localTimeZone)

            Dim neededMatches As Integer = Math.Max(1, options.FromIndex + options.MaxPerSource)
            Dim pageSize As Integer = Math.Max(10, Math.Min(neededMatches, 100))

            Dim url As String =
                $"{GraphV1}/me/calendarView?" &
                $"startDateTime={Uri.EscapeDataString(FormatGraphDateTimeOffset(startBoundary))}" &
                $"&endDateTime={Uri.EscapeDataString(FormatGraphDateTimeOffset(endBoundary))}" &
                $"&$top={pageSize}" &
                $"&$select=id,subject,bodyPreview,organizer,start,end,isAllDay,location,attendees,webLink"

            Dim matchedHits As New List(Of M365SearchHit)()

            Do While Not String.IsNullOrWhiteSpace(url) AndAlso matchedHits.Count < neededMatches
                ct.ThrowIfCancellationRequested()

                Dim page As JObject

                Using req As New HttpRequestMessage(HttpMethod.Get, url)
                    req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                    req.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                    req.Headers.TryAddWithoutValidation("Prefer", $"outlook.timezone=""{localTimeZone.Id}""")

                    Using resp = Await _http.SendAsync(req, ct).ConfigureAwait(False)
                        Dim body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                        If Not resp.IsSuccessStatusCode Then
                            Throw BuildGraphException(resp.StatusCode, body)
                        End If

                        page = If(String.IsNullOrWhiteSpace(body), New JObject(), JObject.Parse(body))
                    End Using
                End Using

                Dim arr = TryCast(page("value"), JArray)
                If arr IsNot Nothing Then
                    For Each item As JToken In arr
                        ct.ThrowIfCancellationRequested()

                        Dim eventObject As JObject = TryCast(item, JObject)
                        If eventObject Is Nothing Then Continue For
                        If Not CalendarViewMatchesQuery(eventObject, query, options.KqlExtra) Then Continue For

                        Dim wrapped As New JObject From {
                            {"resource", eventObject.DeepClone()}
                        }

                        Dim bodyPreview As String = SafeStr(eventObject, "bodyPreview")
                        If Not String.IsNullOrWhiteSpace(bodyPreview) Then
                            wrapped("summary") = bodyPreview
                        End If

                        Dim hit = ParseSearchHit(wrapped, M365SearchSources.Calendar, options)
                        If hit Is Nothing Then Continue For
                        If Not CalendarHitFallsWithinRequestedDateRange(hit, effectiveFromDate, effectiveToDate) Then Continue For

                        matchedHits.Add(hit)
                        If matchedHits.Count >= neededMatches Then Exit For
                    Next
                End If

                url = SafeStr(page, "@odata.nextLink")
            Loop

            Return matchedHits.
                Skip(Math.Max(0, options.FromIndex)).
                Take(Math.Max(1, options.MaxPerSource)).
                ToList()
        End Function

        Private Function BuildDateBoundaryOffset(dateValue As Date,
                                                 timeZone As System.TimeZoneInfo) As DateTimeOffset
            Dim localMidnight As New DateTime(dateValue.Year, dateValue.Month, dateValue.Day, 0, 0, 0, DateTimeKind.Unspecified)
            Dim offset As TimeSpan = timeZone.GetUtcOffset(localMidnight)
            Return New DateTimeOffset(localMidnight, offset)
        End Function

        Private Function CalendarHitFallsWithinRequestedDateRange(hit As M365SearchHit,
                                                                  fromDate As Date,
                                                                  toDate As Date) As Boolean
            If hit Is Nothing OrElse hit.RawJson Is Nothing Then Return False

            Dim resource As JObject = TryCast(hit.RawJson("resource"), JObject)
            If resource Is Nothing Then Return False

            Dim startUtc As DateTime? = TryParseGraphDateTimeTimeZone(TryCast(resource("start"), JObject))
            If Not startUtc.HasValue Then Return False

            Dim localStart As DateTime = System.TimeZoneInfo.ConvertTimeFromUtc(
                DateTime.SpecifyKind(startUtc.Value, DateTimeKind.Utc),
                System.TimeZoneInfo.Local)

            Dim localDate As Date = localStart.Date
            Return localDate >= fromDate AndAlso localDate <= toDate
        End Function

        Private Function FormatGraphDateTimeOffset(value As DateTimeOffset) As String
            Return value.ToString("yyyy-MM-dd'T'HH:mm:sszzz", Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Function CalendarViewMatchesQuery(eventObject As JObject,
                                                  query As String,
                                                  kqlExtra As String) As Boolean
            Dim combinedQuery As String = ""

            If Not String.IsNullOrWhiteSpace(query) AndAlso query.Trim() <> "*" Then
                combinedQuery = query.Trim()
            End If

            If Not String.IsNullOrWhiteSpace(kqlExtra) Then
                combinedQuery =
                    If(String.IsNullOrWhiteSpace(combinedQuery),
                       kqlExtra.Trim(),
                       combinedQuery & " " & kqlExtra.Trim())
            End If

            Dim terms As List(Of String) = ExtractCalendarQueryTerms(combinedQuery)
            If terms.Count = 0 Then Return True

            Dim searchableText As String = BuildCalendarSearchableText(eventObject)
            If String.IsNullOrWhiteSpace(searchableText) Then Return False

            For Each term As String In terms
                If searchableText.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0 Then
                    Return False
                End If
            Next

            Return True
        End Function

        Private Function ExtractCalendarQueryTerms(rawQuery As String) As List(Of String)
            Dim terms As New List(Of String)()
            Dim q As String = If(rawQuery, "").Trim()
            If String.IsNullOrWhiteSpace(q) OrElse q = "*" Then Return terms

            For Each m As Match In Regex.Matches(q, """[^""]+""|\S+")
                Dim term As String = m.Value.Trim().Trim(""""c, "'"c)
                If String.IsNullOrWhiteSpace(term) Then Continue For

                Dim colonIndex As Integer = term.IndexOf(":"c)
                If colonIndex >= 0 AndAlso colonIndex < term.Length - 1 Then
                    term = term.Substring(colonIndex + 1).Trim().Trim(""""c, "'"c)
                End If

                term = term.Trim("("c, ")"c, "["c, "]"c, "{"c, "}"c, ","c, ";"c)
                If String.IsNullOrWhiteSpace(term) Then Continue For

                Select Case term.ToLowerInvariant()
                    Case "*", "and", "or", "not"
                        Continue For
                End Select

                terms.Add(term)
            Next

            Return terms.
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Function BuildCalendarSearchableText(eventObject As JObject) As String
            If eventObject Is Nothing Then Return ""

            Dim organizer As JObject = TryCast(TryCast(eventObject("organizer"), JObject)?("emailAddress"), JObject)
            Dim location As JObject = TryCast(eventObject("location"), JObject)
            Dim parts As New List(Of String) From {
                SafeStr(eventObject, "subject"),
                SafeStr(eventObject, "bodyPreview"),
                SafeStr(organizer, "name"),
                SafeStr(organizer, "address"),
                SafeStr(location, "displayName"),
                SafeStr(eventObject, "webLink")
            }

            Dim attendees = TryCast(eventObject("attendees"), JArray)
            If attendees IsNot Nothing Then
                For Each attendee As JToken In attendees
                    Dim emailAddress As JObject = TryCast(attendee("emailAddress"), JObject)
                    If emailAddress Is Nothing Then Continue For

                    Dim displayName As String = SafeStr(emailAddress, "name")
                    Dim address As String = SafeStr(emailAddress, "address")

                    If Not String.IsNullOrWhiteSpace(displayName) Then parts.Add(displayName)
                    If Not String.IsNullOrWhiteSpace(address) Then parts.Add(address)
                Next
            End If

            Return String.Join(vbLf, parts.Where(Function(s) Not String.IsNullOrWhiteSpace(s)))
        End Function

        Private Function TryParseGraphDateTimeTimeZone(value As JObject) As DateTime?
            If value Is Nothing Then Return Nothing

            Dim rawDateTime As String = SafeStr(value, "dateTime").Trim()
            If String.IsNullOrWhiteSpace(rawDateTime) Then Return Nothing

            Dim dto As DateTimeOffset
            If DateTimeOffset.TryParse(
                rawDateTime,
                Globalization.CultureInfo.InvariantCulture,
                Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
                dto) Then
                Return dto.UtcDateTime
            End If

            Dim localDateTime As DateTime
            Dim localFormats As String() = {
                "yyyy-MM-dd'T'HH:mm:ss",
                "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd HH:mm",
                "yyyy-MM-dd"
            }

            If DateTime.TryParseExact(
                rawDateTime,
                localFormats,
                Globalization.CultureInfo.InvariantCulture,
                Globalization.DateTimeStyles.None,
                localDateTime) OrElse
               DateTime.TryParse(
                rawDateTime,
                Globalization.CultureInfo.InvariantCulture,
                Globalization.DateTimeStyles.None,
                localDateTime) Then

                Dim tz As System.TimeZoneInfo = ResolveGraphTimeZone(SafeStr(value, "timeZone"))
                If tz IsNot Nothing Then
                    Try
                        localDateTime = DateTime.SpecifyKind(localDateTime, DateTimeKind.Unspecified)
                        Return System.TimeZoneInfo.ConvertTimeToUtc(localDateTime, tz)
                    Catch
                    End Try
                End If
            End If

            Return TryDate(value, "dateTime")
        End Function

        Private Function ResolveGraphTimeZone(timeZoneId As String) As System.TimeZoneInfo
            Dim tzId As String = If(timeZoneId, "").Trim()
            If String.IsNullOrWhiteSpace(tzId) Then Return System.TimeZoneInfo.Utc

            Select Case tzId.ToUpperInvariant()
                Case "UTC", "ETC/UTC", "COORDINATED UNIVERSAL TIME"
                    Return System.TimeZoneInfo.Utc
            End Select

            Try
                Return System.TimeZoneInfo.FindSystemTimeZoneById(tzId)
            Catch
                Return Nothing
            End Try
        End Function

        Private Async Function SearchViaQueryEndpointAsync(token As String, query As String, src As M365SearchSources,
                                                           options As M365SearchOptions,
                                                           ct As CancellationToken) As Task(Of List(Of M365SearchHit))
            Dim entityType = MapEntityType(src)
            Dim kql As String = ApplyKqlFilters(query, src, options)

            Dim requestObject As New JObject From {
                {"entityTypes", New JArray From {entityType}},
                {"query", New JObject From {{"queryString", kql}}},
                {"from", Math.Max(0, options.FromIndex)},
                {"size", Math.Max(1, Math.Min(options.MaxPerSource, 500))}
            }

            If src = M365SearchSources.Mail Then
                requestObject("fields") = New JArray From {
                        "id",
                        "subject",
                        "from",
                        "toRecipients",
                        "ccRecipients",
                        "receivedDateTime",
                        "sentDateTime",
                        "internetMessageId",
                        "conversationId",
                        "webLink",
                        "hasAttachments"
                    }
            ElseIf src = M365SearchSources.OneDrive OrElse
                       src = M365SearchSources.SharePoint Then
                requestObject("fields") = New JArray From {
                        "id",
                        "name",
                        "webUrl",
                        "size",
                        "file",
                        "folder",
                        "createdDateTime",
                        "lastModifiedDateTime",
                        "createdBy",
                        "lastModifiedBy",
                        "parentReference"
                    }
            End If

            Dim req As New JObject From {
                {"requests", New JArray From {requestObject}}
            }

            Dim resp = Await GraphPostAsync(token, GraphV1 & "/search/query", req, ct).ConfigureAwait(False)

            Dim hits As New List(Of M365SearchHit)()
            Dim values = TryCast(resp("value"), JArray)
            If values Is Nothing OrElse values.Count = 0 Then Return hits

            For Each respObj In values
                Dim containers = TryCast(respObj("hitsContainers"), JArray)
                If containers Is Nothing Then Continue For
                For Each c In containers
                    Dim hitArr = TryCast(c("hits"), JArray)
                    If hitArr Is Nothing Then Continue For
                    For Each h In hitArr
                        Dim hit = ParseSearchHit(CType(h, JObject), src, options)
                        If hit IsNot Nothing Then hits.Add(hit)
                    Next
                Next
            Next

            Return hits
        End Function

        Private Function MapEntityType(src As M365SearchSources) As String
            Select Case src
                Case M365SearchSources.Mail : Return "message"
                Case M365SearchSources.OneDrive : Return "driveItem"
                Case M365SearchSources.SharePoint : Return "driveItem"
                Case M365SearchSources.SharePointSites : Return "site"
                Case M365SearchSources.SharePointListItems : Return "listItem"
                Case M365SearchSources.Teams : Return "chatMessage"
                Case M365SearchSources.Calendar : Return "event"
                Case Else : Throw New ArgumentOutOfRangeException(NameOf(src))
            End Select
        End Function

        Private Function ApplyKqlFilters(query As String, src As M365SearchSources, options As M365SearchOptions) As String
            Dim q = If(query, "").Trim()
            If String.IsNullOrWhiteSpace(q) Then q = "*"

            Dim sb As New StringBuilder(q)
            Dim dateField As String = Nothing

            Select Case src
                Case M365SearchSources.Mail
                    ' KQL keyword for mail (received/sent) — culture-independent literal.
                    dateField = "received"

                Case M365SearchSources.OneDrive, M365SearchSources.SharePoint,
                     M365SearchSources.SharePointSites, M365SearchSources.SharePointListItems
                    ' SharePoint Search managed property name (NOT the Graph JSON name
                    ' "lastModifiedDateTime", which KQL doesn't recognise).
                    dateField = "LastModifiedTime"

                Case M365SearchSources.Calendar
                    ' /search/query for event entities does not honour a "start" KQL
                    ' refiner reliably — date bounding is applied client-side after
                    ' canonicalisation. Leave dateField Nothing here.
                    dateField = Nothing
            End Select

            If dateField IsNot Nothing Then
                Dim inv = Globalization.CultureInfo.InvariantCulture

                If options.From.HasValue Then
                    Dim fromDate = options.From.Value.Date
                    sb.Append(" AND ").Append(dateField).Append(">=").
                       Append(fromDate.ToString("yyyy-MM-dd", inv))
                End If

                If options.To.HasValue Then
                    Dim toExclusive = options.To.Value.Date.AddDays(1)
                    sb.Append(" AND ").Append(dateField).Append("<").
                       Append(toExclusive.ToString("yyyy-MM-dd", inv))
                End If
            End If

            If Not String.IsNullOrWhiteSpace(options.KqlExtra) Then
                sb.Append(" AND ").Append(options.KqlExtra.Trim())
            End If

            Return sb.ToString()
        End Function

        Private Async Function CanonicalizeSearchHitDatesAsync(context As ISharedContext,
                                                        hits As IList(Of M365SearchHit),
                                                        options As M365SearchOptions,
                                                        ct As CancellationToken) As Task
            If context Is Nothing OrElse hits Is Nothing OrElse hits.Count = 0 Then Return

            Await CanonicalizeMailSearchHitDatesAsync(context, hits, ct).ConfigureAwait(False)

            If Not ShouldUseCalendarView(options) Then
                Await CanonicalizeCalendarSearchHitDatesAsync(context, hits, ct).ConfigureAwait(False)
            End If

            Await CanonicalizeTeamsSearchHitDatesAsync(context, hits, ct).ConfigureAwait(False)
        End Function

        Private Async Function CanonicalizeMailSearchHitDatesAsync(context As ISharedContext,
                                                           hits As IList(Of M365SearchHit),
                                                           ct As CancellationToken) As Task
            Dim mailHits As List(Of M365SearchHit) =
        hits.
            Where(Function(h) h IsNot Nothing AndAlso
                              h.Source = M365SearchSources.Mail AndAlso
                              Not String.IsNullOrWhiteSpace(h.Id)).
            ToList()

            If mailHits.Count = 0 Then Return

            Dim messages As List(Of M365Message) =
        Await GetMessagesBatchAsync(
            context,
            mailHits.Select(Function(h) h.Id),
            M365MessageFields.Headers,
            ct).ConfigureAwait(False)

            Dim byId As New Dictionary(Of String, M365Message)(StringComparer.OrdinalIgnoreCase)
            For Each msg In messages
                If msg Is Nothing OrElse String.IsNullOrWhiteSpace(msg.Id) Then Continue For
                If Not byId.ContainsKey(msg.Id) Then
                    byId.Add(msg.Id, msg)
                End If
            Next

            For Each hit In mailHits
                ct.ThrowIfCancellationRequested()

                Dim msg As M365Message = Nothing
                If Not byId.TryGetValue(hit.Id, msg) OrElse msg Is Nothing Then Continue For

                Dim preferredDate As DateTime? = If(msg.SentUtc, msg.ReceivedUtc)
                If preferredDate.HasValue Then
                    hit.LastModifiedUtc = preferredDate.Value
                End If

                If Not String.IsNullOrWhiteSpace(msg.Subject) Then
                    hit.Title = msg.Subject
                End If

                If Not String.IsNullOrWhiteSpace(msg.WebLink) Then
                    hit.WebUrl = msg.WebLink
                End If

                Dim resource As JObject = EnsureHitResource(hit)
                resource("id") = msg.Id
                resource("subject") = If(msg.Subject, "")
                resource("webLink") = If(msg.WebLink, "")
                resource("internetMessageId") = If(msg.InternetMessageId, "")
                resource("conversationId") = If(msg.ConversationId, "")

                If msg.SentUtc.HasValue Then
                    resource("sentDateTime") = FormatGraphIsoUtc(msg.SentUtc.Value)
                End If

                If msg.ReceivedUtc.HasValue Then
                    resource("receivedDateTime") = FormatGraphIsoUtc(msg.ReceivedUtc.Value)
                End If
            Next
        End Function

        Private Async Function CanonicalizeCalendarSearchHitDatesAsync(context As ISharedContext,
                                                               hits As IList(Of M365SearchHit),
                                                               ct As CancellationToken) As Task
            Dim eventHits As List(Of M365SearchHit) =
        hits.
            Where(Function(h) h IsNot Nothing AndAlso
                              h.Source = M365SearchSources.Calendar AndAlso
                              Not String.IsNullOrWhiteSpace(h.Id)).
            ToList()

            For Each hit In eventHits
                ct.ThrowIfCancellationRequested()

                Dim ev As M365Event = Nothing
                Try
                    ev = Await GetEventAsync(context, hit.Id, ct).ConfigureAwait(False)
                Catch ex As Exception
                    Debug.WriteLine("[M365] CanonicalizeCalendarSearchHitDatesAsync failed for " & hit.Id & ": " & ex.Message)
                End Try

                If ev Is Nothing Then Continue For

                If ev.StartUtc.HasValue Then
                    hit.LastModifiedUtc = ev.StartUtc.Value
                End If

                If Not String.IsNullOrWhiteSpace(ev.Subject) Then
                    hit.Title = ev.Subject
                End If

                If Not String.IsNullOrWhiteSpace(ev.WebLink) Then
                    hit.WebUrl = ev.WebLink
                End If

                Dim resource As JObject = EnsureHitResource(hit)
                resource("id") = ev.Id
                resource("subject") = If(ev.Subject, "")
                resource("webLink") = If(ev.WebLink, "")

                Dim startObj As JObject = TryCast(resource("start"), JObject)
                If startObj Is Nothing Then
                    startObj = New JObject()
                    resource("start") = startObj
                End If

                Dim endObj As JObject = TryCast(resource("end"), JObject)
                If endObj Is Nothing Then
                    endObj = New JObject()
                    resource("end") = endObj
                End If

                If ev.StartUtc.HasValue Then
                    startObj("dateTime") = FormatGraphIsoUtc(ev.StartUtc.Value)
                    startObj("timeZone") = "UTC"
                End If

                If ev.EndUtc.HasValue Then
                    endObj("dateTime") = FormatGraphIsoUtc(ev.EndUtc.Value)
                    endObj("timeZone") = "UTC"
                End If
            Next
        End Function

        Private Async Function CanonicalizeTeamsSearchHitDatesAsync(context As ISharedContext,
                                                            hits As IList(Of M365SearchHit),
                                                            ct As CancellationToken) As Task
            Dim chatHits As List(Of M365SearchHit) =
        hits.
            Where(Function(h) h IsNot Nothing AndAlso
                              h.Source = M365SearchSources.Teams AndAlso
                              Not String.IsNullOrWhiteSpace(h.Id)).
            ToList()

            For Each hit In chatHits
                ct.ThrowIfCancellationRequested()

                Dim resource As JObject = EnsureHitResource(hit)

                Dim chatId As String = SafeStr(resource, "chatId")
                Dim chanIdent As JObject = TryCast(resource("channelIdentity"), JObject)
                Dim teamId As String = SafeStr(chanIdent, "teamId")
                Dim channelId As String = SafeStr(chanIdent, "channelId")

                Dim cm As M365ChatMessage = Nothing

                Try
                    If Not String.IsNullOrWhiteSpace(teamId) AndAlso Not String.IsNullOrWhiteSpace(channelId) Then
                        cm = Await GetChatMessageAsync(
                    context,
                    hit.Id,
                    Nothing,
                    teamId,
                    channelId,
                    ct).ConfigureAwait(False)
                    ElseIf Not String.IsNullOrWhiteSpace(chatId) Then
                        cm = Await GetChatMessageAsync(
                    context,
                    hit.Id,
                    chatId,
                    Nothing,
                    Nothing,
                    ct).ConfigureAwait(False)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("[M365] CanonicalizeTeamsSearchHitDatesAsync failed for " & hit.Id & ": " & ex.Message)
                End Try

                If cm Is Nothing Then Continue For

                If cm.CreatedUtc.HasValue Then
                    hit.LastModifiedUtc = cm.CreatedUtc.Value
                    resource("createdDateTime") = FormatGraphIsoUtc(cm.CreatedUtc.Value)
                End If

                If Not String.IsNullOrWhiteSpace(cm.WebUrl) Then
                    hit.WebUrl = cm.WebUrl
                    resource("webUrl") = cm.WebUrl
                End If
            Next
        End Function

        Private Function EnsureHitResource(hit As M365SearchHit) As JObject
            If hit Is Nothing Then Return New JObject()

            If hit.RawJson Is Nothing Then
                hit.RawJson = New JObject()
            End If

            Dim resource As JObject = TryCast(hit.RawJson("resource"), JObject)
            If resource Is Nothing Then
                resource = New JObject()
                hit.RawJson("resource") = resource
            End If

            Return resource
        End Function

        Private Function FormatGraphIsoUtc(value As DateTime) As String
            Return value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", Globalization.CultureInfo.InvariantCulture)
        End Function


        Private Function ParseSearchHit(h As JObject, src As M365SearchSources,
                                        options As M365SearchOptions) As M365SearchHit
            If h Is Nothing Then Return Nothing
            Dim resource = TryCast(h("resource"), JObject)
            Dim hit As New M365SearchHit() With {
                .Source = src,
                .RawJson = h,
                .Summary = If(options.CleanSummaries, CleanSummary(SafeStr(h, "summary")), SafeStr(h, "summary"))
            }

            If resource Is Nothing Then Return hit

            Select Case src
                Case M365SearchSources.Mail
                    hit.Id = SafeStr(resource, "id")
                    If String.IsNullOrWhiteSpace(hit.Id) Then
                        hit.Id = SafeStr(h, "hitId")
                    End If

                    hit.Title = SafeStr(resource, "subject")
                    hit.WebUrl = SafeStr(resource, "webLink")

                    Dim sentUtc = TryDate(resource, "sentDateTime")
                    Dim receivedUtc = TryDate(resource, "receivedDateTime")
                    hit.LastModifiedUtc = If(sentUtc, receivedUtc)

                    Dim fromObj = TryCast(resource("from"), JObject)?("emailAddress")
                    hit.Author = SafeStr(CType(fromObj, JObject), "name")
                    If String.IsNullOrWhiteSpace(hit.Author) Then
                        hit.Author = SafeStr(CType(fromObj, JObject), "address")
                    End If

                    hit.AdditionalText = BuildRecipientsDisplay(TryCast(resource("toRecipients"), JArray))
                Case M365SearchSources.OneDrive, M365SearchSources.SharePoint
                    hit.Id = SafeStr(resource, "id")
                    Dim parent = TryCast(resource("parentReference"), JObject)
                    hit.ParentId = SafeStr(parent, "driveId")
                    hit.Title = SafeStr(resource, "name")
                    hit.WebUrl = SafeStr(resource, "webUrl")
                    hit.LastModifiedUtc = TryDate(resource, "lastModifiedDateTime")
                    Dim mod_ = TryCast(resource("lastModifiedBy"), JObject)?("user")
                    hit.Author = SafeStr(CType(mod_, JObject), "displayName")
                Case M365SearchSources.SharePointSites
                    hit.Id = SafeStr(resource, "id")
                    hit.Title = If(SafeStr(resource, "displayName"), SafeStr(resource, "name"))
                    hit.WebUrl = SafeStr(resource, "webUrl")
                Case M365SearchSources.SharePointListItems
                    hit.Id = SafeStr(resource, "id")
                    Dim pr = TryCast(resource("parentReference"), JObject)
                    hit.ParentId = SafeStr(pr, "siteId")
                    hit.Title = SafeStr(resource, "name")
                    hit.WebUrl = SafeStr(resource, "webUrl")
                    hit.LastModifiedUtc = TryDate(resource, "lastModifiedDateTime")
                Case M365SearchSources.Teams
                    hit.Id = SafeStr(resource, "id")
                    Dim chatId = SafeStr(resource, "chatId")
                    Dim chanId = SafeStr(TryCast(resource("channelIdentity"), JObject), "channelId")
                    hit.ParentId = If(String.IsNullOrEmpty(chatId), chanId, chatId)
                    hit.Title = TruncatePlain(StripHtml(SafeStr(TryCast(resource("body"), JObject), "content")), 80)
                    hit.WebUrl = SafeStr(resource, "webUrl")
                    hit.LastModifiedUtc = TryDate(resource, "createdDateTime")
                    Dim fromUser = TryCast(TryCast(resource("from"), JObject)?("user"), JObject)
                    hit.Author = SafeStr(fromUser, "displayName")
                Case M365SearchSources.Calendar
                    hit.Id = SafeStr(resource, "id")
                    hit.Title = SafeStr(resource, "subject")
                    hit.WebUrl = SafeStr(resource, "webLink")
                    hit.LastModifiedUtc = TryParseGraphDateTimeTimeZone(TryCast(resource("start"), JObject))
                    Dim org = TryCast(TryCast(resource("organizer"), JObject)?("emailAddress"), JObject)
                    hit.Author = SafeStr(org, "name")
            End Select
            Return hit
        End Function


        Private Function BuildRecipientsDisplay(arr As JArray) As String
            If arr Is Nothing OrElse arr.Count = 0 Then Return ""

            Dim parts As New List(Of String)()

            For Each r In arr
                Dim ea = TryCast(CType(r, JObject)("emailAddress"), JObject)
                If ea Is Nothing Then Continue For

                Dim addr = SafeStr(ea, "address")
                Dim disp = SafeStr(ea, "name")

                If String.IsNullOrWhiteSpace(addr) AndAlso String.IsNullOrWhiteSpace(disp) Then Continue For
                parts.Add(If(String.IsNullOrWhiteSpace(disp), addr, disp))
            Next

            Return String.Join("; ", parts.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Distinct())
        End Function


        ' ── OneNote search (separate endpoint) ────────────────────────────────
        Private Async Function SearchOneNoteAsync(token As String, query As String,
                                                  options As M365SearchOptions,
                                                  ct As CancellationToken) As Task(Of List(Of M365SearchHit))
            Dim url = $"{GraphV1}/me/onenote/pages?$search=""{Uri.EscapeDataString(query)}""&$top={options.MaxPerSource}"
            Dim resp = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Dim hits As New List(Of M365SearchHit)()
            Dim arr = TryCast(resp("value"), JArray)
            If arr Is Nothing Then Return hits
            For Each p In arr
                Dim po = CType(p, JObject)
                hits.Add(New M365SearchHit() With {
                    .Source = M365SearchSources.OneNote,
                    .Id = SafeStr(po, "id"),
                    .Title = SafeStr(po, "title"),
                    .WebUrl = SafeStr(TryCast(TryCast(po("links"), JObject)?("oneNoteWebUrl"), JObject), "href"),
                    .LastModifiedUtc = TryDate(po, "lastModifiedDateTime"),
                    .Summary = "",
                    .RawJson = po
                })
            Next
            Return hits
        End Function

        ' ── People search ─────────────────────────────────────────────────────
        Private Async Function SearchPeopleAsync(token As String, query As String,
                                                 options As M365SearchOptions,
                                                 ct As CancellationToken) As Task(Of List(Of M365SearchHit))
            Dim url = $"{GraphV1}/me/people?$search=""{Uri.EscapeDataString(query)}""&$top={options.MaxPerSource}"
            Dim resp = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Dim hits As New List(Of M365SearchHit)()
            Dim arr = TryCast(resp("value"), JArray)
            If arr Is Nothing Then Return hits
            For Each p In arr
                Dim po = CType(p, JObject)
                Dim email As String = ""
                Dim addrs = TryCast(po("scoredEmailAddresses"), JArray)
                If addrs IsNot Nothing AndAlso addrs.Count > 0 Then
                    email = SafeStr(CType(addrs(0), JObject), "address")
                End If
                hits.Add(New M365SearchHit() With {
                    .Source = M365SearchSources.People,
                    .Id = SafeStr(po, "id"),
                    .Title = SafeStr(po, "displayName"),
                    .Author = email,
                    .Summary = SafeStr(po, "jobTitle"),
                    .RawJson = po
                })
            Next
            Return hits
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  RETRIEVAL — MAIL
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>Retrieves a single Outlook message with the requested fields.</summary>
        Public Async Function GetMessageAsync(context As ISharedContext,
                                              messageId As String,
                                              Optional fields As M365MessageFields = M365MessageFields.Body Or M365MessageFields.Recipients,
                                              Optional ct As CancellationToken = Nothing) As Task(Of M365Message)
            Dim normalizedMessageId = ToGraphUrlSafeId(messageId)
            If String.IsNullOrWhiteSpace(normalizedMessageId) Then Return Nothing

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim selectFields = New List(Of String) From {
                "id", "subject", "from", "receivedDateTime", "sentDateTime", "importance",
                "hasAttachments", "bodyPreview", "internetMessageId", "conversationId", "webLink"
            }
            If (fields And M365MessageFields.Body) <> 0 Then selectFields.Add("body")
            If (fields And M365MessageFields.Recipients) <> 0 Then
                selectFields.Add("toRecipients") : selectFields.Add("ccRecipients") : selectFields.Add("bccRecipients")
            End If
            If (fields And M365MessageFields.Categories) <> 0 Then selectFields.Add("categories")
            If (fields And M365MessageFields.InternetHeaders) <> 0 Then selectFields.Add("internetMessageHeaders")

            Dim url = $"{GraphV1}/me/messages/{Uri.EscapeDataString(normalizedMessageId)}?$select={String.Join(",", selectFields)}"
            If (fields And M365MessageFields.AttachmentsList) <> 0 Then url &= "&$expand=attachments($select=id,name,contentType,size,isInline)"

            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Return ParseMessage(j, fields)
        End Function

        ''' <summary>
        ''' Resolves a message via /me/messages?$filter=internetMessageId eq '...'.
        ''' Returns Nothing when no match is found.
        ''' </summary>
        Public Async Function GetMessageByInternetMessageIdAsync(context As ISharedContext,
                                                                 internetMessageId As String,
                                                                 Optional fields As M365MessageFields = M365MessageFields.Body Or M365MessageFields.Recipients,
                                                                 Optional ct As CancellationToken = Nothing) As Task(Of M365Message)
            If String.IsNullOrWhiteSpace(internetMessageId) Then Return Nothing

            Dim raw As String = internetMessageId.Trim()
            Dim core As String = raw.Trim("<"c, ">"c)

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)

            Dim selectFields = New List(Of String) From {
                "id", "subject", "from", "receivedDateTime", "sentDateTime", "importance",
                "hasAttachments", "bodyPreview", "internetMessageId", "conversationId", "webLink"
            }
            If (fields And M365MessageFields.Body) <> 0 Then selectFields.Add("body")
            If (fields And M365MessageFields.Recipients) <> 0 Then
                selectFields.Add("toRecipients") : selectFields.Add("ccRecipients") : selectFields.Add("bccRecipients")
            End If
            If (fields And M365MessageFields.Categories) <> 0 Then selectFields.Add("categories")
            If (fields And M365MessageFields.InternetHeaders) <> 0 Then selectFields.Add("internetMessageHeaders")

            Dim candidates As New List(Of String) From {
                raw,
                core,
                "<" & core & ">"
            }

            For Each candidate In candidates.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Distinct()
                Dim filter As String = "internetMessageId eq '" & candidate.Replace("'", "''") & "'"
                Dim url As String = $"{GraphV1}/me/messages?$top=1&$select={String.Join(",", selectFields)}&$filter={Uri.EscapeDataString(filter)}"

                Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
                If j Is Nothing Then Continue For

                Dim arr = TryCast(j("value"), JArray)
                If arr Is Nothing OrElse arr.Count = 0 Then Continue For

                Dim first = TryCast(arr(0), JObject)
                If first Is Nothing Then Continue For

                Return ParseMessage(first, fields)
            Next

            Return Nothing
        End Function


        ''' <summary>Retrieves multiple messages in a single Graph $batch request (up to 20 per batch).</summary>
        Public Async Function GetMessagesBatchAsync(context As ISharedContext,
                                                    messageIds As IEnumerable(Of String),
                                                    Optional fields As M365MessageFields = M365MessageFields.Body Or M365MessageFields.Recipients,
                                                    Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365Message))
            Dim ids = messageIds?.
                Select(Function(s) ToGraphUrlSafeId(s)).
                Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim out As New List(Of M365Message)()
            If ids Is Nothing OrElse ids.Count = 0 Then Return out

            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim selectFields = New List(Of String) From {
                "id", "subject", "from", "receivedDateTime", "sentDateTime", "importance",
                "hasAttachments", "bodyPreview", "internetMessageId", "conversationId", "webLink"
            }
            If (fields And M365MessageFields.Body) <> 0 Then selectFields.Add("body")
            If (fields And M365MessageFields.Recipients) <> 0 Then
                selectFields.Add("toRecipients") : selectFields.Add("ccRecipients") : selectFields.Add("bccRecipients")
            End If

            ' Graph $batch caps at 20 requests per call.
            For chunkStart As Integer = 0 To ids.Count - 1 Step 20
                Dim chunk = ids.Skip(chunkStart).Take(20).ToList()
                Dim batchReqs As New JArray()
                For i As Integer = 0 To chunk.Count - 1
                    Dim u = $"/me/messages/{Uri.EscapeDataString(chunk(i))}?$select={String.Join(",", selectFields)}"
                    batchReqs.Add(New JObject From {
                        {"id", (i + 1).ToString()},
                        {"method", "GET"},
                        {"url", u}
                    })
                Next
                Dim batchBody = New JObject From {{"requests", batchReqs}}
                Dim resp = Await GraphPostAsync(token, GraphV1 & "/$batch", batchBody, ct).ConfigureAwait(False)
                Dim responses = TryCast(resp("responses"), JArray)
                If responses Is Nothing Then Continue For
                For Each r In responses
                    Dim status = CInt(If(r("status"), 0))
                    If status >= 200 AndAlso status < 300 Then
                        Dim body = TryCast(r("body"), JObject)
                        If body IsNot Nothing Then out.Add(ParseMessage(body, fields))
                    End If
                Next
            Next
            Return out
        End Function

        ''' <summary>Returns just the body of a message as plain text (HTML stripped if needed).</summary>
        Public Async Function GetMessageBodyAsync(context As ISharedContext,
                                                  messageId As String,
                                                  Optional asPlainText As Boolean = True,
                                                  Optional ct As CancellationToken = Nothing) As Task(Of String)
            Dim m = Await GetMessageAsync(context, messageId, M365MessageFields.Body, ct).ConfigureAwait(False)
            If m Is Nothing OrElse String.IsNullOrEmpty(m.Body) Then Return ""
            If asPlainText AndAlso String.Equals(m.BodyContentType, "html", StringComparison.OrdinalIgnoreCase) Then
                Return StripHtml(m.Body)
            End If
            Return m.Body
        End Function

        ''' <summary>Lists attachments for a message (metadata only — no content).</summary>
        Public Async Function ListAttachmentsAsync(context As ISharedContext,
                                                   messageId As String,
                                                   Optional ct As CancellationToken = Nothing) As Task(Of List(Of M365AttachmentInfo))
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = $"{GraphV1}/me/messages/{Uri.EscapeDataString(messageId)}/attachments?$select=id,name,contentType,size,isInline"
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Dim out As New List(Of M365AttachmentInfo)()
            Dim arr = TryCast(j("value"), JArray)
            If arr Is Nothing Then Return out
            For Each a In arr
                Dim ao = CType(a, JObject)
                out.Add(New M365AttachmentInfo() With {
                    .Id = SafeStr(ao, "id"),
                    .Name = SafeStr(ao, "name"),
                    .ContentType = SafeStr(ao, "contentType"),
                    .Size = CLng(If(ao("size"), 0L)),
                    .IsInline = CBool(If(ao("isInline"), False)),
                    .OdataType = SafeStr(ao, "@odata.type")
                })
            Next
            Return out
        End Function

        ''' <summary>Downloads a single attachment (fileAttachment) to <paramref name="targetPath"/>.</summary>
        Public Async Function DownloadAttachmentAsync(context As ISharedContext,
                                                      messageId As String,
                                                      attachmentId As String,
                                                      targetPath As String,
                                                      Optional ct As CancellationToken = Nothing) As Task(Of String)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            ' /$value returns the raw bytes for fileAttachment.
            Dim url = $"{GraphV1}/me/messages/{Uri.EscapeDataString(messageId)}/attachments/{Uri.EscapeDataString(attachmentId)}/$value"
            Using req As New HttpRequestMessage(HttpMethod.Get, url)
                req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                Using resp = Await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                    If Not resp.IsSuccessStatusCode Then
                        Await ThrowGraphErrorAsync(resp).ConfigureAwait(False)
                    End If
                    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
                    Using fs As New FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await resp.Content.CopyToAsync(fs).ConfigureAwait(False)
                    End Using
                End Using
            End Using
            Return targetPath
        End Function

        ''' <summary>Downloads every (non-inline) attachment of a message into <paramref name="targetFolder"/>.</summary>
        Public Async Function DownloadAllAttachmentsAsync(context As ISharedContext,
                                                          messageId As String,
                                                          targetFolder As String,
                                                          Optional includeInline As Boolean = False,
                                                          Optional ct As CancellationToken = Nothing) As Task(Of List(Of String))
            Dim out As New List(Of String)()
            Dim atts = Await ListAttachmentsAsync(context, messageId, ct).ConfigureAwait(False)
            Directory.CreateDirectory(targetFolder)
            For Each a In atts
                ct.ThrowIfCancellationRequested()
                If a.IsInline AndAlso Not includeInline Then Continue For
                If Not String.Equals(a.OdataType, "#microsoft.graph.fileAttachment", StringComparison.OrdinalIgnoreCase) Then
                    Continue For ' skip itemAttachment / referenceAttachment in this helper
                End If
                Dim safeName = MakeSafeFileName(a.Name, "attachment.bin")
                Dim localPath = System.IO.Path.Combine(targetFolder, safeName)
                Dim n As Integer = 1
                While File.Exists(localPath)
                    localPath = System.IO.Path.Combine(targetFolder,
                        System.IO.Path.GetFileNameWithoutExtension(safeName) & $"_{n}" & System.IO.Path.GetExtension(safeName))
                    n += 1
                End While
                Try
                    Await DownloadAttachmentAsync(context, messageId, a.Id, localPath, ct).ConfigureAwait(False)
                    out.Add(localPath)
                Catch ex As Exception
                    Debug.WriteLine($"[M365] DownloadAttachment '{a.Name}' failed: {ex.Message}")
                End Try
            Next
            Return out
        End Function

        Private Function ParseMessage(j As JObject, fields As M365MessageFields) As M365Message
            If j Is Nothing Then Return Nothing
            Dim m As New M365Message() With {.RawJson = j}
            m.Id = SafeStr(j, "id")
            m.Subject = SafeStr(j, "subject")
            m.Importance = SafeStr(j, "importance")
            m.HasAttachments = CBool(If(j("hasAttachments"), False))
            m.BodyPreview = SafeStr(j, "bodyPreview")
            m.InternetMessageId = SafeStr(j, "internetMessageId")
            m.ConversationId = SafeStr(j, "conversationId")
            m.WebLink = SafeStr(j, "webLink")

            m.ReceivedUtc = TryDate(j, "receivedDateTime")
            m.SentUtc = TryDate(j, "sentDateTime")

            Dim fromObj = TryCast(j("from"), JObject)?("emailAddress")
            If fromObj IsNot Nothing Then
                m.From = SafeStr(CType(fromObj, JObject), "name")
                m.FromAddress = SafeStr(CType(fromObj, JObject), "address")
            End If

            Dim body = TryCast(j("body"), JObject)
            If body IsNot Nothing Then
                m.BodyContentType = SafeStr(body, "contentType")
                m.Body = SafeStr(body, "content")
            End If

            ParseRecipients(j, "toRecipients", m.To_)
            ParseRecipients(j, "ccRecipients", m.Cc)
            ParseRecipients(j, "bccRecipients", m.Bcc)

            Dim cats = TryCast(j("categories"), JArray)
            If cats IsNot Nothing Then m.Categories = cats.Select(Function(x) x.ToString()).ToList()

            Dim atts = TryCast(j("attachments"), JArray)
            If atts IsNot Nothing Then
                For Each a In atts
                    Dim ao = CType(a, JObject)
                    m.Attachments.Add(New M365AttachmentInfo() With {
                        .Id = SafeStr(ao, "id"),
                        .Name = SafeStr(ao, "name"),
                        .ContentType = SafeStr(ao, "contentType"),
                        .Size = CLng(If(ao("size"), 0L)),
                        .IsInline = CBool(If(ao("isInline"), False)),
                        .OdataType = SafeStr(ao, "@odata.type")
                    })
                Next
            End If
            Return m
        End Function

        Private Sub ParseRecipients(j As JObject, name As String, sink As List(Of String))
            Dim arr = TryCast(j(name), JArray)
            If arr Is Nothing Then Return
            For Each r In arr
                Dim ea = TryCast(CType(r, JObject)("emailAddress"), JObject)
                Dim addr = SafeStr(ea, "address")
                Dim disp = SafeStr(ea, "name")
                If String.IsNullOrEmpty(addr) Then Continue For
                sink.Add(If(String.IsNullOrEmpty(disp), addr, $"{disp} <{addr}>"))
            Next
        End Sub

        ' ═══════════════════════════════════════════════════════════════════════
        '  RETRIEVAL — FILES (DriveItem)
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>Fetches DriveItem metadata. driveId may be empty for items returned by personal /me/drive search.</summary>
        Public Async Function GetDriveItemAsync(context As ISharedContext,
                                                driveItemId As String,
                                                Optional driveId As String = Nothing,
                                                Optional ct As CancellationToken = Nothing) As Task(Of M365DriveItem)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url As String
            If String.IsNullOrEmpty(driveId) Then
                url = $"{GraphV1}/me/drive/items/{Uri.EscapeDataString(driveItemId)}"
            Else
                url = $"{GraphV1}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(driveItemId)}"
            End If
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            Return ParseDriveItem(j)
        End Function

        ''' <summary>Streams a file's bytes to disk via Graph's redirected download URL.</summary>
        Public Async Function DownloadFileAsync(context As ISharedContext,
                                                driveItemId As String,
                                                targetPath As String,
                                                Optional driveId As String = Nothing,
                                                Optional ct As CancellationToken = Nothing) As Task(Of String)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url As String
            If String.IsNullOrEmpty(driveId) Then
                url = $"{GraphV1}/me/drive/items/{Uri.EscapeDataString(driveItemId)}/content"
            Else
                url = $"{GraphV1}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(driveItemId)}/content"
            End If
            Using req As New HttpRequestMessage(HttpMethod.Get, url)
                req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                Using resp = Await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(False)
                    If Not resp.IsSuccessStatusCode Then Await ThrowGraphErrorAsync(resp).ConfigureAwait(False)
                    Directory.CreateDirectory(Path.GetDirectoryName(targetPath))
                    Using fs As New FileStream(targetPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await resp.Content.CopyToAsync(fs).ConfigureAwait(False)
                    End Using
                End Using
            End Using
            Return targetPath
        End Function

        Private Function ParseDriveItem(j As JObject) As M365DriveItem
            If j Is Nothing Then Return Nothing
            Dim it As New M365DriveItem() With {.RawJson = j}
            it.Id = SafeStr(j, "id")
            it.Name = SafeStr(j, "name")
            it.Size = CLng(If(j("size"), 0L))
            it.WebUrl = SafeStr(j, "webUrl")
            it.ETag = SafeStr(j, "eTag")
            it.IsFolder = j("folder") IsNot Nothing
            Dim file_ = TryCast(j("file"), JObject)
            If file_ IsNot Nothing Then it.MimeType = SafeStr(file_, "mimeType")
            Dim parent = TryCast(j("parentReference"), JObject)
            If parent IsNot Nothing Then
                it.DriveId = SafeStr(parent, "driveId")
                it.Path = SafeStr(parent, "path")
            End If
            it.CreatedUtc = TryDate(j, "createdDateTime")
            it.LastModifiedUtc = TryDate(j, "lastModifiedDateTime")
            Dim cb = TryCast(TryCast(j("createdBy"), JObject)?("user"), JObject)
            If cb IsNot Nothing Then it.CreatedBy = SafeStr(cb, "displayName")
            Dim mb = TryCast(TryCast(j("lastModifiedBy"), JObject)?("user"), JObject)
            If mb IsNot Nothing Then it.LastModifiedBy = SafeStr(mb, "displayName")
            Return it
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  RETRIEVAL — CALENDAR / TEAMS / ONENOTE / PEOPLE
        ' ═══════════════════════════════════════════════════════════════════════

        Public Async Function GetEventAsync(context As ISharedContext,
                                            eventId As String,
                                            Optional ct As CancellationToken = Nothing) As Task(Of M365Event)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = $"{GraphV1}/me/events/{Uri.EscapeDataString(eventId)}"
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            If j Is Nothing Then Return Nothing
            Dim ev As New M365Event() With {.RawJson = j}
            ev.Id = SafeStr(j, "id")
            ev.Subject = SafeStr(j, "subject")
            ev.BodyPreview = SafeStr(j, "bodyPreview")
            ev.WebLink = SafeStr(j, "webLink")
            ev.IsAllDay = CBool(If(j("isAllDay"), False))
            ev.StartUtc = TryParseGraphDateTimeTimeZone(TryCast(j("start"), JObject))
            ev.EndUtc = TryParseGraphDateTimeTimeZone(TryCast(j("end"), JObject))
            Dim org = TryCast(TryCast(j("organizer"), JObject)?("emailAddress"), JObject)
            If org IsNot Nothing Then ev.Organizer = SafeStr(org, "name")
            Dim loc = TryCast(j("location"), JObject)
            If loc IsNot Nothing Then ev.Location = SafeStr(loc, "displayName")
            Dim att = TryCast(j("attendees"), JArray)
            If att IsNot Nothing Then
                For Each a In att
                    Dim ea = TryCast(TryCast(a("emailAddress"), JObject), JObject)
                    Dim disp = SafeStr(ea, "name")
                    Dim addr = SafeStr(ea, "address")
                    ev.Attendees.Add(If(String.IsNullOrEmpty(disp), addr, $"{disp} <{addr}>"))
                Next
            End If
            Return ev
        End Function

        ''' <summary>
        ''' Retrieves a single Teams chat message. Use the <c>chatId</c> from the search hit's
        ''' <see cref="M365SearchHit.ParentId"/> for 1:1 / group chats, or pass team+channel ids for channel messages.
        ''' </summary>
        Public Async Function GetChatMessageAsync(context As ISharedContext,
                                                  messageId As String,
                                                  Optional chatId As String = Nothing,
                                                  Optional teamId As String = Nothing,
                                                  Optional channelId As String = Nothing,
                                                  Optional ct As CancellationToken = Nothing) As Task(Of M365ChatMessage)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url As String
            If Not String.IsNullOrEmpty(teamId) AndAlso Not String.IsNullOrEmpty(channelId) Then
                url = $"{GraphV1}/teams/{Uri.EscapeDataString(teamId)}/channels/{Uri.EscapeDataString(channelId)}/messages/{Uri.EscapeDataString(messageId)}"
            ElseIf Not String.IsNullOrEmpty(chatId) Then
                url = $"{GraphV1}/chats/{Uri.EscapeDataString(chatId)}/messages/{Uri.EscapeDataString(messageId)}"
            Else
                Throw New ArgumentException("Either chatId or (teamId + channelId) must be provided.")
            End If
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            If j Is Nothing Then Return Nothing
            Dim cm As New M365ChatMessage() With {
                .RawJson = j,
                .Id = SafeStr(j, "id"),
                .ChatId = chatId,
                .TeamId = teamId,
                .ChannelId = channelId,
                .CreatedUtc = TryDate(j, "createdDateTime"),
                .WebUrl = SafeStr(j, "webUrl")
            }
            Dim body = TryCast(j("body"), JObject)
            If body IsNot Nothing Then
                cm.ContentType = SafeStr(body, "contentType")
                cm.Content = SafeStr(body, "content")
            End If
            Dim fromUser = TryCast(TryCast(j("from"), JObject)?("user"), JObject)
            If fromUser IsNot Nothing Then cm.From = SafeStr(fromUser, "displayName")
            Return cm
        End Function

        ''' <summary>Retrieves a OneNote page. When <paramref name="includeContent"/> is True, also fetches HTML body.</summary>
        Public Async Function GetOneNotePageAsync(context As ISharedContext,
                                                  pageId As String,
                                                  Optional includeContent As Boolean = False,
                                                  Optional ct As CancellationToken = Nothing) As Task(Of M365OneNotePage)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim metaUrl = $"{GraphV1}/me/onenote/pages/{Uri.EscapeDataString(pageId)}"
            Dim j = Await GraphGetAsync(token, metaUrl, ct).ConfigureAwait(False)
            If j Is Nothing Then Return Nothing
            Dim p As New M365OneNotePage() With {
                .RawJson = j,
                .Id = SafeStr(j, "id"),
                .Title = SafeStr(j, "title"),
                .CreatedUtc = TryDate(j, "createdDateTime"),
                .LastModifiedUtc = TryDate(j, "lastModifiedDateTime"),
                .ContentUrl = SafeStr(j, "contentUrl")
            }
            If includeContent AndAlso Not String.IsNullOrEmpty(p.ContentUrl) Then
                Using req As New HttpRequestMessage(HttpMethod.Get, p.ContentUrl)
                    req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                    Using resp = Await _http.SendAsync(req, ct).ConfigureAwait(False)
                        If resp.IsSuccessStatusCode Then
                            p.HtmlContent = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                        End If
                    End Using
                End Using
            End If
            Return p
        End Function

        Public Async Function GetPersonAsync(context As ISharedContext,
                                             personId As String,
                                             Optional ct As CancellationToken = Nothing) As Task(Of M365Person)
            Dim token = Await GetAccessTokenAsync(context, ct).ConfigureAwait(False)
            Dim url = $"{GraphV1}/me/people/{Uri.EscapeDataString(personId)}"
            Dim j = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
            If j Is Nothing Then Return Nothing
            Dim per As New M365Person() With {
                .RawJson = j,
                .Id = SafeStr(j, "id"),
                .DisplayName = SafeStr(j, "displayName"),
                .JobTitle = SafeStr(j, "jobTitle"),
                .CompanyName = SafeStr(j, "companyName")
            }
            Dim emails = TryCast(j("scoredEmailAddresses"), JArray)
            If emails IsNot Nothing Then
                For Each e In emails
                    Dim addr = SafeStr(CType(e, JObject), "address")
                    If Not String.IsNullOrEmpty(addr) Then per.EmailAddresses.Add(addr)
                Next
            End If
            Dim phones = TryCast(j("phones"), JArray)
            If phones IsNot Nothing Then
                For Each ph In phones
                    Dim n = SafeStr(CType(ph, JObject), "number")
                    If Not String.IsNullOrEmpty(n) Then per.Phones.Add(n)
                Next
            End If
            Return per
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  LOW-LEVEL GRAPH HELPERS (Friend — for power-user code paths)
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>Issues an authenticated GET against any Graph URL and returns the parsed JSON.</summary>
        Friend Async Function GraphGetAsync(token As String, url As String,
                                            Optional ct As CancellationToken = Nothing) As Task(Of JObject)
            Using req As New HttpRequestMessage(HttpMethod.Get, url)
                req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                req.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                Using resp = Await _http.SendAsync(req, ct).ConfigureAwait(False)
                    Dim body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    If Not resp.IsSuccessStatusCode Then
                        Throw BuildGraphException(resp.StatusCode, body)
                    End If
                    If String.IsNullOrWhiteSpace(body) Then Return New JObject()
                    Return JObject.Parse(body)
                End Using
            End Using
        End Function

        ''' <summary>Issues an authenticated POST with JSON body and returns the parsed response JSON.</summary>
        Friend Async Function GraphPostAsync(token As String, url As String, body As JObject,
                                             Optional ct As CancellationToken = Nothing) As Task(Of JObject)
            Using req As New HttpRequestMessage(HttpMethod.Post, url)
                req.Headers.Authorization = New AuthenticationHeaderValue("Bearer", token)
                req.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                req.Content = New StringContent(body.ToString(), Encoding.UTF8, "application/json")
                Using resp = Await _http.SendAsync(req, ct).ConfigureAwait(False)
                    Dim text = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                    If Not resp.IsSuccessStatusCode Then
                        Throw BuildGraphException(resp.StatusCode, text)
                    End If
                    If String.IsNullOrWhiteSpace(text) Then Return New JObject()
                    Return JObject.Parse(text)
                End Using
            End Using
        End Function

        ''' <summary>Walks a paginated Graph collection (followed via @odata.nextLink), yielding all values.</summary>
        Friend Async Function GraphPageAsync(token As String, startUrl As String,
                                             Optional maxItems As Integer = 1000,
                                             Optional ct As CancellationToken = Nothing) As Task(Of List(Of JObject))
            Dim out As New List(Of JObject)()
            Dim url = startUrl
            While Not String.IsNullOrEmpty(url) AndAlso out.Count < maxItems
                ct.ThrowIfCancellationRequested()
                Dim page = Await GraphGetAsync(token, url, ct).ConfigureAwait(False)
                Dim arr = TryCast(page("value"), JArray)
                If arr IsNot Nothing Then
                    For Each v In arr
                        If out.Count >= maxItems Then Exit For
                        Dim jo = TryCast(v, JObject)
                        If jo IsNot Nothing Then out.Add(jo)
                    Next
                End If
                url = SafeStr(page, "@odata.nextLink")
            End While
            Return out
        End Function

        Private Async Function ThrowGraphErrorAsync(resp As HttpResponseMessage) As Task
            Dim body = ""
            Try
                body = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
            Catch
            End Try
            Throw BuildGraphException(resp.StatusCode, body)
        End Function

        Private Function BuildGraphException(status As HttpStatusCode, body As String) As M365GraphException
            Dim msg As String = $"Graph error {CInt(status)} {status}"
            Try
                If Not String.IsNullOrWhiteSpace(body) Then
                    Dim j = JObject.Parse(body)
                    Dim err = TryCast(j("error"), JObject)
                    If err IsNot Nothing Then
                        msg &= ": " & SafeStr(err, "code") & " — " & SafeStr(err, "message")
                    End If
                End If
            Catch
                msg &= ": " & body
            End Try
            Return New M365GraphException(msg, CInt(status))
        End Function

        ''' <summary>
        ''' Converts a Graph mail message id to the URL-safe form accepted in
        ''' /me/messages/{id}. Accepts raw Base64-style ids, already URL-safe ids,
        ''' and defensively unescapes once when a percent-encoded id is supplied.
        ''' </summary>
        Friend Function ToGraphUrlSafeId(value As String) As String
            Dim s As String = If(value, "").Trim()
            If String.IsNullOrWhiteSpace(s) Then Return ""

            Try
                If s.IndexOf("%"c) >= 0 Then
                    s = Uri.UnescapeDataString(s)
                End If
            Catch
            End Try

            s = s.Replace(" ", "")
            Return s.Replace("+"c, "-"c).Replace("/"c, "_"c)
        End Function

        ' ═══════════════════════════════════════════════════════════════════════
        '  STRING / JSON UTILITIES
        ' ═══════════════════════════════════════════════════════════════════════

        Private Function SafeStr(j As JObject, name As String) As String
            If j Is Nothing Then Return ""
            Dim t = j(name)
            If t Is Nothing OrElse t.Type = JTokenType.Null Then Return ""
            Return t.ToString()
        End Function

        Private Function TryDate(j As JObject, name As String) As DateTime?
            If j Is Nothing Then Return Nothing

            Dim t = j(name)
            If t Is Nothing OrElse t.Type = JTokenType.Null Then Return Nothing

            Dim s As String = t.ToString().Trim()
            If String.IsNullOrWhiteSpace(s) Then Return Nothing

            Dim d As DateTime

            ' 1) First handle explicit day-first dotted formats returned by some Graph search responses.
            Dim dottedDayFirstFormats As String() = {
        "dd.MM.yyyy HH:mm:ss",
        "d.MM.yyyy HH:mm:ss",
        "dd.M.yyyy HH:mm:ss",
        "d.M.yyyy HH:mm:ss",
        "dd.MM.yyyy HH:mm",
        "d.MM.yyyy HH:mm",
        "dd.M.yyyy HH:mm",
        "d.M.yyyy HH:mm",
        "dd.MM.yyyy",
        "d.MM.yyyy",
        "dd.M.yyyy",
        "d.M.yyyy"
    }

            If DateTime.TryParseExact(
        s,
        dottedDayFirstFormats,
        Globalization.CultureInfo.InvariantCulture,
        Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
        d) Then
                Return d
            End If

            ' 2) Then try strict ISO-style formats commonly returned by Graph.
            Dim isoFormats As String() = {
        "yyyy-MM-dd'T'HH:mm:ss'Z'",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'",
        "yyyy-MM-dd'T'HH:mm:ssK",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-dd HH:mm",
        "yyyy-MM-dd"
    }

            If DateTime.TryParseExact(
        s,
        isoFormats,
        Globalization.CultureInfo.InvariantCulture,
        Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
        d) Then
                Return d
            End If

            ' 3) Fallback: invariant parse for anything else.
            If DateTime.TryParse(
        s,
        Globalization.CultureInfo.InvariantCulture,
        Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
        d) Then
                Return d
            End If

            Return Nothing
        End Function

        Private Function StripHtml(html As String) As String
            If String.IsNullOrEmpty(html) Then Return ""
            Dim s = Regex.Replace(html, "<\s*(script|style)[^>]*>.*?<\s*/\s*\1\s*>", "",
                                  RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            s = Regex.Replace(s, "<[^>]+>", " ")
            s = s.Replace("&nbsp;", " ").Replace("&amp;", "&").Replace("&lt;", "<") _
                 .Replace("&gt;", ">").Replace("&quot;", """").Replace("&apos;", "'")
            s = Regex.Replace(s, "[ \t]{2,}", " ")
            s = Regex.Replace(s, "(\r?\n){3,}", vbCrLf & vbCrLf)
            Return s.Trim()
        End Function

        Private Function CleanSummary(s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            ' Graph wraps matched terms with <c0>…</c0>; remove markers but keep text.
            Dim cleaned = Regex.Replace(s, "</?c\d+>", "")
            Return StripHtml(cleaned)
        End Function

        Private Function TruncatePlain(s As String, maxLen As Integer) As String
            If String.IsNullOrEmpty(s) Then Return ""
            If s.Length <= maxLen Then Return s
            Return s.Substring(0, maxLen).TrimEnd() & "…"
        End Function

        Private Function MakeSafeFileName(name As String, fallback As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return fallback
            For Each c In Path.GetInvalidFileNameChars()
                name = name.Replace(c, "_"c)
            Next
            Return name
        End Function


        Private Sub EnsureModernTls()
            Try
                ServicePointManager.SecurityProtocol =
                    ServicePointManager.SecurityProtocol Or
                    SecurityProtocolType.Tls12 Or
                    CType(12288, SecurityProtocolType) ' Tls13 = 12288 on older .NET Framework
            Catch
                Try
                    ServicePointManager.SecurityProtocol =
                        ServicePointManager.SecurityProtocol Or
                        SecurityProtocolType.Tls12
                Catch
                End Try
            End Try
        End Sub

    End Module



End Namespace
