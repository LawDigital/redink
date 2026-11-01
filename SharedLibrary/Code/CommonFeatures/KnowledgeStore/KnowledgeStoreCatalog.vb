' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreCatalog.vb
' Purpose:
'   Manages the catalog of configured Knowledge Stores across central and local
'   configuration sources.
'
' Responsibilities:
'   - Load Knowledge Store definitions from central and user-local catalog files.
'   - Merge catalogs into a single effective store list.
'   - Resolve store metadata paths, wiki paths, and source-root paths.
'   - Create new local store definitions with default metadata.
'   - Save local catalog updates with safe write semantics.
'   - Expose helper methods for lookup, labeling, and permission checks.
'
' Notes:
'   - Central and local catalogs are merged into the effective runtime view.
'   - Per-store runtime artifacts live below each store's `.redink` directory.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Manages the catalog of named Knowledge Store definitions, loaded from central and local JSON files.
    ''' </summary>
    Public Class KnowledgeStoreCatalog

#Region "Data Model"

        ''' <summary>
        ''' Defines a single Knowledge Store: a watched directory with metadata.
        ''' </summary>
        Public Class KnowledgeStoreDefinition
            ''' <summary>Human-readable unique name (e.g., "Contracts", "Research").</summary>
            Public Property Name As String = ""

            ''' <summary>Root directory path to watch for source documents (supports env variables).</summary>
            Public Property SourcePath As String = ""

            ''' <summary>
            ''' Owner username (Environment.UserName). Only this user may write to the store's
            ''' index if Role = "shared". Empty = no restriction.
            ''' </summary>
            Public Property Owner As String = ""

            ''' <summary>
            ''' Role: "personal" (single-user), "shared" (multi-user, owner-restricted writes),
            ''' "readonly" (no indexing, query only).
            ''' </summary>
            Public Property Role As String = "personal"

            ''' <summary>Whether this store is active (monitored and queryable).</summary>
            Public Property Active As Boolean = True

            ''' <summary>Whether to scan subdirectories of SourcePath.</summary>
            Public Property ScanSubdirectories As Boolean = True

            ' ── Runtime-only fields (not serialized) ──

            ''' <summary>True if this definition came from the central catalog file.</summary>
            <JsonIgnore>
            Public Property IsFromCentralCatalog As Boolean = False

            ''' <summary>Expanded, resolved SourcePath (computed at load time).</summary>
            <JsonIgnore>
            Public Property ResolvedSourcePath As String = ""


            <JsonIgnore>
            Public Property StoreId As String = ""

            <JsonIgnore>
            Public Property DisplayLabel As String = ""
        End Class

#End Region

#Region "Constants"

        ''' <summary>Subfolder inside each store's SourcePath that holds index metadata and wiki pages.</summary>
        Public Const MetadataFolder As String = ".redink"

        ''' <summary>Identifies reserved Knowledge Store output, including Raw and Wiki.
        ''' Intake consumers must not recursively treat these artifacts as original sources.
        ''' Does not inspect, create, modify or delete any file.</summary>
        Public Shared Function IsGeneratedMetadataPath(path As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(path) Then Return False
            For Each component As System.String In path.Replace("\", "/").Split("/"c)
                If System.String.Equals(component, MetadataFolder, System.StringComparison.OrdinalIgnoreCase) Then Return True
            Next
            Return False
        End Function

        ''' <summary>Default filename used when a configured catalog path points to a directory only.</summary>
        Public Const DefaultCatalogFileName As String = "redink-ks-catalog.json"

        ''' <summary>Manifest file inside .redink/ listing all indexed documents.</summary>
        Public Const ManifestFile As String = "manifest.json"

        ''' <summary>Wiki subfolder inside .redink/ containing LLM-generated wiki pages.</summary>
        Public Const WikiFolder As String = "Wiki"

        ''' <summary>Auto-maintained catalog of all wiki pages with one-line summaries.</summary>
        Public Const IndexFile As String = "index.md"

        ''' <summary>Append-only chronological record of ingest/query/lint operations.</summary>
        Public Const LogFile As String = "log.md"

        ''' <summary>Per-store schema steering ingest, query, and linting behavior.</summary>
        Public Const SchemaFile As String = "schema.json"

#End Region

#Region "Load / Merge"

        ''' <summary>
        ''' Loads and merges Knowledge Store definitions from both catalog files.
        ''' Local definitions override central ones with the same Name (case-insensitive).
        ''' Auto-creates missing local catalog files with an empty array.
        ''' </summary>
        ' Replace LoadAll and add the helper methods below it.

        Public Shared Function LoadAll(context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            EnsureCatalogFile(context.INI_KnowledgeStorePathLocal, isLocal:=True)
            EnsureCatalogFile(context.INI_KnowledgeStorePath, isLocal:=False)
            Return LoadOverviewReadOnly(context)
        End Function

        ''' <summary>Same catalog resolution/merge as LoadAll, without initialization, repair or writes.</summary>
        Public Shared Function LoadOverviewReadOnly(context As ISharedContext) As System.Collections.Generic.List(Of KnowledgeStoreDefinition)
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            Dim merged As New System.Collections.Generic.List(Of KnowledgeStoreDefinition)()
            merged.AddRange(LoadConfiguredStores(context.INI_KnowledgeStorePathLocal, isFromCentral:=False, context:=context))
            merged.AddRange(LoadConfiguredStores(context.INI_KnowledgeStorePath, isFromCentral:=True, context:=context))

            ApplyRuntimeMetadata(merged)

            Dim deduped As New System.Collections.Generic.List(Of KnowledgeStoreDefinition)()
            Dim seenIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            For Each def In merged
                Dim id = If(def.StoreId, "").Trim()
                If id <> "" AndAlso seenIds.Contains(id) Then Continue For
                If id <> "" Then seenIds.Add(id)
                deduped.Add(def)
            Next

            ApplyRuntimeMetadata(deduped)
            Return deduped
        End Function

        Private Shared Function LoadConfiguredStores(rawPath As String,
                                             isFromCentral As Boolean,
                                             context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            Dim defs = LoadCatalogFile(rawPath, isFromCentral)
            If defs.Count > 0 Then Return defs

            Dim directPath = ExpandEnvironmentVariables(StripQuotes(rawPath))
            If String.IsNullOrWhiteSpace(directPath) OrElse Not Directory.Exists(directPath) Then
                Return defs
            End If

            Dim owner As String = ""
            If context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStoreOwner) Then
                owner = context.INI_KnowledgeStoreOwner.Trim()
            End If

            Dim defaultName = Path.GetFileName(directPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
            If String.IsNullOrWhiteSpace(defaultName) Then
                defaultName = If(isFromCentral, "Central Knowledge Store", "Local Knowledge Store")
            End If

            defs.Add(New KnowledgeStoreDefinition() With {
        .Name = defaultName.Trim(),
        .SourcePath = StripQuotes(rawPath),
        .Owner = owner,
        .Role = If(isFromCentral, "shared", "personal"),
        .Active = True,
        .ScanSubdirectories = True,
        .IsFromCentralCatalog = isFromCentral,
        .ResolvedSourcePath = directPath
    })

            Return defs
        End Function

        Private Shared Sub ApplyRuntimeMetadata(definitions As IEnumerable(Of KnowledgeStoreDefinition))
            If definitions Is Nothing Then Return

            Dim nameCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For Each def In definitions
                If def Is Nothing Then Continue For

                def.StoreId = BuildStoreId(def)

                Dim cleanName = If(def.Name, "").Trim()
                If cleanName <> "" Then
                    If nameCounts.ContainsKey(cleanName) Then
                        nameCounts(cleanName) += 1
                    Else
                        nameCounts.Add(cleanName, 1)
                    End If
                End If
            Next

            For Each def In definitions
                If def Is Nothing Then Continue For

                Dim cleanName = If(def.Name, "").Trim()
                Dim sourceLabel = If(def.IsFromCentralCatalog, "Central", "Local")
                Dim includePath As Boolean = False

                If cleanName = "" Then
                    def.DisplayLabel = If(String.IsNullOrWhiteSpace(def.ResolvedSourcePath), sourceLabel, $"{sourceLabel}: {def.ResolvedSourcePath}")
                    Continue For
                End If

                If nameCounts.ContainsKey(cleanName) AndAlso nameCounts(cleanName) > 1 Then
                    includePath = True
                End If

                If includePath AndAlso Not String.IsNullOrWhiteSpace(def.ResolvedSourcePath) Then
                    def.DisplayLabel = $"{cleanName} ({sourceLabel}: {def.ResolvedSourcePath})"
                Else
                    def.DisplayLabel = $"{cleanName} ({sourceLabel})"
                End If
            Next
        End Sub

        Private Shared Function BuildStoreId(def As KnowledgeStoreDefinition) As String
            If def Is Nothing Then Return ""

            Dim sourceKey = NormalizeStorePathForId(def.ResolvedSourcePath)
            If sourceKey = "" Then
                sourceKey = NormalizeStorePathForId(def.SourcePath)
            End If

            Dim sourceKind = If(def.IsFromCentralCatalog, "CENTRAL", "LOCAL")
            Dim namePart = If(def.Name, "").Trim().ToUpperInvariant()

            Return $"{sourceKind}|{namePart}|{sourceKey}"
        End Function

        Private Shared Function NormalizeStorePathForId(pathValue As String) As String
            Dim cleaned = ExpandEnvironmentVariables(StripQuotes(If(pathValue, "").Trim()))
            If String.IsNullOrWhiteSpace(cleaned) Then Return ""

            cleaned = cleaned.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)

            Try
                Return Path.GetFullPath(cleaned).ToUpperInvariant()
            Catch
                Return cleaned.ToUpperInvariant()
            End Try
        End Function

        Public Shared Function GetDisplayLabel(def As KnowledgeStoreDefinition) As String
            If def Is Nothing Then Return ""
            If Not String.IsNullOrWhiteSpace(def.DisplayLabel) Then Return def.DisplayLabel
            Return If(def.Name, "").Trim()
        End Function

        Public Shared Function GetStoreById(storeId As String, context As ISharedContext) As KnowledgeStoreDefinition
            If String.IsNullOrWhiteSpace(storeId) Then Return Nothing

            Return LoadAll(context).FirstOrDefault(
        Function(d) String.Equals(d.StoreId, storeId, StringComparison.OrdinalIgnoreCase))
        End Function

        Public Shared Function GetStoresByName(name As String, context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            If String.IsNullOrWhiteSpace(name) Then
                Return New List(Of KnowledgeStoreDefinition)()
            End If

            Dim value = name.Trim()

            Return LoadAll(context).Where(
        Function(d) String.Equals(If(d.Name, "").Trim(), value, StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(If(d.DisplayLabel, "").Trim(), value, StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(If(d.StoreId, "").Trim(), value, StringComparison.OrdinalIgnoreCase)).ToList()
        End Function

        Public Shared Function GetStoreByName(name As String, context As ISharedContext) As KnowledgeStoreDefinition
            Return GetStoresByName(name, context).FirstOrDefault()
        End Function

        ''' <summary>
        ''' Resolves a configured catalog path. When a directory path is supplied instead
        ''' of a JSON filename, the default catalog filename is appended automatically.
        ''' </summary>
        Friend Shared Function ResolveCatalogPath(rawPath As String) As String
            If String.IsNullOrWhiteSpace(rawPath) Then Return ""

            Dim cleaned = StripQuotes(rawPath)
            If String.IsNullOrWhiteSpace(cleaned) Then Return ""

            Dim expanded = ExpandEnvironmentVariables(cleaned.Trim())
            If String.IsNullOrWhiteSpace(expanded) Then Return ""

            Dim endsWithSeparator =
                expanded.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) OrElse
                expanded.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)

            Dim lastSegment As String = ""
            Try
                lastSegment = Path.GetFileName(expanded.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
            Catch
                lastSegment = ""
            End Try

            Dim treatAsDirectory =
                endsWithSeparator OrElse
                Directory.Exists(expanded) OrElse
                String.IsNullOrWhiteSpace(Path.GetExtension(lastSegment))

            If Not treatAsDirectory Then
                Return expanded
            End If

            Dim basePath = expanded
            If endsWithSeparator Then
                basePath = basePath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                If basePath.EndsWith(":", StringComparison.Ordinal) Then
                    basePath &= Path.DirectorySeparatorChar
                End If
            End If

            Return Path.Combine(basePath, DefaultCatalogFileName)
        End Function

        ''' <summary>
        ''' Ensures a catalog file exists. If the path is configured but the file is missing,
        ''' auto-creates it with an empty JSON array (local only). If the file exists but
        ''' contains invalid JSON, logs an error and resets it (local only).
        ''' </summary>
        Private Shared Sub EnsureCatalogFile(rawPath As String, isLocal As Boolean)
            If String.IsNullOrWhiteSpace(rawPath) Then Return

            Try
                Dim expanded = ResolveCatalogPath(rawPath)
                If String.IsNullOrWhiteSpace(expanded) Then Return

                ' Auto-create missing file (local catalog only — central is admin-managed)
                If Not File.Exists(expanded) Then
                    If isLocal Then
                        Dim dir = Path.GetDirectoryName(expanded)
                        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                            Directory.CreateDirectory(dir)
                        End If
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Created empty local catalog at '{expanded}'")
                    Else
                        Debug.WriteLine($"KnowledgeStoreCatalog: Central catalog not found at '{expanded}' — skipping.")
                    End If
                    Return
                End If

                ' Check for empty/whitespace file — but try to recover from .tmp first
                Dim json = File.ReadAllText(expanded, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then
                    ' The file exists but is empty — likely a crash during write.
                    ' Try to recover from .tmp file (atomic write leftovers)
                    Dim tmpPath = expanded & ".tmp"
                    If File.Exists(tmpPath) Then
                        Dim tmpJson = File.ReadAllText(tmpPath, System.Text.Encoding.UTF8)
                        If Not String.IsNullOrWhiteSpace(tmpJson) Then
                            Try
                                Dim tmpToken = JToken.Parse(tmpJson)
                                If tmpToken.Type = JTokenType.Array Then
                                    File.WriteAllText(expanded, tmpJson, System.Text.Encoding.UTF8)
                                    Debug.WriteLine($"KnowledgeStoreCatalog: Recovered catalog from '{tmpPath}'")
                                    Return
                                End If
                            Catch
                                ' .tmp is also bad — fall through
                            End Try
                        End If
                    End If

                    ' Also try .bak
                    Dim bakPath = expanded & ".bak"
                    If File.Exists(bakPath) Then
                        Dim bakJson = File.ReadAllText(bakPath, System.Text.Encoding.UTF8)
                        If Not String.IsNullOrWhiteSpace(bakJson) Then
                            Try
                                Dim bakToken = JToken.Parse(bakJson)
                                If bakToken.Type = JTokenType.Array Then
                                    File.WriteAllText(expanded, bakJson, System.Text.Encoding.UTF8)
                                    Debug.WriteLine($"KnowledgeStoreCatalog: Recovered catalog from '{bakPath}'")
                                    Return
                                End If
                            Catch
                            End Try
                        End If
                    End If

                    If isLocal Then
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Empty local catalog — no recovery source found, reset to '[]'.")
                    End If
                    Return
                End If

                ' Validate JSON structure
                Try
                    Dim token = JToken.Parse(json)
                    If token.Type <> JTokenType.Array Then
                        Dim msg = $"KnowledgeStoreCatalog: '{expanded}' does not contain a JSON array (found {token.Type}). " &
                                  "Expected format: [ {{ ""Name"": ""..."", ""SourcePath"": ""..."", ... }} ]"
                        Debug.WriteLine(msg)
                        If isLocal Then
                            Dim backup = expanded & ".bak"
                            Try : File.Copy(expanded, backup, overwrite:=True) : Catch : End Try
                            File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                            Debug.WriteLine($"KnowledgeStoreCatalog: Local catalog was invalid — backed up to '{backup}' and reset.")
                        End If
                    End If
                Catch ex As JsonReaderException
                    Dim msg = $"KnowledgeStoreCatalog: '{expanded}' contains invalid JSON: {ex.Message}. " &
                              "Expected format: [ {{ ""Name"": ""..."", ""SourcePath"": ""..."", ... }} ]"
                    Debug.WriteLine(msg)
                    If isLocal Then
                        Dim backup = expanded & ".bak"
                        Try : File.Copy(expanded, backup, overwrite:=True) : Catch : End Try
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Local catalog was malformed — backed up to '{backup}' and reset.")
                    End If
                End Try
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreCatalog: EnsureCatalogFile error for '{rawPath}': {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Loads store definitions from a single catalog JSON file.
        ''' </summary>
        Private Shared Function LoadCatalogFile(rawPath As String, isFromCentral As Boolean) As List(Of KnowledgeStoreDefinition)
            Dim result As New List(Of KnowledgeStoreDefinition)()
            If String.IsNullOrWhiteSpace(rawPath) Then Return result

            Try
                Dim expanded = ResolveCatalogPath(rawPath)
                If String.IsNullOrWhiteSpace(expanded) OrElse Not File.Exists(expanded) Then Return result

                Dim json = File.ReadAllText(expanded, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then Return result

                Dim arr = JArray.Parse(json)
                For Each token As JToken In arr
                    Try
                        If token.Type <> JTokenType.Object Then Continue For
                        Dim obj = CType(token, JObject)

                        Dim def As New KnowledgeStoreDefinition() With {
                            .Name = If(obj.Value(Of String)("Name"), "").Trim(),
                            .SourcePath = StripQuotes(If(obj.Value(Of String)("SourcePath"), "")),
                            .Owner = If(obj.Value(Of String)("Owner"), "").Trim(),
                            .Role = If(obj.Value(Of String)("Role"), "personal").Trim(),
                            .Active = If(obj("Active") IsNot Nothing, obj.Value(Of Boolean)("Active"), True),
                            .ScanSubdirectories = If(obj("ScanSubdirectories") IsNot Nothing, obj.Value(Of Boolean)("ScanSubdirectories"), True),
                            .IsFromCentralCatalog = isFromCentral
                        }
                        ' Resolve the source path
                        def.ResolvedSourcePath = ExpandEnvironmentVariables(def.SourcePath)

                        If Not String.IsNullOrWhiteSpace(def.Name) Then
                            result.Add(def)
                        End If
                    Catch
                        ' Skip malformed entries
                    End Try
                Next
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreCatalog: Error loading '{rawPath}': {ex.Message}")
            End Try

            Return result
        End Function

#End Region

#Region "Save"

        ''' <summary>
        ''' Saves local store definitions to the local catalog file.
        ''' Only definitions where IsFromCentralCatalog = False are written.
        ''' Uses atomic write (tmp + move) to prevent corruption.
        ''' </summary>
        Public Shared Sub SaveLocalCatalog(definitions As List(Of KnowledgeStoreDefinition), context As ISharedContext)
            If String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal is not configured.")
            End If

            Dim expanded = ResolveCatalogPath(context.INI_KnowledgeStorePathLocal)
            If String.IsNullOrWhiteSpace(expanded) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal could not be expanded.")
            End If

            Dim localDefs = definitions.Where(Function(d) Not d.IsFromCentralCatalog).ToList()

            Dim arr As New JArray()
            For Each def In localDefs
                Dim obj As New JObject()
                obj("Name") = def.Name
                obj("SourcePath") = def.SourcePath
                obj("Owner") = If(def.Owner, "")
                obj("Role") = If(def.Role, "personal")
                obj("Active") = def.Active
                obj("ScanSubdirectories") = def.ScanSubdirectories
                arr.Add(obj)
            Next

            ' Ensure directory exists
            Dim dir = Path.GetDirectoryName(expanded)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' Atomic write: write to .tmp, then replace original
            Dim tmpPath = expanded & ".tmp"
            File.WriteAllText(tmpPath, arr.ToString(Formatting.Indented), System.Text.Encoding.UTF8)
            If File.Exists(expanded) Then
                File.Delete(expanded)
            End If
            File.Move(tmpPath, expanded)
        End Sub

#End Region

#Region "Query Helpers"


        ''' <summary>
        ''' Returns True if at least one catalog file is configured.
        ''' </summary>
        Public Shared Function IsConfigured(context As ISharedContext) As Boolean
            Return Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePath) OrElse
                   Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal)
        End Function

        ''' <summary>
        ''' Returns all active store definitions.
        ''' </summary>
        Public Shared Function GetActiveStores(context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            Return LoadAll(context).Where(Function(d) d.Active).ToList()
        End Function


        ''' <summary>
        ''' Returns True if the current user is allowed to write (index) to this store.
        ''' Checks both Environment.UserName and the configured INI_KnowledgeStoreOwner.
        ''' </summary>
        Public Shared Function CanCurrentUserWrite(def As KnowledgeStoreDefinition, Optional context As ISharedContext = Nothing) As Boolean
            If def Is Nothing Then Return False
            If String.Equals(def.Role, "readonly", StringComparison.OrdinalIgnoreCase) Then Return False
            If String.IsNullOrWhiteSpace(def.Owner) Then Return True

            ' Match against current Windows username
            If String.Equals(def.Owner, Environment.UserName, StringComparison.OrdinalIgnoreCase) Then Return True

            ' Match against configured owner identity (if different from Windows username)
            If context IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStoreOwner) AndAlso
               String.Equals(def.Owner, context.INI_KnowledgeStoreOwner, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Returns the .redink/ metadata folder path for a given store.
        ''' Creates it if it does not exist and the user has write permission.
        ''' </summary>
        Public Shared Function GetMetadataPath(def As KnowledgeStoreDefinition, Optional createIfMissing As Boolean = True) As String
            If def Is Nothing OrElse String.IsNullOrWhiteSpace(def.ResolvedSourcePath) Then Return ""
            Dim metaDir = Path.Combine(def.ResolvedSourcePath, MetadataFolder)
            If createIfMissing AndAlso Not Directory.Exists(metaDir) Then
                Try
                    Directory.CreateDirectory(metaDir)
                Catch
                    Return ""
                End Try
            End If
            Return metaDir
        End Function

        ''' <summary>
        ''' Returns the wiki/ folder path for a given store.
        ''' Creates it if it does not exist and the user has write permission.
        ''' </summary>
        Public Shared Function GetWikiPath(def As KnowledgeStoreDefinition, Optional createIfMissing As Boolean = True) As String
            Dim metaDir = GetMetadataPath(def, createIfMissing)
            If String.IsNullOrWhiteSpace(metaDir) Then Return ""
            Dim wikiDir = Path.Combine(metaDir, WikiFolder)
            If createIfMissing AndAlso Not Directory.Exists(wikiDir) Then
                Try
                    Directory.CreateDirectory(wikiDir)
                Catch
                    Return ""
                End Try
            End If
            Return wikiDir
        End Function

        ''' <summary>
        ''' Creates a new KnowledgeStoreDefinition with the Owner defaulted from
        ''' <see cref="ISharedContext.INI_KnowledgeStoreOwner"/> (falls back to %USERNAME%).
        ''' </summary>
        Public Shared Function CreateDefinition(name As String, sourcePath As String, context As ISharedContext) As KnowledgeStoreDefinition
            Dim cleanPath = StripQuotes(sourcePath)
            Return New KnowledgeStoreDefinition() With {
                .Name = If(name, "").Trim(),
                .SourcePath = cleanPath,
                .Owner = If(String.IsNullOrWhiteSpace(context.INI_KnowledgeStoreOwner),
                            Environment.UserName,
                            context.INI_KnowledgeStoreOwner),
                .Role = "personal",
                .Active = True,
                .ScanSubdirectories = True,
                .ResolvedSourcePath = ExpandEnvironmentVariables(cleanPath)
            }
        End Function


        ''' <summary>
        ''' Strips leading/trailing quotes and whitespace from a path string.
        ''' Paths pasted from Windows Explorer or wrapped in INI values often contain quotes
        ''' that cause File.Exists, Directory.GetFiles, and Path.Combine to fail.
        ''' </summary>
        Friend Shared Function StripQuotes(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then Return ""
            Return path.Trim().Trim(""""c, "'"c).Trim()
        End Function


#End Region

    End Class

End Namespace