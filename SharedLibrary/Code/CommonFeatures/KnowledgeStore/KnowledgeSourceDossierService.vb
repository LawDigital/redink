' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeSourceDossierService.vb
' Purpose:
'   Maintains stable Knowledge Store source dossiers and deduplicated source
'   blobs under the store wiki, returning durable references that other
'   indexing and wiki services can link to.
'
' Architecture:
'   Resolves source paths against the Knowledge Store root, hashes/copies the
'   source into the configured blob folder, derives a stable dossier slug, and
'   creates or reuses the corresponding Markdown dossier page. It is a shared
'   persistence helper; retrieval/ranking and LLM processing live elsewhere.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions

Namespace SharedLibrary
    Friend Module KnowledgeSourceDossierService

        Friend Class DossierRef
            Public Property Slug As String = ""               ' file name without .md
            Public Property Title As String = ""              ' e.g. original file name
            Public Property DossierRelPath As String = ""     ' "Sources/<slug>.md"
            Public Property BlobRelPath As String = ""        ' "_files/<hash><ext>" (relative to docs root)
            Public Property OriginalPath As String = ""
            Public Property Hash As String = ""
        End Class

        ''' <summary>
        ''' Ensures the file exists once under Wiki/_files/ (hash-named if dedup),
        ''' and that a dossier page Wiki/Sources/&lt;slug&gt;.md exists.
        ''' Returns a DossierRef the caller can link to.
        ''' </summary>
        Friend Function EnsureDossier(kbRootPath As String,
                                      sourceFilePath As String,
                                      schema As KnowledgeStoreSchema) As DossierRef
            If String.IsNullOrWhiteSpace(sourceFilePath) Then Return Nothing

            ' Callers may hand us either an absolute path or a path that has
            ' already been normalized to be relative to the store root. Resolve
            ' to an absolute, on-disk path before doing anything that needs the
            ' actual bytes (hash, copy, dossier title).
            Dim resolvedSourcePath As String = ResolveSourcePath(kbRootPath, sourceFilePath)
            If String.IsNullOrWhiteSpace(resolvedSourcePath) OrElse Not File.Exists(resolvedSourcePath) Then
                Return Nothing
            End If

            Dim filesFolder As String =
                    If(schema IsNot Nothing AndAlso schema.SourceAccess IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(schema.SourceAccess.FilesFolderName),
                       schema.SourceAccess.FilesFolderName, "_files")

            Dim sourcesFolder As String =
                    If(schema IsNot Nothing AndAlso schema.SourceRegistry IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(schema.SourceRegistry.FolderName),
                       schema.SourceRegistry.FolderName, "Sources")


            Dim wikiRoot = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            Dim filesRoot = Path.Combine(wikiRoot, filesFolder)
            Dim sourcesRoot = Path.Combine(wikiRoot, sourcesFolder)

            If Not Directory.Exists(filesRoot) Then Directory.CreateDirectory(filesRoot)
            If Not Directory.Exists(sourcesRoot) Then Directory.CreateDirectory(sourcesRoot)

            Dim ext As String = Path.GetExtension(resolvedSourcePath)

            Dim allowed As String() = Nothing
            If schema IsNot Nothing AndAlso schema.SourceAccess IsNot Nothing Then
                allowed = schema.SourceAccess.AllowedExtensions
            End If

            If allowed IsNot Nothing AndAlso allowed.Length > 0 AndAlso
               Not allowed.Any(Function(a) Not String.IsNullOrWhiteSpace(a) AndAlso
                                            a.Equals(ext, StringComparison.OrdinalIgnoreCase)) Then
                Return Nothing
            End If

            ' --- 1. Hash & blob copy ---
            Dim hash = ComputeSha256(resolvedSourcePath)
            Dim blobName As String
            Dim deduplicate As Boolean = True
            If schema IsNot Nothing AndAlso schema.SourceAccess IsNot Nothing Then
                deduplicate = schema.SourceAccess.DeduplicateByHash
            End If

            If deduplicate Then
                blobName = hash & ext.ToLowerInvariant()
            Else
                blobName = SafeFileName(Path.GetFileNameWithoutExtension(resolvedSourcePath)) & "_" & hash.Substring(0, 8) & ext.ToLowerInvariant()
            End If
            Dim blobPath = Path.Combine(filesRoot, blobName)

            If Not File.Exists(blobPath) Then
                Try
                    File.Copy(resolvedSourcePath, blobPath, overwrite:=False)
                Catch
                End Try
            End If

            ' --- 2. Dossier page ---
            Dim slug = SafeFileName(Path.GetFileNameWithoutExtension(resolvedSourcePath))
            If String.IsNullOrWhiteSpace(slug) Then slug = hash.Substring(0, 12)
            Dim dossierPath = Path.Combine(sourcesRoot, slug & ".md")

            Dim ref As New DossierRef With {
                .Slug = slug,
                .Title = Path.GetFileName(resolvedSourcePath),
                .DossierRelPath = (sourcesFolder & "/" & slug & ".md").Replace("\"c, "/"c),
                .BlobRelPath = (filesFolder & "/" & blobName).Replace("\"c, "/"c),
                .OriginalPath = resolvedSourcePath,
                .Hash = hash
            }

            If Not File.Exists(dossierPath) Then
                File.WriteAllText(dossierPath, BuildInitialDossier(ref, schema), Encoding.UTF8)
            Else
                ' touch: guarantee the Download link points at the current blob
                File.WriteAllText(dossierPath,
                                  ReconcileDossierBlobLink(File.ReadAllText(dossierPath, Encoding.UTF8), ref, schema),
                                  Encoding.UTF8)
            End If

            Return ref
        End Function

        ''' <summary>
        ''' Resolves an incoming source-path token to an absolute on-disk path.
        ''' Accepts:
        '''   - an already-absolute local or UNC path;
        '''   - a path with environment variables;
        '''   - a path that has been normalized to be relative to the store root.
        ''' </summary>
        Private Function ResolveSourcePath(kbRootPath As String, sourcePath As String) As String
            Dim cleaned As String
            Try
                cleaned = SharedMethods.ExpandEnvironmentVariables(
                    If(sourcePath, "").Trim().Trim(""""c, "'"c))
            Catch
                cleaned = If(sourcePath, "").Trim().Trim(""""c, "'"c)
            End Try

            If String.IsNullOrWhiteSpace(cleaned) Then Return ""

            ' Already absolute (drive-letter or UNC)?
            If Path.IsPathRooted(cleaned) Then
                If File.Exists(cleaned) Then
                    Try
                        Return Path.GetFullPath(cleaned)
                    Catch
                        Return cleaned
                    End Try
                End If
                Return cleaned
            End If

            ' Relative — resolve against the store root.
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return cleaned

            Dim normalizedRoot As String
            Try
                normalizedRoot = Path.GetFullPath(kbRootPath).
                    TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Catch
                normalizedRoot = kbRootPath.Trim().
                    TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            End Try

            Try
                Dim combined = Path.GetFullPath(
                    Path.Combine(normalizedRoot,
                                 cleaned.Replace("/"c, Path.DirectorySeparatorChar)))
                Return combined
            Catch
                Return cleaned
            End Try
        End Function

        ''' <summary>
        ''' Adds/updates a "Cited By" entry in the dossier so users can navigate
        ''' from a source back to all pages that reference it.
        ''' </summary>
        Friend Sub AddCitedBy(kbRootPath As String,
                              dossier As DossierRef,
                              citingTitle As String,
                              citingPageRelPathFromWikiRoot As String)
            If dossier Is Nothing Then Return
            Dim wikiRoot = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            Dim dossierFull = Path.Combine(wikiRoot, dossier.DossierRelPath.Replace("/"c, Path.DirectorySeparatorChar))
            If Not File.Exists(dossierFull) Then Return

            Dim text = File.ReadAllText(dossierFull, Encoding.UTF8)
            ' Dossier sits in Sources\, citing page rel path is from wiki root.
            ' Convert to relative-from-dossier:
            Dim linkTarget = "../" & citingPageRelPathFromWikiRoot.Replace("\"c, "/"c)
            Dim entry = $"- [{citingTitle}]({linkTarget})"

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim citedByHeading As String = If(schema Is Nothing, "## Cited By", schema.GetLabel("CitedByHeading", "## Cited By"))
            text = UpsertBulletInSection(text, citedByHeading, entry)
            File.WriteAllText(dossierFull, text, Encoding.UTF8)
        End Sub

        ' ---------- helpers ----------

        Private Function BuildInitialDossier(ref As DossierRef, schema As KnowledgeStoreSchema) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("---")
            sb.AppendLine($"title: ""{ref.Title.Replace("""", "\""")}""")
            sb.AppendLine($"summary: ""Source dossier for {ref.Title.Replace("""", "\""")}.""")
            sb.AppendLine("kind: ""SourceDossier""")
            sb.AppendLine("status: ""stable""")
            sb.AppendLine($"hash: ""{ref.Hash}""")
            sb.AppendLine("source_paths:")
            sb.AppendLine($"  - '{ref.OriginalPath.Replace("'", "''")}'")
            sb.AppendLine($"blob_path: ""{ref.BlobRelPath}""")
            ' Schema-defined empty metadata fields, to be filled in by LLM or user.
            If schema IsNot Nothing AndAlso
                   schema.SourceRegistry IsNot Nothing AndAlso
                   schema.SourceRegistry.MetadataFields IsNot Nothing Then
                For Each fld In schema.SourceRegistry.MetadataFields
                    If Not String.IsNullOrWhiteSpace(fld) Then
                        sb.AppendLine($"{fld}: """"")
                    End If
                Next
            End If
            sb.AppendLine($"updated_utc: ""{DateTime.UtcNow:o}""")
            sb.AppendLine("---")
            sb.AppendLine()
            Dim downloadHeading As String = If(schema Is Nothing, "## Download", schema.GetLabel("DownloadHeading", "## Download"))
            Dim originalLocationHeading As String = If(schema Is Nothing, "## Original Location", schema.GetLabel("OriginalLocationHeading", "## Original Location"))
            Dim citedByHeading As String = If(schema Is Nothing, "## Cited By", schema.GetLabel("CitedByHeading", "## Cited By"))
            Dim noneYetPlaceholder As String = If(schema Is Nothing, "- _None yet._", "- " & schema.GetLabel("PlaceholderNoneYet", "_None yet._"))

            sb.AppendLine(downloadHeading)
            sb.AppendLine()
            ' Dossier sits at <SourcesFolder>\<slug>.md, blob at <FilesFolder>\<hash><ext>.
            sb.AppendLine($"- [{ref.Title}](../{ref.BlobRelPath})")
            sb.AppendLine()
            sb.AppendLine(originalLocationHeading)
            sb.AppendLine()
            sb.AppendLine($"- `{ref.OriginalPath}`")
            sb.AppendLine()
            sb.AppendLine(citedByHeading)
            sb.AppendLine()
            sb.AppendLine(noneYetPlaceholder)
            Return sb.ToString()
        End Function

        Private Function ReconcileDossierBlobLink(existing As String, ref As DossierRef, schema As KnowledgeStoreSchema) As String
            Dim downloadHeading As String = If(schema Is Nothing, "## Download", schema.GetLabel("DownloadHeading", "## Download"))

            ' The regex must match BOTH the configured heading and the legacy English
            ' default, so dossiers created before localization can still be reconciled.
            Dim configuredEsc = Regex.Escape(downloadHeading.Replace("##", "").Trim())
            Dim pattern = "(?ims)(^##\s*(?:Download|" & configuredEsc & ")\s*$\s*)(.*?)(?=^##\s+|\z)"

            Dim newSection = downloadHeading & vbCrLf & vbCrLf &
                             $"- [{ref.Title}](../{ref.BlobRelPath})" & vbCrLf & vbCrLf

            If Regex.IsMatch(existing, pattern) Then
                Return Regex.Replace(existing, pattern, newSection)
            End If
            Return existing.TrimEnd() & vbCrLf & vbCrLf & newSection
        End Function

        Private Function UpsertBulletInSection(text As String, heading As String, bullet As String) As String
            Dim pattern = "(?ims)^" & Regex.Escape(heading) & "\s*$(?<body>.*?)(?=^##\s+|\z)"
            Dim m = Regex.Match(text, pattern)
            If Not m.Success Then
                Return text.TrimEnd() & vbCrLf & vbCrLf & heading & vbCrLf & vbCrLf & bullet & vbCrLf
            End If

            ' Extract the new bullet's link target so we dedupe by destination, not by literal string.
            Dim newTarget As String = ExtractMarkdownLinkTarget(bullet)
            Dim body = m.Groups("body").Value

            For Each rawLine In body.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim line = rawLine.Trim()
                If Not line.StartsWith("- ", StringComparison.Ordinal) Then Continue For
                Dim existingTarget = ExtractMarkdownLinkTarget(line.Substring(2))
                If Not String.IsNullOrWhiteSpace(existingTarget) AndAlso
           Not String.IsNullOrWhiteSpace(newTarget) AndAlso
           existingTarget.Equals(newTarget, StringComparison.OrdinalIgnoreCase) Then
                    Return text   ' already cited
                End If
            Next

            body = Regex.Replace(body, "(?im)^\s*-\s*_None yet\._\s*$\s*", "")
            Dim rebuilt = heading & vbCrLf & vbCrLf & body.Trim() & vbCrLf & bullet & vbCrLf & vbCrLf
            Return text.Substring(0, m.Index) & rebuilt & text.Substring(m.Index + m.Length)
        End Function

        Private Function ExtractMarkdownLinkTarget(value As String) As String
            Dim m = Regex.Match(If(value, ""), "\[[^\]]+\]\(([^)]+)\)")
            If m.Success Then Return m.Groups(1).Value.Trim()
            Return ""
        End Function



        Private Function ComputeSha256(path As String) As String
            Using sha = SHA256.Create()
                Using fs = File.OpenRead(path)
                    Return BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "").ToLowerInvariant()
                End Using
            End Using
        End Function

        Private Function SafeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(If(name, "").Where(Function(c) Not invalid.Contains(c)).ToArray()).Trim()
            cleaned = Regex.Replace(cleaned, "\s+", " ")
            If cleaned.Length > 120 Then cleaned = cleaned.Substring(0, 120).Trim()
            Return cleaned
        End Function

    End Module
End Namespace