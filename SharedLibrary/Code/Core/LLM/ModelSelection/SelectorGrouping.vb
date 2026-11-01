' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SelectorGrouping.vb
' PURPOSE
'   Defines optional group layouts, selection state and central/local presentation stores.
'
' KEY ARCHITECTURE
'   • Data: Validates group definitions and preserves stable item selections.
'   • Precedence: Personal layouts replace central layouts without changing permissions.
' =============================================================================

'
' Optional selector presentation metadata. No tool availability, permission, check-state,
' prompt or execution semantics belong in this file. Local files replace (not merge)
' central layouts. A missing/invalid file cannot remove entries from a selector.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class SelectorGroupDefinition
        Public Property Name As System.String = ""
        Public Property Expanded As System.Boolean = True
        Public Property Items As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public NotInheritable Class SelectorGroupItem
        Public Property Id As System.String = ""
        Public Property Label As System.String = ""
    End Class

    Public NotInheritable Class SelectorGroupDocument
        Public Const MaximumBytes As System.Int32 = 524288
        Public Property Groups As New System.Collections.Generic.List(Of SelectorGroupDefinition)()

        Public Function Copy() As SelectorGroupDocument
            Return Parse(ToJson())
        End Function

        Public Function GroupFor(itemId As System.String) As System.String
            For Each definition As SelectorGroupDefinition In Groups
                For Each assignedId As System.String In definition.Items
                    If System.String.Equals(itemId, assignedId, System.StringComparison.OrdinalIgnoreCase) Then Return definition.Name
                Next
            Next
            Return ""
        End Function

        Public Function ToJson() As System.String
            Dim groupArray As New Newtonsoft.Json.Linq.JArray()
            For Each definition As SelectorGroupDefinition In Groups
                Dim itemArray As New Newtonsoft.Json.Linq.JArray()
                For Each itemId As System.String In definition.Items
                    itemArray.Add(itemId)
                Next
                Dim groupObject As New Newtonsoft.Json.Linq.JObject()
                groupObject("name") = definition.Name
                groupObject("expanded") = definition.Expanded
                groupObject("items") = itemArray
                groupArray.Add(groupObject)
            Next
            Dim root As New Newtonsoft.Json.Linq.JObject()
            root("schema_version") = 1
            root("groups") = groupArray
            Return root.ToString(Newtonsoft.Json.Formatting.Indented) & System.Environment.NewLine
        End Function

        Public Shared Function Parse(json As System.String) As SelectorGroupDocument
            If json Is Nothing OrElse System.Text.Encoding.UTF8.GetByteCount(json) > MaximumBytes Then
                Throw New System.FormatException("The group file is empty or too large.")
            End If
            Dim settings As New Newtonsoft.Json.Linq.JsonLoadSettings() With {
                .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error
            }
            Dim root As Newtonsoft.Json.Linq.JObject
            Using textReader As New System.IO.StringReader(json.TrimStart(System.Convert.ToChar(&HFEFF)))
                Using reader As New Newtonsoft.Json.JsonTextReader(textReader)
                    reader.MaxDepth = 8
                    reader.DateParseHandling = Newtonsoft.Json.DateParseHandling.None
                    root = Newtonsoft.Json.Linq.JObject.Load(reader, settings)
                    While reader.Read()
                        If reader.TokenType <> Newtonsoft.Json.JsonToken.Comment Then
                            Throw New System.FormatException("Only one group document is allowed.")
                        End If
                    End While
                End Using
            End Using
            ValidateProperties(root, New System.String() {"schema_version", "groups"})
            If root("schema_version") Is Nothing OrElse root("schema_version").Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse
               root("schema_version").ToString() <> "1" Then
                Throw New System.FormatException("The group file must use schema_version 1.")
            End If
            Dim groupArray As Newtonsoft.Json.Linq.JArray = TryCast(root("groups"), Newtonsoft.Json.Linq.JArray)
            If groupArray Is Nothing OrElse groupArray.Count > 200 Then Throw New System.FormatException("The group list is missing or too large.")
            Dim result As New SelectorGroupDocument()
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim assignedIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each token As Newtonsoft.Json.Linq.JToken In groupArray
                Dim groupObject As Newtonsoft.Json.Linq.JObject = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If groupObject Is Nothing Then Throw New System.FormatException("Each group must be an object.")
                ValidateProperties(groupObject, New System.String() {"name", "expanded", "items"})
                Dim definition As New SelectorGroupDefinition() With {.Name = ReadName(groupObject("name"), 100)}
                If Not names.Add(definition.Name) Then Throw New System.FormatException("Group names must be unique.")
                If groupObject("expanded") IsNot Nothing Then
                    If groupObject("expanded").Type <> Newtonsoft.Json.Linq.JTokenType.Boolean Then Throw New System.FormatException("expanded must be true or false.")
                    definition.Expanded = groupObject("expanded").ToObject(Of System.Boolean)()
                End If
                Dim itemArray As Newtonsoft.Json.Linq.JArray = TryCast(groupObject("items"), Newtonsoft.Json.Linq.JArray)
                If itemArray Is Nothing Then Throw New System.FormatException("Each group needs an items array.")
                For Each itemToken As Newtonsoft.Json.Linq.JToken In itemArray
                    Dim itemId As System.String = ReadName(itemToken, 256)
                    If Not assignedIds.Add(itemId) Then Throw New System.FormatException("Each entry may belong to only one group.")
                    If assignedIds.Count > 10000 Then Throw New System.FormatException("Too many group assignments.")
                    definition.Items.Add(itemId)
                Next
                result.Groups.Add(definition)
            Next
            Return result
        End Function

        Private Shared Function ReadName(token As Newtonsoft.Json.Linq.JToken, maximumLength As System.Int32) As System.String
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.FormatException("A name or entry ID is missing.")
            Dim value As System.String = token.ToObject(Of System.String)().Trim()
            If value.Length = 0 OrElse value.Length > maximumLength Then Throw New System.FormatException("A name or entry ID has an invalid length.")
            For Each character As System.Char In value
                If System.Char.IsControl(character) Then Throw New System.FormatException("Names and entry IDs cannot contain control characters.")
            Next
            Return value
        End Function

        Private Shared Sub ValidateProperties(value As Newtonsoft.Json.Linq.JObject, allowed As System.String())
            For Each propertyValue As Newtonsoft.Json.Linq.JProperty In value.Properties()
                If System.Array.IndexOf(allowed, propertyValue.Name) < 0 Then Throw New System.FormatException("Unknown group file property: " & propertyValue.Name)
            Next
        End Sub

        ' Apply only rows actually displayed by this editor. Entries belonging to another
        ' host, Advanced Tools, unavailable connectors or future tools survive unchanged.
        Public Function WithAssignments(assignments As System.Collections.Generic.IDictionary(Of System.String, System.String)) As SelectorGroupDocument
            Dim result As SelectorGroupDocument = Copy()
            Dim changedIds As New System.Collections.Generic.HashSet(Of System.String)(assignments.Keys, System.StringComparer.OrdinalIgnoreCase)
            For Each definition As SelectorGroupDefinition In result.Groups
                definition.Items.RemoveAll(Function(itemId As System.String) changedIds.Contains(itemId))
            Next
            For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, System.String) In assignments
                Dim name As System.String = If(pair.Value, "").Trim()
                If name.Length = 0 Then Continue For
                Dim target As SelectorGroupDefinition = Nothing
                For Each definition As SelectorGroupDefinition In result.Groups
                    If System.String.Equals(definition.Name, name, System.StringComparison.OrdinalIgnoreCase) Then
                        target = definition
                        Exit For
                    End If
                Next
                If target Is Nothing Then
                    target = New SelectorGroupDefinition() With {.Name = name}
                    result.Groups.Add(target)
                End If
                target.Items.Add(pair.Key)
            Next
            result.Groups.RemoveAll(Function(definition As SelectorGroupDefinition) definition.Items.Count = 0)
            ' Use the same strict contract for hand-written files and editor output.
            Return Parse(result.ToJson())
        End Function
    End Class

    ' Pure selection state: IDs, not labels, positions, tree nodes or expanded state.
    Public NotInheritable Class SelectorGroupSelection
        Private ReadOnly _items As New System.Collections.Generic.List(Of SelectorGroupItem)()
        Private ReadOnly _selected As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _known As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        Public Property Layout As New SelectorGroupDocument()
        Public Property Filter As System.String = ""

        Public Sub New(items As System.Collections.Generic.IEnumerable(Of SelectorGroupItem), initiallySelected As System.Collections.Generic.IEnumerable(Of System.String))
            For Each item As SelectorGroupItem In items
                If item Is Nothing OrElse System.String.IsNullOrWhiteSpace(item.Id) OrElse Not _known.Add(item.Id) Then
                    Throw New System.ArgumentException("Grouping requires unique, nonempty entry IDs.")
                End If
                _items.Add(New SelectorGroupItem() With {.Id = item.Id, .Label = item.Label})
            Next
            If initiallySelected IsNot Nothing Then
                For Each itemId As System.String In initiallySelected
                    If _known.Contains(itemId) Then _selected.Add(itemId)
                Next
            End If
        End Sub

        Public ReadOnly Property Items As System.Collections.Generic.IReadOnlyList(Of SelectorGroupItem)
            Get
                Return _items.AsReadOnly()
            End Get
        End Property

        Public Function IsSelected(itemId As System.String) As System.Boolean
            Return _selected.Contains(itemId)
        End Function

        Public Sub SetSelected(itemId As System.String, value As System.Boolean)
            If Not _known.Contains(itemId) Then Return
            If value Then
                _selected.Add(itemId)
            Else
                _selected.Remove(itemId)
            End If
        End Sub

        Public Function SelectedIds() As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            For Each item As SelectorGroupItem In _items
                If _selected.Contains(item.Id) Then result.Add(item.Id)
            Next
            Return result
        End Function

        Public Function Matches(item As SelectorGroupItem) As System.Boolean
            Dim search As System.String = If(Filter, "").Trim()
            Return search.Length = 0 OrElse
                If(item.Label, "").IndexOf(search, System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                item.Id.IndexOf(search, System.StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                Layout.GroupFor(item.Id).IndexOf(search, System.StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Public Function MatchingItems() As System.Collections.Generic.List(Of SelectorGroupItem)
            Return _items.FindAll(Function(item As SelectorGroupItem) Matches(item))
        End Function

        Public Function AllMatchingSelected() As System.Boolean
            Dim matches As System.Collections.Generic.List(Of SelectorGroupItem) = MatchingItems()
            If matches.Count = 0 Then Return False
            For Each item As SelectorGroupItem In matches
                If Not IsSelected(item.Id) Then Return False
            Next
            Return True
        End Function

        Public Sub ToggleMatching()
            Dim value As System.Boolean = Not AllMatchingSelected()
            For Each item As SelectorGroupItem In MatchingItems()
                SetSelected(item.Id, value)
            Next
        End Sub
    End Class

    ' Files are loaded once per opening. No timers, watchers or file IO while filtering.
    Public NotInheritable Class SelectorGroupingStore
        Public Const FileName As System.String = "selector-groups.json"
        Public ReadOnly Property CentralFile As System.String
        Public ReadOnly Property LocalFile As System.String
        Private _document As New SelectorGroupDocument()
        Private _sourceFile As System.String = ""
        Private _warning As System.String = ""
        Public ReadOnly Property Document As SelectorGroupDocument
            Get
                Return _document
            End Get
        End Property
        Public ReadOnly Property SourceFile As System.String
            Get
                Return _sourceFile
            End Get
        End Property
        Public ReadOnly Property Warning As System.String
            Get
                Return _warning
            End Get
        End Property
        Private _localStamp As System.String = "missing"
        Private _centralStamp As System.String = "missing"

        Public Sub New(centralFile As System.String, localFile As System.String)
            Me.CentralFile = If(centralFile, "")
            Me.LocalFile = If(localFile, "")
        End Sub

        ' Adapter only: the form and data model do not know resource conventions.
        Public Shared Function ForResourceRoots(centralRoot As System.String, localRoot As System.String) As SelectorGroupingStore
            Dim local As System.String = localRoot
            ' No file/folder is created until the user explicitly chooses Save personal.
            If System.String.IsNullOrWhiteSpace(localRoot) Then
                Dim applicationData As System.String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
                local = If(System.String.IsNullOrWhiteSpace(applicationData), "", System.IO.Path.Combine(applicationData, "RedInk"))
            End If
            Return New SelectorGroupingStore(ConfigurationPath(centralRoot), ConfigurationPath(local))
        End Function

        Private Shared Function ConfigurationPath(root As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(root) Then Return ""
            Try
                Dim expanded As System.String = SharedMethods.ExpandEnvironmentVariables(root)
                If System.String.IsNullOrWhiteSpace(expanded) Then Throw New System.ArgumentException("The resource folder could not be resolved.")
                Return System.IO.Path.Combine(System.IO.Path.GetFullPath(expanded), FileName)
            Catch ex As System.Exception
                Log("invalid-resource-folder-flat-fallback", root, ex)
                Return ""
            End Try
        End Function

        Public ReadOnly Property CanSaveLocal As System.Boolean
            Get
                Return Not System.String.IsNullOrWhiteSpace(LocalFile) AndAlso Not SamePath(LocalFile, CentralFile)
            End Get
        End Property

        Public Sub Reload()
            _document = New SelectorGroupDocument()
            _sourceFile = ""
            _warning = ""
            Dim localDocument As SelectorGroupDocument = ReadCandidate(LocalFile, _localStamp)
            If localDocument IsNot Nothing Then
                _document = localDocument
                _sourceFile = LocalFile
                _centralStamp = "not-read"
                Return
            End If
            ' A valid personal file is a complete override; it needs no shared-folder IO.
            If SamePath(LocalFile, CentralFile) Then
                _centralStamp = _localStamp
                Return
            End If
            Dim centralDocument As SelectorGroupDocument = ReadCandidate(CentralFile, _centralStamp)
            If centralDocument IsNot Nothing Then
                _document = centralDocument
                _sourceFile = CentralFile
            End If
        End Sub

        Private Function ReadCandidate(path As System.String, ByRef stamp As System.String) As SelectorGroupDocument
            stamp = "missing"
            If System.String.IsNullOrWhiteSpace(path) Then Return Nothing
            Try
                Dim text As System.String = ReadText(path)
                If text Is Nothing Then Return Nothing
                stamp = Hash(text)
                Return SelectorGroupDocument.Parse(text)
            Catch ex As System.Exception
                If stamp = "missing" Then stamp = "unreadable"
                _warning = "A group file could not be read. A valid alternative or the flat list is being used. Open Groups to review the file locations."
                Log("load-fallback", path, ex)
                Return Nothing
            End Try
        End Function

        Public Sub SaveLocal(document As SelectorGroupDocument)
            If Not CanSaveLocal Then Throw New System.InvalidOperationException("A separate personal resource folder is required for local saving. Export the layout instead.")
            If _localStamp = "unreadable" Then Throw New System.IO.IOException("The personal file could not be read. It will not be overwritten.")
            ' Do not turn a stale central snapshot into a new overriding personal file.
            If SamePath(SourceFile, CentralFile) AndAlso CurrentStamp(CentralFile) <> _centralStamp Then
                Throw New System.IO.IOException("The shared layout changed while this window was open. Reopen Groups before saving.")
            End If
            WriteAtomic(LocalFile, document, _localStamp)
            Reload()
        End Sub

        Public Sub UseCentralOrFlat()
            If Not CanSaveLocal Then Throw New System.InvalidOperationException("The personal and shared paths must be different.")
            Dim current As System.String = CurrentStamp(LocalFile)
            If current <> _localStamp Then Throw New System.IO.IOException("The personal layout changed. Reopen Groups before resetting it.")
            If current = "missing" Then
                Reload()
                Return
            End If
            Using fileLock As New System.IO.FileStream(LocalFile & ".lock", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                If CurrentStamp(LocalFile) <> _localStamp Then Throw New System.IO.IOException("The personal layout changed. Reopen Groups before resetting it.")
                ' Keep the complete previous file; never change the shared resource.
                System.IO.File.Move(LocalFile, LocalFile & ".disabled-" & System.Guid.NewGuid().ToString("N") & ".bak")
            End Using
            Reload()
        End Sub

        Public Shared Function ReadDocument(path As System.String) As SelectorGroupDocument
            Dim text As System.String = ReadText(path)
            If text Is Nothing Then Throw New System.IO.FileNotFoundException("The group file was not found.")
            Return SelectorGroupDocument.Parse(text)
        End Function

        Public Shared Sub ExportDocument(path As System.String, document As SelectorGroupDocument)
            WriteAtomic(path, document, CurrentStamp(path))
        End Sub

        Private Shared Function ReadText(path As System.String) As System.String
            Try
                Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite Or System.IO.FileShare.Delete)
                    If stream.Length > SelectorGroupDocument.MaximumBytes Then Throw New System.IO.InvalidDataException("The group file is too large.")
                    Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), True)
                        Dim buffer(4095) As System.Char
                        Dim text As New System.Text.StringBuilder()
                        Do
                            Dim count As System.Int32 = reader.Read(buffer, 0, buffer.Length)
                            If count = 0 Then Exit Do
                            text.Append(buffer, 0, count)
                            If text.Length > SelectorGroupDocument.MaximumBytes Then Throw New System.IO.InvalidDataException("The group file is too large.")
                        Loop
                        Return text.ToString()
                    End Using
                End Using
            Catch ex As System.IO.FileNotFoundException
                Return Nothing
            Catch ex As System.IO.DirectoryNotFoundException
                Return Nothing
            End Try
        End Function

        Private Shared Function CurrentStamp(path As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(path) Then Return "missing"
            Dim text As System.String = ReadText(path)
            Return If(text Is Nothing, "missing", Hash(text))
        End Function

        Private Shared Function Hash(text As System.String) As System.String
            Using algorithm As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(algorithm.ComputeHash(System.Text.Encoding.UTF8.GetBytes(text))).Replace("-", "")
            End Using
        End Function

        Private Shared Sub WriteAtomic(path As System.String, document As SelectorGroupDocument, expectedStamp As System.String)
            Dim normalized As SelectorGroupDocument = SelectorGroupDocument.Parse(document.ToJson())
            Dim fullPath As System.String = System.IO.Path.GetFullPath(path)
            Dim directory As System.String = System.IO.Path.GetDirectoryName(fullPath)
            System.IO.Directory.CreateDirectory(directory)
            Dim temporary As System.String = fullPath & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                ' Serializes cooperating editors in Word and Outlook; a stale snapshot is rejected.
                Using fileLock As New System.IO.FileStream(fullPath & ".lock", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                    If CurrentStamp(fullPath) <> expectedStamp Then Throw New System.IO.IOException("The group file changed while this window was open. Reopen Groups before saving.")
                    Dim bytes As System.Byte() = New System.Text.UTF8Encoding(False).GetBytes(normalized.ToJson())
                    Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                        stream.Write(bytes, 0, bytes.Length)
                        stream.Flush(True)
                    End Using
                    If CurrentStamp(fullPath) <> expectedStamp Then Throw New System.IO.IOException("The group file changed during saving. Reopen Groups before saving.")
                    If expectedStamp = "missing" Then
                        System.IO.File.Move(temporary, fullPath)
                    Else
                        ' No non-atomic delete/copy fallback: on unsupported shares fail safely.
                        System.IO.File.Replace(temporary, fullPath, fullPath & ".bak")
                    End If
                End Using
                Log("saved", fullPath, Nothing)
            Finally
                If System.IO.File.Exists(temporary) Then
                    Try
                        System.IO.File.Delete(temporary)
                    Catch ex As System.Exception
                        Log("temporary-cleanup-failed", temporary, ex)
                    End Try
                End If
            End Try
        End Sub

        Public Shared Function SamePath(left As System.String, right As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(left) OrElse System.String.IsNullOrWhiteSpace(right) Then Return False
            Return System.String.Equals(System.IO.Path.GetFullPath(left), System.IO.Path.GetFullPath(right), System.StringComparison.OrdinalIgnoreCase)
        End Function

        Friend Shared Sub Log(stage As System.String, path As System.String, failure As System.Exception)
            Dim details As System.String = stage & "; file=" & If(path, "")
            If failure IsNot Nothing Then details &= "; " & failure.ToString()
            System.Diagnostics.Trace.WriteLine("[SelectorGrouping] " & details)
            RiCrashLogger.LogMarker("SELECTOR_GROUPING", details)
        End Sub
    End Class
End Namespace
