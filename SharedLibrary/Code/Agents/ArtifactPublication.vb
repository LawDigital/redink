' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Host-owned publication receipts and opt-in stable output-name binding.
' This layer never selects deliverables, changes their lifecycle, or infers opaque identities.
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq

Namespace Agents
    Public NotInheritable Class OutputPublicationReceipt
        Public ReadOnly Property RequestedRelativeName As System.String
        Public ReadOnly Property PublishedPath As System.String
        Public ReadOnly Property Sha256 As System.String

        Private Sub New(requested As System.String, published As System.String, digest As System.String)
            RequestedRelativeName = requested
            PublishedPath = published
            Sha256 = digest
        End Sub

        Public Shared Function Capture(requestedRelativeName As System.String, publishedPath As System.String,
                                       expectedSha256 As System.String) As OutputPublicationReceipt
            Dim relative = If(requestedRelativeName, "").Replace("\"c, "/"c)
            If System.String.IsNullOrWhiteSpace(relative) OrElse relative.StartsWith("/", System.StringComparison.Ordinal) OrElse
               relative.Contains(":") OrElse relative.Split("/"c).Any(Function(part) part = "" OrElse part = "." OrElse part = "..") Then
                Throw New System.IO.InvalidDataException("A publication receipt requires an exact relative producer name.")
            End If
            For Each part In relative.Split("/"c)
                OutputNameBinder.ValidateBasename(part)
            Next
            Dim full = System.IO.Path.GetFullPath(publishedPath)
            Dim digest = OutputNameBinder.FileHash(full)
            If System.String.IsNullOrWhiteSpace(expectedSha256) OrElse Not System.String.Equals(digest, expectedSha256, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.IO.InvalidDataException("Published output differs from the verified producer bytes.")
            End If
            Return New OutputPublicationReceipt(relative, full, digest)
        End Function
    End Class

    ' In-memory, per-run telemetry only. A receipt is NOT permission to deliver a file.
    ' Nested runs share this object just as they share the artifact registry. It is not
    ' reconstructed from model JSON or from old logs/checkpoints.
    Public NotInheritable Class OutputPublicationRegistry
        Private ReadOnly _sync As New System.Object()
        Private ReadOnly _receipts As New System.Collections.Generic.List(Of OutputPublicationReceipt)()

        Public Sub Record(receipts As System.Collections.Generic.IEnumerable(Of OutputPublicationReceipt))
            If receipts Is Nothing Then Return
            SyncLock _sync
                For Each receipt In receipts
                    If receipt Is Nothing Then Throw New System.ArgumentException("A publication receipt cannot be Nothing.", NameOf(receipts))
                    _receipts.Add(receipt)
                Next
            End SyncLock
        End Sub

        Public Function GetRequestedName(publishedPath As System.String, currentSha256 As System.String) As System.String
            Dim full = System.IO.Path.GetFullPath(publishedPath)
            Dim names As System.Collections.Generic.List(Of System.String)
            SyncLock _sync
                names = _receipts.Where(Function(receipt) System.String.Equals(receipt.PublishedPath, full, System.StringComparison.OrdinalIgnoreCase) AndAlso
                    System.String.Equals(receipt.Sha256, currentSha256, System.StringComparison.OrdinalIgnoreCase)).
                    Select(Function(receipt) receipt.RequestedRelativeName).Distinct(System.StringComparer.Ordinal).ToList()
            End SyncLock
            If names.Count > 1 Then Throw New System.IO.InvalidDataException("Conflicting publication names for one verified file: " & full)
            Return If(names.Count = 0, "", names(0))
        End Function

        Public Shared Function AddToToolResult(payload As System.String, receipts As System.Collections.Generic.IEnumerable(Of OutputPublicationReceipt)) As System.String
            If receipts Is Nothing Then Return payload
            Dim entries As New Newtonsoft.Json.Linq.JArray()
            For Each receipt In receipts
                entries.Add(New Newtonsoft.Json.Linq.JObject From {
                    {"requested_name", receipt.RequestedRelativeName},
                    {"published_name", System.IO.Path.GetFileName(receipt.PublishedPath)},
                    {"published_path", receipt.PublishedPath},
                    {"sha256", receipt.Sha256}})
            Next
            If entries.Count = 0 Then Return payload
            Dim result = Newtonsoft.Json.Linq.JObject.Parse(payload)
            result("host_publications") = entries
            result("publication_note") = "Published paths may have collision-safe names. Use the actual published file for further tool calls; do not rerun work merely to remove a collision suffix. These receipts do not replace artifact identity or delivery validation."
            Return result.ToString(Newtonsoft.Json.Formatting.None)
        End Function
    End Class

    Public NotInheritable Class OutputNameBinding
        Public Property Name As System.String = ""
        Public Property LogicalDeliverableId As System.String = ""
        Public Property OutputSlotId As System.String = ""
    End Class

    ' Only a caller's already-authorized final delivery set may enter the binder.
    Public NotInheritable Class OutputDeliveryCandidate
        Public Property Path As System.String = ""
        Public Property LogicalDeliverableId As System.String = ""
        Public Property OutputSlotId As System.String = ""
    End Class

    Public NotInheritable Class BoundOutputFile
        Public Property SourcePath As System.String = ""
        Public Property SourceSha256 As System.String = ""
        Public Property Binding As OutputNameBinding
        Public Property Evidence As System.String = ""
    End Class

    Public NotInheritable Class OutputNameBinder
        Private Sub New()
        End Sub

        Public Shared Sub ValidateBasename(name As System.String)
            If System.String.IsNullOrWhiteSpace(name) OrElse name = "." OrElse name = ".." OrElse
               name.EndsWith(".", System.StringComparison.Ordinal) OrElse name.EndsWith(" ", System.StringComparison.Ordinal) OrElse
               name.Any(Function(ch) System.Char.IsControl(ch) OrElse "<>:""/\|?*".IndexOf(ch) >= 0) Then
                Throw New System.IO.InvalidDataException("Invalid output basename: " & If(name, "(null)"))
            End If
            Dim stem = name.Split("."c)(0)
            Dim devices = New System.String() {"CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"}
            If devices.Contains(stem, System.StringComparer.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Reserved output basename: " & name)
        End Sub

        Public Shared Function FileHash(path As System.String) As System.String
            Dim info As New System.IO.FileInfo(path)
            If Not info.Exists OrElse (info.Attributes And (System.IO.FileAttributes.Directory Or System.IO.FileAttributes.ReparsePoint)) <> 0 Then
                Throw New System.IO.InvalidDataException("Output is not an existing regular file: " & path)
            End If
            Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Using sha = System.Security.Cryptography.SHA256.Create()
                    Return System.BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
                End Using
            End Using
        End Function

        Public Shared Sub ValidateBindings(bindings As System.Collections.Generic.IEnumerable(Of OutputNameBinding))
            If bindings Is Nothing Then Throw New System.ArgumentNullException(NameOf(bindings))
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim seen As New System.Collections.Generic.List(Of OutputNameBinding)()
            For Each binding In bindings
                If binding Is Nothing Then Throw New System.IO.InvalidDataException("Missing output binding.")
                ValidateBasename(binding.Name)
                ValidateIdentity(binding.LogicalDeliverableId, binding.OutputSlotId)
                If Not names.Add(binding.Name) Then Throw New System.IO.InvalidDataException("Duplicate output basename: " & binding.Name)
                If HasIdentity(binding.LogicalDeliverableId) AndAlso seen.Any(Function(prior) SameSlot(prior, binding.LogicalDeliverableId, binding.OutputSlotId)) Then
                    Throw New System.IO.InvalidDataException("More than one output name is bound to the same logical slot.")
                End If
                seen.Add(binding)
            Next
        End Sub

        Private Shared Function HasIdentity(value As System.String) As System.Boolean
            Return Not System.String.IsNullOrWhiteSpace(value)
        End Function

        Private Shared Sub ValidateIdentity(logicalId As System.String, slotId As System.String)
            If HasIdentity(logicalId) <> HasIdentity(slotId) Then Throw New System.IO.InvalidDataException("Incomplete output-slot binding.")
        End Sub

        Private Shared Function SameSlot(binding As OutputNameBinding, logicalId As System.String, slotId As System.String) As System.Boolean
            Return HasIdentity(logicalId) AndAlso HasIdentity(binding.LogicalDeliverableId) AndAlso
                System.String.Equals(binding.LogicalDeliverableId, logicalId, System.StringComparison.Ordinal) AndAlso
                System.String.Equals(binding.OutputSlotId, slotId, System.StringComparison.Ordinal)
        End Function

        Public Shared Function Resolve(candidates As System.Collections.Generic.IEnumerable(Of OutputDeliveryCandidate),
                                       previous As System.Collections.Generic.IEnumerable(Of OutputNameBinding),
                                       publications As OutputPublicationRegistry) As System.Collections.Generic.List(Of BoundOutputFile)
            If candidates Is Nothing Then Throw New System.ArgumentNullException(NameOf(candidates))
            Dim prior = If(previous, System.Linq.Enumerable.Empty(Of OutputNameBinding)()).ToList()
            ValidateBindings(prior)
            Dim result As New System.Collections.Generic.List(Of BoundOutputFile)()
            For Each candidate In candidates
                If candidate Is Nothing Then Throw New System.IO.InvalidDataException("Missing delivery candidate.")
                ValidateIdentity(candidate.LogicalDeliverableId, candidate.OutputSlotId)
                Dim full = System.IO.Path.GetFullPath(candidate.Path)
                Dim physicalName = System.IO.Path.GetFileName(full)
                ValidateBasename(physicalName)
                Dim digest = FileHash(full)
                Dim requested = If(publications Is Nothing, "", publications.GetRequestedName(full, digest))
                Dim declaredName = If(requested = "", "", requested.Split("/"c).Last())
                Dim slotMatch = prior.SingleOrDefault(Function(binding) SameSlot(binding, candidate.LogicalDeliverableId, candidate.OutputSlotId))
                ' A receipt outranks a transport basename: a collision suffix can itself
                ' equal another genuine output name. Never strip suffixes or match stems.
                Dim evidencedName = If(declaredName = "", physicalName, declaredName)
                Dim nameMatch = prior.SingleOrDefault(Function(binding) System.String.Equals(binding.Name, evidencedName, System.StringComparison.OrdinalIgnoreCase))
                Dim matched = slotMatch
                Dim reason = "logical_slot"
                If slotMatch IsNot Nothing Then
                    If declaredName <> "" AndAlso nameMatch IsNot Nothing AndAlso Not System.String.Equals(slotMatch.Name, nameMatch.Name, System.StringComparison.OrdinalIgnoreCase) Then
                        Throw New System.IO.InvalidDataException("Publication name conflicts with the retained logical output slot.")
                    End If
                Else
                    matched = nameMatch
                    reason = If(declaredName = "", "exact_delivered_name", "verified_publication_name")
                    If matched IsNot Nothing AndAlso HasIdentity(matched.LogicalDeliverableId) Then
                        Throw New System.IO.InvalidDataException("A carried output changed or omitted its logical slot identity: " & matched.Name)
                    End If
                End If
                Dim name = If(matched Is Nothing, evidencedName, matched.Name)
                ValidateBasename(name)
                If Not System.String.Equals(System.IO.Path.GetExtension(name), System.IO.Path.GetExtension(physicalName), System.StringComparison.OrdinalIgnoreCase) Then
                    Throw New System.IO.InvalidDataException("Output-name binding cannot change the file format: " & name)
                End If
                result.Add(New BoundOutputFile With {.SourcePath = full, .SourceSha256 = digest,
                    .Binding = New OutputNameBinding With {.Name = name, .LogicalDeliverableId = If(candidate.LogicalDeliverableId, ""), .OutputSlotId = If(candidate.OutputSlotId, "")}, .Evidence = reason})
            Next
            ValidateBindings(result.Select(Function(output) output.Binding))
            For Each binding In prior
                If Not result.Any(Function(output) System.String.Equals(output.Binding.Name, binding.Name, System.StringComparison.OrdinalIgnoreCase)) Then
                    Throw New System.IO.InvalidDataException("No authorized current output matches the retained binding: " & binding.Name & ". No stale input or filename guess was substituted.")
                End If
            Next
            Return result
        End Function
    End Class
End Namespace
