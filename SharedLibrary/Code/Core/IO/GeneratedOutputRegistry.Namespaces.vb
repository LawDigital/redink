' Shared artifact namespaces are recognized by an immutable protocol marker.
' No per-document registry entries are needed, including for another producer's output.

' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: GeneratedOutputRegistry.Namespaces.vb
' Purpose:
'   Immutable protocol markers for recognizing shared generated-artifact namespaces.
'
' Architecture / Function:
'   Identifies marked output namespaces without per-document registrations or directory-
'   name guessing.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace SharedLibrary

    Public NotInheritable Partial Class GeneratedOutputRegistry

        Public Const ArtifactNamespaceName As System.String = ".redink-sa"
        Public Const ArtifactNamespaceMarkerName As System.String = ".redink-sa.namespace.json"
        Private Const ArtifactNamespaceFormat As System.String = "redink-semantic-archive-namespace-v1"
        Private Shared ReadOnly NamespaceMarkerBytes As System.Byte() =
            New System.Text.UTF8Encoding(False, True).GetBytes("{""format"":""" & ArtifactNamespaceFormat & """}" & System.Environment.NewLine)

        ''' <summary>
        ''' Claims a new, empty artifact namespace before any payload is written.
        ''' Existing source folders are never adopted or excluded merely by name.
        ''' The marker is content-free and may inherit the namespace's read rights.
        ''' </summary>
        Public Shared Sub EnsureNamespaceMarker(namespaceDirectory As System.String)
            Dim physical As System.String = NormalizePhysical(namespaceDirectory)
            RequireNamespaceDirectory(physical)
            Dim marker As System.String = System.IO.Path.Combine(physical, ArtifactNamespaceMarkerName)
            If ReadNamespaceMarker(physical) Then Return
            ' Callers create the namespace with their chosen security descriptor.
            ' Creating or changing the original parent directory is not required here.
            If Not System.IO.Directory.Exists(physical) Then Throw New System.IO.DirectoryNotFoundException("The generated artifact namespace has not been created.")
            Using entries As System.Collections.Generic.IEnumerator(Of System.String) = System.IO.Directory.EnumerateFileSystemEntries(physical).GetEnumerator()
                If entries.MoveNext() Then Throw New System.IO.IOException("The unmarked .redink-sa directory contains existing files and cannot be adopted as generated output.")
            End Using
            Dim temporary As System.String = System.IO.Path.Combine(physical, ".ns-" & System.Guid.NewGuid().ToString("N") & ".tmp")
            RequireNamespacePathBudget(marker)
            RequireNamespacePathBudget(temporary)
            Try
                Using stream As New System.IO.FileStream(temporary, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    stream.Write(NamespaceMarkerBytes, 0, NamespaceMarkerBytes.Length)
                    stream.Flush(True)
                End Using
                Try
                    System.IO.File.Move(temporary, marker)
                Catch ex As System.IO.IOException
                    ' Another publisher may have installed the exact same marker.
                    ' An invalid, inaccessible or absent winner is never accepted.
                    If Not ReadNamespaceMarker(physical) Then Throw
                End Try
                If Not ReadNamespaceMarker(physical) Then Throw New System.IO.IOException("The generated artifact namespace could not be verified after publication.")
            Finally
                If System.IO.File.Exists(temporary) Then
                    Try
                        System.IO.File.Delete(temporary)
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("GeneratedOutputRegistry: namespace marker staging cleanup deferred: " & ex.Message)
                    End Try
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Exact shared namespace discovery, including namespaces made by another user.
        ''' An inaccessible or malformed marker is an unknown decision and throws.
        ''' </summary>
        Public Shared Function IsMarkedNamespacePath(candidatePath As System.String) As System.Boolean
            Return IsMarkedNamespacePathCore(NormalizePhysical(candidatePath))
        End Function

        Private Shared Function IsMarkedNamespacePathCore(physicalPath As System.String) As System.Boolean
            Dim current As System.String = physicalPath
            Do While Not System.String.IsNullOrWhiteSpace(current)
                If System.String.Equals(System.IO.Path.GetFileName(current), ArtifactNamespaceName, System.StringComparison.OrdinalIgnoreCase) AndAlso
                    ReadNamespaceMarker(current) Then Return True
                Dim parent As System.String = System.IO.Path.GetDirectoryName(current)
                If System.String.Equals(parent, current, System.StringComparison.OrdinalIgnoreCase) Then Exit Do
                current = parent
            Loop
            Return False
        End Function

        Private Shared Function ReadNamespaceMarker(namespaceDirectory As System.String) As System.Boolean
            Dim marker As System.String = System.IO.Path.Combine(namespaceDirectory, ArtifactNamespaceMarkerName)
            Try
                Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(marker)
                If (attributes And (System.IO.FileAttributes.ReparsePoint Or System.IO.FileAttributes.Directory)) <> 0 Then
                    Throw New System.IO.InvalidDataException("The generated artifact namespace marker is not a regular file.")
                End If
            Catch ex As System.IO.FileNotFoundException
                Return False
            Catch ex As System.IO.DirectoryNotFoundException
                Return False
            End Try
            RequireNamespaceDirectory(namespaceDirectory)
            ' Pin the marker while reading; a writer cannot replace or mutate it.
            Using stream As New System.IO.FileStream(marker, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                If stream.Length <= 0 OrElse stream.Length > 256 Then Throw New System.IO.InvalidDataException("The generated artifact namespace marker is invalid.")
                Using reader As New System.IO.StreamReader(stream, New System.Text.UTF8Encoding(False, True), False)
                    Dim actual As System.String = reader.ReadToEnd()
                    Dim expected As System.String = "{""format"":""" & ArtifactNamespaceFormat & """}"
                    If Not System.String.Equals(actual.TrimEnd(ChrW(10), ChrW(13)), expected, System.StringComparison.Ordinal) Then
                        Throw New System.IO.InvalidDataException("The generated artifact namespace marker has an unsupported format.")
                    End If
                End Using
            End Using
            Return True
        End Function

        Private Shared Sub RequireNamespaceDirectory(path As System.String)
            If Not System.String.Equals(System.IO.Path.GetFileName(path), ArtifactNamespaceName, System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.ArgumentException("The exact generated artifact namespace name is required.", NameOf(path))
            End If
            Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(path)
            If (attributes And System.IO.FileAttributes.Directory) = 0 OrElse (attributes And System.IO.FileAttributes.ReparsePoint) <> 0 Then
                Throw New System.IO.IOException("The generated artifact namespace must be a regular directory.")
            End If
        End Sub

        Private Shared Sub RequireNamespacePathBudget(path As System.String)
            ' The shared Office/export pipeline retains a conservative legacy Win32
            ' path budget; a long-path prefix alone does not update every consumer.
            SemanticArchivePathGuard.RequireWindowsCompatiblePath(path)
        End Sub

    End Class

End Namespace
