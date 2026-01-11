' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.MyStyleHelpers.vb
' Purpose: Reads a MyStyle prompt definition file and lets the user select one
'          applicable prompt for a given calling application ("Word" / "Outlook").
'
' File Format (INI-like, line-based):
'  - Empty lines are ignored.
'  - Comment lines starting with ";" are ignored.
'  - Supported entry formats:
'      * "App|Title|Prompt"  (App: "Word", "Outlook", or "All")
'      * "Title|Prompt"      (legacy format; treated as "All|Title|Prompt")
'
' Behavior:
'  - Filters entries by application ("All" or the normalized calling application).
'  - Sorts entries by title (case-insensitive).
'  - Displays entries as "Title (App)" and enforces unique display labels.
'  - Optional "None" entry (ID 0) returns "NONE".
'  - Uses SharedMethods.SelectValue to show the picker UI and returns the selected
'    prompt string; returns "ERROR" on validation/processing errors.
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Helpers for loading MyStyle prompt definitions and showing a selection UI.
        ''' </summary>
        Public NotInheritable Class MyStyleHelpers

            ''' <summary>
            ''' Reads a MyStyle prompt definition file, filters prompts for the specified application,
            ''' shows a picker UI, and returns the selected prompt text.
            ''' </summary>
            ''' <param name="iniPath">Full path to the MyStyle prompt definition file.</param>
            ''' <param name="callingApplication">Application name ("Word" or "Outlook").</param>
            ''' <param name="defaultValue">Default selected item ID passed to the picker.</param>
            ''' <param name="promptText">Prompt text shown in the picker UI.</param>
            ''' <param name="headerText">Optional header text shown in the picker UI.</param>
            ''' <param name="AddNone">If <c>True</c>, adds a "None" option returning "NONE" (ID 0).</param>
            ''' <returns>
            ''' The selected prompt text, or "NONE" if "None" was selected or no applicable items exist,
            ''' or "ERROR" on validation/processing errors.
            ''' </returns>
            Public Shared Function SelectPromptFromMyStyle(ByVal iniPath As System.String,
                                                   ByVal callingApplication As System.String,
                                                   Optional ByVal defaultValue As System.Int32 = 0,
                                                   Optional ByVal promptText As System.String = "Please choose …",
                                                   Optional ByVal headerText As System.String = Nothing,
                                                   Optional ByVal AddNone As Boolean = True) As System.String
                Try
                    ' --- Validate inputs ---
                    If iniPath Is Nothing OrElse iniPath.Trim().Length = 0 Then
                        ShowCustomMessageBox($"Invalid MyStyle prompt file path ({iniPath}).")
                        Return "ERROR"
                    End If

                    If callingApplication Is Nothing OrElse callingApplication.Trim().Length = 0 Then
                        ShowCustomMessageBox("Invalid calling application (expected 'Word' or 'Outlook').")
                        Return "ERROR"
                    End If

                    Dim appNorm As System.String = NormalizeAppName(callingApplication)
                    If appNorm Is Nothing Then
                        ShowCustomMessageBox("Unknown application '" & callingApplication & "'. Use 'Word' or 'Outlook'.")
                        Return "ERROR"
                    End If

                    If System.IO.File.Exists(iniPath) = False Then
                        ShowCustomMessageBox("MyStyle prompt file not found at: " & iniPath)
                        Return "ERROR"
                    End If

                    ' --- Parse file into entries ---
                    Dim entries As System.Collections.Generic.List(Of MyStyleEntry) = New System.Collections.Generic.List(Of MyStyleEntry)()
                    For Each raw As System.String In System.IO.File.ReadLines(iniPath)
                        If raw Is Nothing Then
                            Continue For
                        End If

                        Dim line As System.String = raw.Trim()
                        If line.Length = 0 Then
                            Continue For
                        End If
                        If line.StartsWith(";", System.StringComparison.Ordinal) Then
                            Continue For
                        End If

                        ' Parse into App|Title|Prompt (legacy Title|Prompt → All|Title|Prompt)
                        Dim p1 As System.Int32 = line.IndexOf("|"c)
                        If p1 < 0 Then
                            Continue For
                        End If
                        Dim p2 As System.Int32 = line.IndexOf("|"c, p1 + 1)

                        Dim app As System.String
                        Dim title As System.String
                        Dim prompt As System.String

                        If p2 >= 0 Then
                            app = line.Substring(0, p1).Trim()
                            title = line.Substring(p1 + 1, p2 - (p1 + 1)).Trim()
                            prompt = line.Substring(p2 + 1).Trim()
                        Else
                            app = "All"
                            title = line.Substring(0, p1).Trim()
                            prompt = line.Substring(p1 + 1).Trim()
                        End If

                        If title.Length = 0 OrElse prompt.Length = 0 Then
                            Continue For
                        End If

                        Dim appForEntry As System.String = NormalizeAppName(app)
                        If appForEntry Is Nothing Then
                            Continue For
                        End If

                        If appForEntry.Equals("All", System.StringComparison.OrdinalIgnoreCase) _
                   OrElse appForEntry.Equals(appNorm, System.StringComparison.OrdinalIgnoreCase) Then
                            entries.Add(New MyStyleEntry With {.App = appForEntry, .Title = title, .Prompt = prompt})
                        End If
                    Next

                    ' --- Build List(Of SharedMethods.SelectionItem) ---
                    Dim items As System.Collections.Generic.List(Of SharedMethods.SelectionItem) =
                New System.Collections.Generic.List(Of SharedMethods.SelectionItem)()

                    ' ID → Prompt map
                    Dim idToPrompt As System.Collections.Generic.Dictionary(Of System.Int32, System.String) =
                New System.Collections.Generic.Dictionary(Of System.Int32, System.String)()

                    ' Ensure unique display strings
                    Dim seenDisplays As System.Collections.Generic.HashSet(Of System.String) =
                New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

                    If AddNone And entries.Count > 0 Then
                        ' add NONE (ID = 0)
                        items.Add(New SharedMethods.SelectionItem("None", 0))
                        seenDisplays.Add("None")
                        idToPrompt(0) = "NONE"
                    End If

                    If entries.Count > 0 Then
                        entries.Sort(Function(a As MyStyleEntry, b As MyStyleEntry) _
                    System.String.Compare(a.Title, b.Title, System.StringComparison.OrdinalIgnoreCase))

                        Dim nextId As System.Int32 = 1
                        For Each e As MyStyleEntry In entries
                            Dim display As System.String = e.Title & " (" & e.App & ")"
                            display = MakeUniqueDisplay(display, seenDisplays)

                            items.Add(New SharedMethods.SelectionItem(display, nextId))
                            idToPrompt(nextId) = e.Prompt
                            nextId += 1
                        Next
                    End If

                    If items.Count = 0 Then
                        ShowCustomMessageBox($"No styles applicable for {appNorm} found in your MyStyle prompt file ({iniPath}).",
                                                                                                         extraButtonText:="Edit MyStyle prompt file",
                                                            extraButtonAction:=Sub()
                                                                                   ShowTextFileEditor(iniPath, "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                                                                               End Sub)

                        Return "NONE"
                    End If

                    ' --- Show picker (uses your SharedMethods.SelectValue) ---
                    Dim chosenId As System.Int32 = SharedMethods.SelectValue(items, defaultValue, promptText, headerText)

                    If chosenId = 0 Then
                        Return "NONE"
                    End If

                    Dim outPrompt As System.String = Nothing
                    If idToPrompt.TryGetValue(chosenId, outPrompt) Then
                        Return outPrompt
                    End If

                    ShowCustomMessageBox("Unexpected selection result.")
                    Return "ERROR"

                Catch ex As System.Exception
                    ShowCustomMessageBox($"Error reading the MyStyle prompt file ({iniPath}): " & ex.Message)
                    Return "ERROR"
                End Try
            End Function

            ' ------- Helpers (Shared) -------

            ''' <summary>
            ''' Normalizes an application name to "Word", "Outlook", or "All".
            ''' Returns <c>Nothing</c> for unknown/empty inputs.
            ''' </summary>
            ''' <param name="input">Raw application name read from file or passed in by the caller.</param>
            ''' <returns>"Word", "Outlook", "All", or <c>Nothing</c>.</returns>
            Private Shared Function NormalizeAppName(ByVal input As System.String) As System.String
                If input Is Nothing Then
                    Return Nothing
                End If
                Dim s As System.String = input.Trim()
                If s.Length = 0 Then
                    Return Nothing
                End If
                If s.Equals("Word", System.StringComparison.OrdinalIgnoreCase) Then
                    Return "Word"
                ElseIf s.Equals("Outlook", System.StringComparison.OrdinalIgnoreCase) Then
                    Return "Outlook"
                ElseIf s.Equals("All", System.StringComparison.OrdinalIgnoreCase) Then
                    Return "All"
                End If
                Return Nothing
            End Function

            ''' <summary>
            ''' Ensures a unique display label by appending a numeric suffix if needed.
            ''' </summary>
            ''' <param name="display">The proposed display label.</param>
            ''' <param name="seen">Case-insensitive set used to track already-used display labels.</param>
            ''' <returns>The original display label if unique; otherwise a suffixed variant.</returns>
            Private Shared Function MakeUniqueDisplay(ByVal display As System.String,
                                              ByVal seen As System.Collections.Generic.HashSet(Of System.String)) As System.String
                If seen.Contains(display) = False Then
                    seen.Add(display)
                    Return display
                End If
                Dim n As System.Int32 = 2
                While True
                    Dim candidate As System.String = display & " [" & n.ToString(System.Globalization.CultureInfo.InvariantCulture) & "]"
                    If seen.Contains(candidate) = False Then
                        seen.Add(candidate)
                        Return candidate
                    End If
                    n += 1
                End While
            End Function

            ''' <summary>
            ''' Container for one parsed MyStyle entry (application, title, prompt).
            ''' </summary>
            Private NotInheritable Class MyStyleEntry
                ''' <summary>
                ''' Normalized application name ("Word", "Outlook", or "All").
                ''' </summary>
                Public Property App As System.String

                ''' <summary>
                ''' Entry title shown to the user.
                ''' </summary>
                Public Property Title As System.String

                ''' <summary>
                ''' Prompt text returned when this entry is selected.
                ''' </summary>
                Public Property Prompt As System.String
            End Class

        End Class


    End Class

End Namespace