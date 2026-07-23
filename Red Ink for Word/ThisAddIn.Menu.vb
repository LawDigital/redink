' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Menu.vb
' Purpose: Manages Word context menu customization for Red Ink add-in functionality.
'          Adds Red Ink menu items to relevant Word context menus, configures keyboard
'          shortcuts, and handles menu cleanup.
'
' Architecture:
'  - Context Menu Integration: Iterates through Word's CommandBars collection and adds
'    Red Ink menu items to relevant popup menus (text, tables, fields, etc.).
'  - Dynamic Menu Building: Creates hierarchical menus based on INI configuration,
'    including translation options, correction features, improvement tools, and Word helpers.
'  - Keyboard Shortcut Assignment: Parses INI_ShortcutsWordExcel configuration string
'    and assigns keyboard shortcuts to menu commands via Word's KeyBindings API.
'  - VBA Bridge Integration: Menu commands trigger VBA macros (OnAction property) that
'    invoke BridgeSubs methods to execute Red Ink functionality.
'  - Menu State Management: Tracks menu addition state (MenusAdded) and handles cleanup
'    of outdated menu items during updates.
'  - Shortcut Key Parsing: Converts text-based shortcut definitions (e.g., "CTRL-SHIFT-F5")
'    into Word KeyCode values using BuildKeyCodeFromText function.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Adds Red Ink context menu to relevant Word context menus.
    ''' Initializes the ribbon, validates configuration, and dynamically creates menu structure.
    ''' </summary>
    ''' <remarks>
    ''' Requires VBA module to be working and INI configuration to be loaded.
    ''' Menu items are temporary and recreated each time to reflect current settings.
    ''' </remarks>
    Public Sub AddContextMenu()
        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

        If MenusAdded Then Return

        ' Remove existing context menus from relevant context menus
        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Return

        If Not VBAModuleWorking() Then Return

        If INIloaded = False Then Return

        MenusAdded = True

        ' List of relevant context menus where Red Ink menu should appear
        Dim contextMenus As String() = {
            "Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
            "Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
            "Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
            "Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
            "Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
            "Business Card", "Equation Popup", "WordArt Context Menu",
            "Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
            "Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
        }
        Dim application As Word.Application = Globals.ThisAddIn.Application

        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Check if the menu already exists
                    If Not ContextMenuExists(cb, RIMenu) Then
                        Dim myControl As CommandBarPopup = Nothing
                        Try
                            myControl = CType(cb.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
                        Catch ex As System.Exception
                            ' Silently handle errors when menu cannot be added (e.g., protected menus)
                            Continue For
                        End Try

                        If myControl IsNot Nothing Then
                            myControl.Caption = RIMenu
                            myControl.Visible = True
                            myControl.Enabled = True

                            ' Add submenu items
                            AddSubMenuItems(myControl)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Checks if a context menu already contains a popup menu item with the specified name.
    ''' </summary>
    ''' <param name="cb">The CommandBar to check.</param>
    ''' <param name="menuName">The caption of the menu item to search for.</param>
    ''' <returns>True if the menu exists, otherwise False.</returns>
    Private Function ContextMenuExists(cb As CommandBar, menuName As String) As Boolean
        For Each ctrl As CommandBarControl In cb.Controls
            If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = menuName Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Adds submenu items to the Red Ink context menu popup.
    ''' Creates translation, correction, improvement, and helper tool menu items.
    ''' Assigns keyboard shortcuts from INI configuration.
    ''' </summary>
    ''' <param name="myControl">The CommandBarPopup to add items to.</param>
    ''' <remarks>
    ''' OnAction properties reference VBA macros that bridge to BridgeSubs methods.
    ''' Menu structure includes conditional items based on configuration (INI_Language1/2, INI_SecondAPI, INI_DoMarkupWord).
    ''' </remarks>
    Public Sub AddSubMenuItems(myControl As CommandBarPopup)
        Try
            Dim subControl As CommandBarButton
            Dim wordHelpersMenu As CommandBarPopup
            Dim improveMenu As CommandBarPopup
            Dim subSubControl As CommandBarButton
            Dim shortcutsArray() As String
            Dim shortcutPair() As String
            Dim shortcutDict As New Dictionary(Of String, String)
            Dim i As Integer

            ' Parse the shortcuts from INI_ShortcutsWordExcel (format: "Menu Name=CTRL-F1;Other Menu=SHIFT-F2")
            shortcutsArray = INI_ShortcutsWordExcel.Split(";"c)

            ' Populate the dictionary with menu name/shortcut pairs
            For i = 0 To shortcutsArray.Length - 1
                If shortcutsArray(i).Contains("=") Then
                    shortcutPair = shortcutsArray(i).Split("="c)
                    shortcutDict(shortcutPair(0).Trim()) = shortcutPair(1).Trim()
                End If
            Next
            myControl.Visible = True

            ' Add menu items and assign shortcuts
            ' OnAction properties reference VBA macros that invoke BridgeSubs methods

            ' Translation menu items (conditional based on configured languages)
            If Not String.IsNullOrWhiteSpace(INI_Language1) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language1
                subControl.FaceId = 6112 ' Translation icon
                subControl.Visible = True
                subControl.OnAction = "CallInLanguage1"
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If

            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112 ' Translation icon
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If

            ' "To Other" translation option (always available)
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112 ' Translation icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Correction feature (with optional markup mode)
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329 ' Checkmark icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create "Improve" submenu with style options
            improveMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            improveMenu.Caption = "Improve"

            ' Improve submenu items
            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Improve" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallImprove"
            subSubControl.FaceId = 329 ' Checkmark icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "No Filler Words" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallNoFillers"
            subSubControl.FaceId = 4242 ' Filter icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Friendly" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallFriendly"
            subSubControl.FaceId = 59 ' Smiley icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(improveMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "More Convincing" & If(INI_DoMarkupWord, " (Markup)", "")
            subSubControl.OnAction = "CallConvincing"
            subSubControl.FaceId = 343 ' Lightbulb icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            ' Additional text manipulation features
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292 ' Compress icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502 ' Privacy icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party" & If(INI_DoMarkupWord, " (Markup)", "")
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327 ' Swap icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Summarize"
            subControl.OnAction = "CallSummarize"
            subControl.FaceId = 602 ' Document icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Freestyle prompt feature (primary model)
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346 ' Custom prompt icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Freestyle prompt feature (secondary model, if configured)
            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (2nd)"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346 ' Custom prompt icon
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
                End If
            End If

            ' Context search feature
            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Context Search"
            subControl.OnAction = "CallContextSearch"
            subControl.FaceId = 46 ' Search icon
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut " & shortcutDict(subControl.Caption)
            End If

            ' Create "Word helpers" submenu with utility functions
            wordHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            wordHelpersMenu.Caption = "Word helpers"

            ' Word helpers submenu items
            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Self-Compare Selection"
            subSubControl.OnAction = "CallCompareSelectionHalves"
            subSubControl.FaceId = 304 ' Compare icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Accept Format Changes"
            subSubControl.OnAction = "CallAcceptFormatting"
            subSubControl.FaceId = 161 ' Accept icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Markup Time Span"
            subSubControl.OnAction = "CallCalculateUserMarkupTimeSpan"
            subSubControl.FaceId = 33 ' Clock icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288 ' Find/Replace icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(wordHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Import Text File"
            subSubControl.OnAction = "CallImportTextFile"
            subSubControl.FaceId = 2311 ' Import icon
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut " & shortcutDict(subSubControl.Caption)
            End If

            ' Assign keyboard shortcuts if configured
            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then
                ' Assign shortcuts for main menu items
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other", "CallInOther", shortcutDict)
                AssignShortcut("Correct (Markup)", "CallCorrect", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Improve (Markup)", "CallImprove", shortcutDict)
                AssignShortcut("Improve", "CallImprove", shortcutDict)
                AssignShortcut("No Filler Words (Markup)", "CallNoFillers", shortcutDict)
                AssignShortcut("No Filler Words", "CallNoFillers", shortcutDict)
                AssignShortcut("More Friendly (Markup)", "CallFriendly", shortcutDict)
                AssignShortcut("More Friendly", "CallFriendly", shortcutDict)
                AssignShortcut("More Convincing (Markup)", "CallConvincing", shortcutDict)
                AssignShortcut("More Convincing", "CallConvincing", shortcutDict)
                AssignShortcut("Shorten (Markup)", "CallShorten", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize (Markup)", "CallAnonymize", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party (Markup)", "CallSwitchParty", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Summarize", "CallSummarize", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)
                AssignShortcut("Context Search", "CallContextSearch", shortcutDict)

                ' Assign shortcuts for secondary API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (2nd)", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for Word helpers submenu
                AssignShortcut("Self-Compare Selection", "CallCompareSelectionHalves", shortcutDict)
                AssignShortcut("Accept Format Changes", "CallAcceptFormatting", shortcutDict)
                AssignShortcut("Markup Time Span", "CallCalculateUserMarkupTimeSpan", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Import Text File", "CallImportTextFile", shortcutDict)
            End If

        Catch ex As System.Exception
            ' Silently handle errors during menu construction
            ' Prevents menu building failures from disrupting Word functionality
        End Try
    End Sub

    ''' <summary>
    ''' Assigns a keyboard shortcut to a VBA macro command.
    ''' </summary>
    ''' <param name="controlName">The menu item caption to look up in the shortcut dictionary.</param>
    ''' <param name="macro">The VBA macro name to bind the shortcut to.</param>
    ''' <param name="shortcutDict">Dictionary mapping menu captions to shortcut key strings.</param>
    ''' <remarks>
    ''' KeyBindings are added to the Normal template (CustomizationContext).
    ''' Shortcut keys are built using BuildKeyCodeFromText to convert text format to Word KeyCode.
    ''' </remarks>
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCode As Long
        Try
            ' Check if there is a shortcut assigned for this menu item
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Return ' No shortcut assigned
            End If

            ' Build KeyCode from shortcutKey text (e.g., "CTRL-SHIFT-F5")
            keyCode = BuildKeyCodeFromText(shortcutKey)

            If keyCode > 0 Then
                Globals.ThisAddIn.Application.CustomizationContext = Globals.ThisAddIn.Application.NormalTemplate
                Globals.ThisAddIn.Application.KeyBindings.Add(KeyCode:=CInt(keyCode), KeyCategory:=WdKeyCategory.wdKeyCategoryMacro, Command:=macro)
            End If
        Catch ex As System.Exception
            ' Silently handle errors (e.g., shortcut already in use, invalid key combination)
        End Try
    End Sub

    ''' <summary>
    ''' Converts a text-based shortcut key definition to a Word KeyCode value.
    ''' </summary>
    ''' <param name="shortcutKey">Shortcut string (e.g., "CTRL-SHIFT-F5", "ALT-A").</param>
    ''' <returns>Word KeyCode as Long, or 0 if parsing fails.</returns>
    ''' <remarks>
    ''' Supports modifier keys (CTRL, SHIFT, ALT), function keys (F1-F12), 
    ''' digits (0-9), letters (A-Z), navigation keys (arrow keys, HOME, END, etc.),
    ''' and special keys (ESC, TAB, SPACE, etc.).
    ''' Uses Word's Application.BuildKeyCode method to combine multiple key values.
    ''' </remarks>
    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As Long
        Dim parts() As String
        Dim keysCollection As New List(Of Integer)()
        Dim keyCode As Long = 0

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    ' Modifier keys
                    Case "CTRL"
                        keysCollection.Add(WdKey.wdKeyControl)
                    Case "SHIFT"
                        keysCollection.Add(WdKey.wdKeyShift)
                    Case "ALT"
                        keysCollection.Add(WdKey.wdKeyAlt)

                    ' Digit keys
                    Case "0"
                        keysCollection.Add(WdKey.wdKey0)
                    Case "1"
                        keysCollection.Add(WdKey.wdKey1)
                    Case "2"
                        keysCollection.Add(WdKey.wdKey2)
                    Case "3"
                        keysCollection.Add(WdKey.wdKey3)
                    Case "4"
                        keysCollection.Add(WdKey.wdKey4)
                    Case "5"
                        keysCollection.Add(WdKey.wdKey5)
                    Case "6"
                        keysCollection.Add(WdKey.wdKey6)
                    Case "7"
                        keysCollection.Add(WdKey.wdKey7)
                    Case "8"
                        keysCollection.Add(WdKey.wdKey8)
                    Case "9"
                        keysCollection.Add(WdKey.wdKey9)

                    ' Function keys
                    Case "F1"
                        keysCollection.Add(WdKey.wdKeyF1)
                    Case "F2"
                        keysCollection.Add(WdKey.wdKeyF2)
                    Case "F3"
                        keysCollection.Add(WdKey.wdKeyF3)
                    Case "F4"
                        keysCollection.Add(WdKey.wdKeyF4)
                    Case "F5"
                        keysCollection.Add(WdKey.wdKeyF5)
                    Case "F6"
                        keysCollection.Add(WdKey.wdKeyF6)
                    Case "F7"
                        keysCollection.Add(WdKey.wdKeyF7)
                    Case "F8"
                        keysCollection.Add(WdKey.wdKeyF8)
                    Case "F9"
                        keysCollection.Add(WdKey.wdKeyF9)
                    Case "F10"
                        keysCollection.Add(WdKey.wdKeyF10)
                    Case "F11"
                        keysCollection.Add(WdKey.wdKeyF11)
                    Case "F12"
                        keysCollection.Add(WdKey.wdKeyF12)

                    ' Navigation and special keys
                    Case "LEFT"
                        keysCollection.Add(CustomWdKey.wdKeyLeft)
                    Case "RIGHT"
                        keysCollection.Add(CustomWdKey.wdKeyRight)
                    Case "UP"
                        keysCollection.Add(CustomWdKey.wdKeyUp)
                    Case "DOWN"
                        keysCollection.Add(CustomWdKey.wdKeyDown)
                    Case "HOME"
                        keysCollection.Add(WdKey.wdKeyHome)
                    Case "END"
                        keysCollection.Add(WdKey.wdKeyEnd)
                    Case "PAGEUP"
                        keysCollection.Add(WdKey.wdKeyPageUp)
                    Case "PAGEDOWN"
                        keysCollection.Add(WdKey.wdKeyPageDown)
                    Case "ESC"
                        keysCollection.Add(WdKey.wdKeyEsc)
                    Case "TAB"
                        keysCollection.Add(WdKey.wdKeyTab)
                    Case "BACKSPACE"
                        keysCollection.Add(WdKey.wdKeyBackspace)
                    Case "DELETE"
                        keysCollection.Add(WdKey.wdKeyDelete)
                    Case "INSERT"
                        keysCollection.Add(WdKey.wdKeyInsert)
                    Case "SPACE"
                        keysCollection.Add(CustomWdKey.wdKeySpace)

                    ' Letter keys (A-Z)
                    Case "A" To "Z"
                        keysCollection.Add([Enum].Parse(GetType(WdKey), "wdKey" & part.Trim().ToUpper()))
                    Case Else
                        ' Unknown key - return 0 to indicate parsing failure
                        Return 0
                End Select
            Next

            ' Build the KeyCode using Application.BuildKeyCode (supports 1-4 keys)
            Select Case keysCollection.Count
                Case 1
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0))
                Case 2
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1))
                Case 3
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2))
                Case 4
                    keyCode = Globals.ThisAddIn.Application.BuildKeyCode(keysCollection(0), keysCollection(1), keysCollection(2), keysCollection(3))
                Case Else
                    ' Invalid key combination
                    Return 0
            End Select

            Return keyCode

        Catch ex As System.Exception
            ' Handle parsing errors gracefully
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Removes old Red Ink context menus from relevant Word context menus.
    ''' </summary>
    ''' <remarks>
    ''' Called during menu update to ensure clean state before adding new menus.
    ''' Uses RIMenu constant to identify current menu version.
    ''' </remarks>
    Public Sub RemoveOldContextMenu()
        Dim application As Word.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {
            "Text", "Spelling", "Grammar", "Grammar (2)", "Linked Text", "Lists", "Headings", "Rotate Text", "Table Text",
            "Footnotes", "Endnotes", "Frames", "Fields", "Form Fields", "Display Fields", "Field Display List Numbers", "Field AutoText",
            "Comment", "Track Changes", "Track Changes Indicator", "Hyperlink Context Menu",
            "Table Cells", "Whole Table", "Linked Table", "Table Lists", "Table Pictures",
            "Inline Picture", "Floating Picture", "OLE Object", "ActiveX Control", "Inline ActiveX Control",
            "Business Card", "Equation Popup", "WordArt Context Menu",
            "Drop Caps", "Font Popup", "Font Paragraph", "Format consistency",
            "Format Inspector Popup in Normal Mode", "Format Inspector Popup in Compare Mode", "AutoSignature Popup"
        }

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = RIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Silently handle errors (e.g., menu already deleted, protected menu)
                            End Try
                        End If
                    Next
                End If
            End If
        Next
    End Sub

End Class
