' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: Program.vb
' Purpose: Main Windows entry point and application initialization.
' Version: V091026 (9 October 2026). Update on each dated delivery.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Third-party elements: Markdig (BSD-2-Clause), Microsoft.Web.WebView2 (Microsoft license),
' .NET runtime components (applicable Microsoft/.NET licenses), adapted Lucide/Feather SVG paths
' (ISC/MIT) and adapted Office Web styling; full notices: LICENSE.txt.
'
' Part of Red Ink (AskInky). Standalone WinForms message loop, not a VSTO add-in.
Friend Module Program
    Public Const Version As String = "V091026"
    <System.STAThread>
    Public Sub Main()
        System.Windows.Forms.Application.SetHighDpiMode(System.Windows.Forms.HighDpiMode.PerMonitorV2)
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        System.Windows.Forms.Application.Run(New AssistantContext())
    End Sub
End Module
