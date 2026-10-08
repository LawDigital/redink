' Part of Red Ink (AskInky). Standalone WinForms message loop, not a VSTO add-in.
Friend Module Program
    <System.STAThread>
    Public Sub Main()
        System.Windows.Forms.Application.SetHighDpiMode(System.Windows.Forms.HighDpiMode.PerMonitorV2)
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        System.Windows.Forms.Application.Run(New AssistantContext())
    End Sub
End Module
