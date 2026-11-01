' Part of Red Ink (Prompt Picker). New standalone Windows tray application.
Friend Module Program
    <System.STAThread>
    Public Sub Main()
        Dim created As Boolean
        Using instance As New System.Threading.Mutex(True, "Local\LawDigital.RedInk.PromptPicker", created)
            If Not created Then
                System.Windows.Forms.MessageBox.Show("Red Ink Prompt Picker läuft bereits.", "Red Ink Prompt Picker", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Return
            End If
            AddHandler System.AppDomain.CurrentDomain.UnhandledException,
                Sub(sender As Object, e As System.UnhandledExceptionEventArgs)
                    AppLog.Write("unhandled", TryCast(e.ExceptionObject, System.Exception))
                End Sub
            System.Windows.Forms.Application.SetHighDpiMode(System.Windows.Forms.HighDpiMode.PerMonitorV2)
            System.Windows.Forms.Application.EnableVisualStyles()
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
            System.Windows.Forms.Application.Run(New PickerContext())
            System.GC.KeepAlive(instance)
        End Using
    End Sub
End Module
