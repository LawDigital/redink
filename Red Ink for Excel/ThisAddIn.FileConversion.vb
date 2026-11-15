' Part of "Red Ink for Excel".
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Explicit On
Option Strict On

Partial Public Class ThisAddIn
    <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
    Private Shared Function LoadConversionLibrary(path As System.String) As System.IntPtr
    End Function

    Private Shared Sub EnsureConversionPdfiumLoaded()
        Dim assemblyDirectory As System.String = System.IO.Path.GetDirectoryName(GetType(ThisAddIn).Assembly.Location)
        Dim architecture As System.String = If(System.Environment.Is64BitProcess, "x64", "x86")
        For Each directory As System.String In New System.String() {assemblyDirectory, System.AppDomain.CurrentDomain.BaseDirectory}
            For Each relativePath As System.String In New System.String() {System.IO.Path.Combine(architecture, "pdfium.dll"), "pdfium.dll"}
                Dim candidate As System.String = System.IO.Path.Combine(directory, relativePath)
                If System.IO.File.Exists(candidate) AndAlso LoadConversionLibrary(candidate) <> System.IntPtr.Zero Then Return
            Next
        Next
        Throw New System.IO.FileNotFoundException("Restore the PdfiumViewer native packages and rebuild Excel to enable PDF flattening.")
    End Sub
End Class
