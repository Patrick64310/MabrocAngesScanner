
Imports System
Imports System.Windows.Forms

Module Program
    <STAThread>
    Sub Main()
        ApplicationConfiguration.Initialize()
        Application.Run(New Form1())
    End Sub
End Module

