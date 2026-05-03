
Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Module Program

    ' ✅ AppUserModelID : icône correcte dans la barre des tâches / tray
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Sub SetCurrentProcessExplicitAppUserModelID(appID As String)
    End Sub

    <STAThread>
    Sub Main()
        ' Identifiant Windows stable
        SetCurrentProcessExplicitAppUserModelID("MabrocAnges.ScannerEtsy")
        ' ✅ Initialisation WinForms CORRECTE pour VB.NET
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form1())

End Module
