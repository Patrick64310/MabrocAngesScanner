
Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Module Program

    ' ✅ AppUserModelID : indispensable pour l’icône barre des tâches
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Private Sub SetCurrentProcessExplicitAppUserModelID(appID As String)
    End Sub
 
    <STAThread>
    Sub Main()

        ' ✅ ID unique pour Windows (choisis une valeur stable)
        SetCurrentProcessExplicitAppUserModelID("MabrocAnges.ScannerEtsy")
        ApplicationConfiguration.Initialize()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Application.Run(New Form1())
    End Sub

End Module
