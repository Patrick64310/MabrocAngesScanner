
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    Public Sub New()
        Me.Text = "Mabroc'Anges Scanner"
        Me.Width = 500
        Me.Height = 300
    End Sub

    <STAThread>
    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.Run(New Form1())
    End Sub
End Class
