
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.WinForms

Public Class Form1
    Inherits Form

    Private web As WebView2

    Public Sub New()
        Me.Text = "Test WebView2"
        Me.Width = 800
        Me.Height = 600

        web = New WebView2()
        web.Dock = DockStyle.Fill
        Me.Controls.Add(web)

        AddHandler Me.Load, AddressOf OnLoad
    End Sub

    
Private Sub Form1_Load(sender As Object, e As EventArgs) _
    Handles MyBase.Load

        Await web.EnsureCoreWebView2Async()
        web.Source = New Uri("https://www.etsy.com")
    End Sub

End Class
