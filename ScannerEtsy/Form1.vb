
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Public Class Form1
    Inherits Form

    Private web As WebView2
    Private PageNumber As Integer = 1
    Private StopRequested As Boolean = False

    Private Const ShopBaseUrl As String =
        "https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={0}&sort_order=date_desc#items"

    Private ReadOnly ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()
        InitializeComponent()

        web = New WebView2()
        web.Dock = DockStyle.Fill
        web.Visible = False ' caché
        Me.Controls.Add(web)

        AddHandler web.NavigationCompleted, AddressOf OnPageLoaded
        web.Source = New Uri(String.Format(ShopBaseUrl, PageNumber))
    End Sub

    Private Async Sub OnPageLoaded(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)

        If StopRequested Then Exit Sub

        ' Lire le DOM réel
        Dim html As String =
            Await web.ExecuteScriptAsync("document.documentElement.outerHTML")

        html = html.Replace("\""", """") ' nettoyage JSON

        Dim matches = ListingRegex.Matches(html)

        If matches.Count = 0 Then
            Exit Sub
        End If

        For Each m As Match In matches
            Debug.WriteLine(m.Groups(1).Value)
        Next

        PageNumber += 1
        web.Source = New Uri(String.Format(ShopBaseUrl, PageNumber))
    End Sub
End Class

