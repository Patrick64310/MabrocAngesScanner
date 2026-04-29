
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class Form1
    Inherits Form

    ' ========= HTML EN MÉMOIRE =========
    Private PagesHtml As New Dictionary(Of Integer, String)

    ' ========= LISTES =========
    Private ArticlesUrl As New List(Of String)

    ' ========= ÉTAT =========
    Private Running As Boolean = False
    Private LoopRunning As Boolean = False
    Private TotalClicks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private LoopCount As Integer = 1

    ' ========= TEMPS =========
    Private LoopStartTime As DateTime
    Private uiTimer As Timer

    ' ========= UI =========
    Private lblClicks, lblArticles, lblTime As Label
    Private lblCurrentArticle, lblArticleTitle, lblProgress As Label
    Private picThumbnail As PictureBox
    Private btnStart, btnStop As Button
    Private chkVisualiserArticles As CheckBox

    ' ========= WEBVIEW =========
    Private webPages As WebView2
    Private webArticle As WebView2

    ' ========= REGEX =========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()

        Me.Text = "Mabroc'Anges – Scanner Etsy (Excel final)"
        Me.Width = 1000
        Me.Height = 620
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI
    End Sub

    ' ========= UI =========
    Private Sub InitializeUI()

        lblClicks = New Label() With {.Left = 20, .Top = 20, .Width = 950}
        lblArticles = New Label() With {.Left = 20, .Top = 45, .Width = 950}
        lblTime = New Label() With {.Left = 20, .Top = 70, .Width = 950}

        lblCurrentArticle = New Label() With {.Left = 20, .Top = 100, .Width = 950}
        lblArticleTitle = New Label() With {.Left = 20, .Top = 130, .Width = 950}
        lblProgress = New Label() With {.Left = 20, .Top = 160, .Width = 950}

        picThumbnail = New PictureBox() With {
            .Left = 20,
            .Top = 190,
            .Width = 200,
            .Height = 200,
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle
        }

        chkVisualiserArticles = New CheckBox() With {
            .Text = "Visualiser articles",
            .Left = 250,
            .Top = 190,
            .Checked = True
        }

        btnStart = New Button() With {.Text = "START", .Left = 420, .Top = 185, .Width = 120}
        btnStop = New Button() With {.Text = "STOP", .Left = 560, .Top = 185, .Width = 120}

        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        webPages = New WebView2() With {.Visible = False}
        webArticle = New WebView2() With {
            .Left = 250,
            .Top = 230,
            .Width = 700,
            .Height = 330,
            .Visible = False
        }

        Controls.AddRange({
            lblClicks, lblArticles, lblTime,
            lblCurrentArticle, lblArticleTitle, lblProgress,
            picThumbnail,
            chkVisualiserArticles, btnStart, btnStop,
            webPages, webArticle
        })
    End Sub

    ' ========= START =========
    Private Async Sub StartAsync(sender As Object, e As EventArgs)

        Running = True
        ArticlesUrl.Clear()
        LoopCount = 1
        TotalClicks = 0

        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        ' === Capture pages boutique ===
        For page = 1 To 20
            webPages.Source = New Uri(
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items")
            Await Task.Delay(6000)

            Dim html = Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")
            html = html.Replace("\""", """")

            If html.Contains("Aucun article en vente pour le moment") Then Exit For

            For Each m As Match In ListingRegex.Matches(html)
                Dim url = m.Value
                If url.Length < 100 AndAlso Not ArticlesUrl.Contains(url) Then
                    ArticlesUrl.Add(url)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        LoopRunning = True
        LoopStartTime = DateTime.Now
        uiTimer.Start()

        Dim index As Integer = 0

        If chkVisualiserArticles.Checked Then webArticle.Visible = True

        ' === BOUCLE INFINIE ARTICLES ===
        While Running

            Dim articleUrl = ArticlesUrl(index)
            lblCurrentArticle.Text = "Article : " & articleUrl
            lblProgress.Text = $"Article {index + 1} / {ArticlesFound} (tour n°{LoopCount})"
            TotalClicks += 1

            If chkVisualiserArticles.Checked Then
                webArticle.CoreWebView2.Navigate(articleUrl)
                Await Task.Delay(1500)

                ' Titre
                lblArticleTitle.Text =
                    Await webArticle.ExecuteScriptAsync("document.title").ContinueWith(
                        Function(t) t.Result.Replace("""", ""))

                ' Miniature
                Dim imgUrl =
                    Await webArticle.ExecuteScriptAsync(
                        "document.querySelector('meta[property=""og:image""]')?.content")

                imgUrl = imgUrl.Replace("""", "")
                If imgUrl.StartsWith("http") Then
                    picThumbnail.LoadAsync(imgUrl)
                End If
            End If

            Await Task.Delay(4000)

            index += 1
            If index >= ArticlesFound Then
                index = 0
                LoopCount += 1
            End If
        End While
    End Sub

    ' ========= STOP =========
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        uiTimer.Stop()
        webArticle.Visible = False
        picThumbnail.Image = Nothing
        lblArticleTitle.Text = ""
        lblProgress.Text = ""
    End Sub

    ' ========= UI =========
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblClicks.Text = "Clics cumulés : " & TotalClicks
        lblArticles.Text = "Articles trouvés : " & ArticlesFound
        lblTime.Text = "Temps activité : " & (DateTime.Now - LoopStartTime).ToString("hh\:mm\:ss")
    End Sub

End Class
