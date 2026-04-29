
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Drawing
Imports Microsoft.Web.WebView2.WinForms

Public Class Form1
    Inherits Form

    ' ========= DONNÉES =========
    Private ArticlesUrl As New List(Of String)

    ' ========= ÉTAT =========
    Private Running As Boolean
    Private TotalClicks As Integer
    Private ArticlesFound As Integer
    Private DeadLinks As Integer
    Private LoopCount As Integer = 1

    ' ========= TEMPS =========
    Private LoopStartTime As DateTime
    Private uiTimer As Timer

    ' ========= UI =========
    Private lblCurrentArticle As Label
    Private lblArticleTitle As Label
    Private lblProgress As Label
    Private lblClicks As Label
    Private lblArticles As Label
    Private lblDead As Label
    Private lblTime As Label
    Private btnStart As Button
    Private btnStop As Button
    Private picThumbnail As PictureBox

    ' ========= WEBVIEW =========
    Private webPages As WebView2
    Private webArticle As WebView2

    ' ========= REGEX =========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner Etsy (Final)"
        Me.Width = 1150
        Me.Height = 650
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI
    End Sub

    ' ================= UI =================
    Private Sub InitializeUI()

        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .RowCount = 3,
            .Padding = New Padding(10)
        }

        root.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 320))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 130))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 6))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        ' === URL + Description ===
        lblCurrentArticle = New Label With {
            .AutoSize = False,
            .Height = 30,
            .Dock = DockStyle.Top,
            .Font = New Font("Arial", 14, FontStyle.Regular),
            .Text = "Article :"
        }

        lblArticleTitle = New Label With {
            .AutoSize = False,
            .Height = 60,
            .Dock = DockStyle.Top,
            .Font = New Font("Arial", 14, FontStyle.Regular),
            .Text = "Description :"
        }

        Dim header As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown
        }
        header.Controls.Add(lblCurrentArticle)
        header.Controls.Add(lblArticleTitle)

        root.SetColumnSpan(header, 2)
        root.Controls.Add(header, 0, 0)

        ' === Séparateur ===
        Dim separator As New Panel With {
            .Dock = DockStyle.Fill,
            .Height = 2,
            .BackColor = Color.DarkGray
        }
        root.SetColumnSpan(separator, 2)
        root.Controls.Add(separator, 0, 1)

        ' === Image ===
        Dim imagePanel As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.FromArgb(245, 245, 245),
            .Padding = New Padding(10)
        }

        picThumbnail = New PictureBox With {
            .Width = 300,
            .Height = 300,
            .Dock = DockStyle.None,
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle
        }
        imagePanel.Controls.Add(picThumbnail)
        root.Controls.Add(imagePanel, 0, 2)

        ' === Zone droite ===
        Dim rightPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown
        }

        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        btnStart = New Button With {.Text = "START", .Width = 160, .Height = 40, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 160, .Height = 40, .Font = fnt}

        AddHandler btnStart.MouseEnter, Sub() btnStart.BackColor = Color.LightGreen
        AddHandler btnStart.MouseLeave, Sub() btnStart.BackColor = SystemColors.Control
        AddHandler btnStop.MouseEnter, Sub() btnStop.BackColor = Color.LightCoral
        AddHandler btnStop.MouseLeave, Sub() btnStop.BackColor = SystemColors.Control

        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        lblProgress = New Label With {.AutoSize = False, .Height = 30, .Width = 450, .Font = fnt}
        lblClicks = New Label With {.AutoSize = False, .Height = 30, .Width = 450, .Font = fnt}
        lblArticles = New Label With {.AutoSize = False, .Height = 30, .Width = 450, .Font = fnt}
        lblDead = New Label With {.AutoSize = False, .Height = 30, .Width = 450, .Font = fnt}
        lblTime = New Label With {.AutoSize = False, .Height = 30, .Width = 450, .Font = fnt}

        rightPanel.Controls.Add(btnStart)
        rightPanel.Controls.Add(btnStop)
        rightPanel.Controls.Add(lblProgress)
        rightPanel.Controls.Add(lblClicks)
        rightPanel.Controls.Add(lblArticles)
        rightPanel.Controls.Add(lblDead)
        rightPanel.Controls.Add(lblTime)

        root.Controls.Add(rightPanel, 1, 2)
        Me.Controls.Add(root)

        webPages = New WebView2 With {.Visible = False}
        webArticle = New WebView2 With {.Visible = False}
        Me.Controls.Add(webPages)
        Me.Controls.Add(webArticle)
    End Sub

    ' ================= START =================
    Private Async Sub StartAsync(sender As Object, e As EventArgs)

        Running = True
        ArticlesUrl.Clear()
        TotalClicks = 0
        DeadLinks = 0
        LoopCount = 1

        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        ' === Pages boutique ===
        For page = 1 To 20
            webPages.Source = New Uri(
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}")
            Await Task.Delay(6000)

            Dim html = Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")
            html = html.Replace("""", "")

            If html.Contains("Aucun article en vente pour le moment") Then Exit For

            For Each m As Match In ListingRegex.Matches(html)
                Dim url = m.Value
                If url.Length < 100 AndAlso Not ArticlesUrl.Contains(url) Then
                    ArticlesUrl.Add(url)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        LoopStartTime = DateTime.Now
        uiTimer.Start()

        Dim rnd As New Random()
        Dim index As Integer = 0

        ' === Boucle infinie ===
        While Running
            Dim url = ArticlesUrl(index)
            lblCurrentArticle.Text = "Article : " & url
            lblProgress.Text = $"Article {index + 1} / {ArticlesFound} (tour n°{LoopCount})"
            TotalClicks += 1

            Try
                webArticle.CoreWebView2.Navigate(url)
                Await Task.Delay(1200)

                lblArticleTitle.Text =
                    (Await webArticle.ExecuteScriptAsync("document.title")).Replace("""", "")

                Dim img =
                    Await webArticle.ExecuteScriptAsync(
                        "document.querySelector('meta[property=""og:image""]')?.content")
                img = img.Replace("""", "")
                If img.StartsWith("http") Then picThumbnail.LoadAsync(img)
            Catch
                DeadLinks += 1
            End Try

            Await Task.Delay(rnd.Next(3000, 9000))

            index += 1
            If index >= ArticlesFound Then index = 0 : LoopCount += 1
        End While
    End Sub

    ' ================= STOP =================
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        uiTimer.Stop()
        picThumbnail.Image = Nothing
    End Sub

    ' ================= UI TIMER =================
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblClicks.Text = $"Clics cumulés : {TotalClicks}"
        lblArticles.Text = $"Articles trouvés : {ArticlesFound}"
        lblDead.Text = $"Liens morts : {DeadLinks}"
        lblTime.Text = $"Temps activité : {(DateTime.Now - LoopStartTime):hh\:mm\:ss}"
    End Sub

End Class

