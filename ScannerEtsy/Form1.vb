
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports Microsoft.Web.WebView2.WinForms

Public Class Form1
    Inherits Form

    ' ========= DONNEES =========
    Private ArticlesUrl As New List(Of String)

    ' ========= ETAT =========
    Private Running As Boolean
    Private TotalClicks As Integer
    Private ArticlesFound As Integer
    Private DeadLinks As Integer
    Private LoopCount As Integer = 1

    ' ========= TEMPS =========
    Private LoopStartTime As DateTime
    Private uiTimer As Timer
    Private statusTimer As Timer
    Private fadeValue As Integer = 80
    Private fadeDir As Integer = 1

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
    Private pnlStatus As Panel

    Private picThumbnail As PictureBox
    Private picLogo As PictureBox

    ' ========= WEBVIEW =========
    Private webPages As WebView2
    Private webArticle As WebView2

    ' ========= REGEX =========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()
        Me.Text = "Mabroc'Anges - Scanner Etsy"
        Me.Width = 1150
        Me.Height = 720
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.AliceBlue

        InitializeUI()
        ApplyWindowIcon()

        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        statusTimer = New Timer()
        statusTimer.Interval = 60
        AddHandler statusTimer.Tick, AddressOf AnimateStatus
    End Sub

    ' ========= ICONE =========
    Private Sub ApplyWindowIcon()
        Try
            For Each r As String In GetType(Form1).Assembly.GetManifestResourceNames()
                If r.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) Then
                    Using s = GetType(Form1).Assembly.GetManifestResourceStream(r)
                        If s IsNot Nothing Then
                            Me.Icon = New Icon(s)
                            Exit For
                        End If
                    End Using
                End If
            Next
        Catch
        End Try
    End Sub

    ' ========= UI =========
    Private Sub InitializeUI()

        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 4
        root.Padding = New Padding(10)

        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 6))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 320))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        ' ----- HEADER -----
        lblCurrentArticle = New Label()
        lblCurrentArticle.Width = 1050
        lblCurrentArticle.Height = 28
        lblCurrentArticle.Font = New Font("Arial", 10)
        lblCurrentArticle.ForeColor = Color.DarkBlue

        lblArticleTitle = New Label()
        lblArticleTitle.Width = 1050
        lblArticleTitle.Height = 56
        lblArticleTitle.Font = New Font("Arial", 11)
        lblArticleTitle.ForeColor = Color.DarkGreen
        lblArticleTitle.Text = "Cliquer sur START pour commencer"

        Dim header As New FlowLayoutPanel()
        header.AutoSize = True
        header.FlowDirection = FlowDirection.TopDown
        header.Controls.Add(lblCurrentArticle)
        header.Controls.Add(lblArticleTitle)

        root.Controls.Add(header, 0, 0)
        root.Controls.Add(New Panel() With {.BackColor = Color.DarkGray, .Height = 2, .Dock = DockStyle.Fill}, 0, 1)

        ' ----- IMAGES -----
        Dim imagesRow As New TableLayoutPanel()
        imagesRow.ColumnCount = 2
        imagesRow.Dock = DockStyle.Fill
        imagesRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        imagesRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

        picThumbnail = New PictureBox()
        picThumbnail.Width = 300
        picThumbnail.Height = 300
        picThumbnail.SizeMode = PictureBoxSizeMode.Zoom
        picThumbnail.BorderStyle = BorderStyle.FixedSingle
        picThumbnail.Anchor = AnchorStyles.None

        picLogo = New PictureBox()
        picLogo.Width = 300
        picLogo.Height = 300
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        picLogo.BorderStyle = BorderStyle.FixedSingle
        picLogo.Anchor = AnchorStyles.None

        For Each r As String In GetType(Form1).Assembly.GetManifestResourceNames()
            If r.EndsWith(".logo.png", StringComparison.OrdinalIgnoreCase) Then
                Using s = GetType(Form1).Assembly.GetManifestResourceStream(r)
                    picLogo.Image = Image.FromStream(s)
                End Using
                Exit For
            End If
        Next

        imagesRow.Controls.Add(picThumbnail, 0, 0)
        imagesRow.Controls.Add(picLogo, 1, 0)
        root.Controls.Add(imagesRow, 0, 2)

        ' ----- BAS -----
        Dim bottomPanel As New TableLayoutPanel()
        bottomPanel.Dock = DockStyle.Fill
        bottomPanel.ColumnCount = 2
        bottomPanel.Padding = New Padding(10)
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))

        Dim counters As New FlowLayoutPanel()
        counters.FlowDirection = FlowDirection.TopDown

        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        lblTime = New Label() With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.DarkGreen, .Margin = New Padding(0, 0, 0, 10)}
        lblClicks = New Label() With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.Red, .Margin = New Padding(0, 0, 0, 10)}
        lblArticles = New Label() With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.DarkBlue, .Margin = New Padding(0, 0, 0, 10)}
        lblDead = New Label() With {.Width = 520, .Height = 30, .Font = fnt, .Margin = New Padding(0, 0, 0, 10)}
        lblProgress = New Label() With {.Width = 520, .Height = 30, .Font = fnt}

        counters.Controls.Add(lblTime)
        counters.Controls.Add(lblClicks)
        counters.Controls.Add(lblArticles)
        counters.Controls.Add(lblDead)
        counters.Controls.Add(lblProgress)

        bottomPanel.Controls.Add(counters, 0, 0)

        Dim actions As New FlowLayoutPanel()
        actions.FlowDirection = FlowDirection.TopDown
        actions.Dock = DockStyle.Bottom
        actions.Padding = New Padding(0, 0, 60, 0)

        pnlStatus = New Panel()
        pnlStatus.Width = 60
        pnlStatus.Height = 60
        pnlStatus.BackColor = Color.Red
        pnlStatus.Margin = New Padding(0, 0, 0, 60)

        btnStart = New Button() With {.Text = "START", .Width = 160, .Height = 60, .Font = fnt}
        btnStop = New Button() With {.Text = "STOP", .Width = 160, .Height = 60, .Font = fnt, .Visible = False}

        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        actions.Controls.Add(pnlStatus)
        actions.Controls.Add(btnStart)
        actions.Controls.Add(btnStop)

        bottomPanel.Controls.Add(actions, 1, 0)
        root.Controls.Add(bottomPanel, 0, 3)

        Me.Controls.Add(root)

        webPages = New WebView2() With {.Visible = False}
        webArticle = New WebView2() With {.Visible = False}
        Me.Controls.Add(webPages)
        Me.Controls.Add(webArticle)
    End Sub

    ' ========= VOYANT =========
    Private Sub AnimateStatus(sender As Object, e As EventArgs)
        If Not Running Then Exit Sub

        fadeValue += fadeDir * 8
        If fadeValue >= 255 Then fadeDir = -1
        If fadeValue <= 80 Then fadeDir = 1

        pnlStatus.BackColor = Color.FromArgb(0, fadeValue, 0)
    End Sub

    ' ========= START =========
    Private Async Sub StartAsync(sender As Object, e As EventArgs)

        Running = True
        btnStart.Visible = False
        btnStop.Visible = True
        statusTimer.Start()

        lblArticleTitle.Text = "Recherche en cours..."
        ArticlesUrl.Clear()
        TotalClicks = 0
        DeadLinks = 0
        LoopCount = 1

        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        For page As Integer = 1 To 20
            webPages.Source = New Uri("https://www.etsy.com/fr/shop/mabrocanges?page=" & page)
            Await Task.Delay(900)

            Dim html As String = Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")
            html = html.Replace("""", "")

            For Each m As Match In ListingRegex.Matches(html)
                If m.Value.Length < 100 AndAlso Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        LoopStartTime = DateTime.Now
        uiTimer.Start()

        Dim rnd As New Random()
        Dim i As Integer = 0

        While Running
            Dim url As String = ArticlesUrl(i)
            lblCurrentArticle.Text = "Lien de l'article : " & url

            If LoopCount = 1 Then
                lblProgress.Text = "Article " & (i + 1) & "  /  " & ArticlesFound & " (1er tour)"
            Else
                lblProgress.Text = "Article " & (i + 1) & "  /  " & ArticlesFound & " (" & LoopCount & "eme tour)"
            End If

            TotalClicks += 1

            Try
                webArticle.CoreWebView2.Navigate(url)
                Await Task.Delay(1200)

                lblArticleTitle.Text = (Await webArticle.ExecuteScriptAsync("document.title")).Replace("""", "")

                Dim img As String = Await webArticle.ExecuteScriptAsync("document.querySelector('meta[property=""og:image""]')?.content")
                img = img.Replace("""", "")
                If img.StartsWith("http") Then picThumbnail.LoadAsync(img)
            Catch
                DeadLinks += 1
            End Try

            Await Task.Delay(rnd.Next(2000, 9000))
            i = (i + 1) Mod ArticlesFound
            If i = 0 Then LoopCount += 1
        End While
    End Sub

    ' ========= STOP =========
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        btnStart.Visible = True
        btnStop.Visible = False
        statusTimer.Stop()
        pnlStatus.BackColor = Color.Red
        uiTimer.Stop()
        picThumbnail.Image = Nothing
        lblArticleTitle.Text = ""
        lblCurrentArticle.Text = ""
    End Sub

    ' ========= UI TIMER =========
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblClicks.Text = "Clics cumules :    " & TotalClicks
        lblArticles.Text = "Articles trouves :    " & ArticlesFound
        lblDead.Text = "Liens morts :    " & DeadLinks
        lblTime.Text = "Temps activite :    " & (DateTime.Now - LoopStartTime).ToString("hh\:mm\:ss")
    End Sub

End Class
