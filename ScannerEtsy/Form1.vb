
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
    Private picLogo As PictureBox
    Private pnlStatus As Panel   ' voyant RUN / STOP

    ' ========= WEBVIEW =========
    Private webPages As WebView2
    Private webArticle As WebView2

    ' ========= REGEX =========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner Etsy"
        Me.Width = 1050
        Me.Height = 700
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.AliceBlue

        InitializeUI()

        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI
    End Sub

    ' ================= UI =================
    Private Sub InitializeUI()

        ' ===== ROOT =====
        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(10)
        }

        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))       ' Header
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 6))    ' Separator
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 320))  ' Images FIXE
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))   ' Controls

        ' ===== HEADER =====
        lblCurrentArticle = New Label With {
            .Height = 32,
            .Dock = DockStyle.Top,
            .Font = New Font("Arial", 10, FontStyle.Regular),
            .Text = "Article :"
        }

        lblArticleTitle = New Label With {
            .Height = 60,
            .Dock = DockStyle.Top,
            .Font = New Font("Arial", 12, FontStyle.Regular),
            .Text = "Description :"
        }

        Dim header As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown,
            .AutoSize = True,
            .WrapContents = False
        }

        header.Controls.Add(lblCurrentArticle)
        header.Controls.Add(lblArticleTitle)
        root.Controls.Add(header, 0, 0)

        ' ===== SEPARATOR =====
        root.Controls.Add(New Panel With {
            .Dock = DockStyle.Fill,
            .Height = 2,
            .BackColor = Color.DarkGray
        }, 0, 1)

        ' ===== IMAGES ROW (MINIATURE | LOGO) =====
        Dim imagesRow As New TableLayoutPanel With {
            .ColumnCount = 2,
            .RowCount = 1,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10)
        }

        imagesRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        imagesRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

        picThumbnail = New PictureBox With {
            .Width = 300,
            .Height = 300,
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .Anchor = AnchorStyles.None
        }

        picLogo = New PictureBox With {
            .Width = 300,
            .Height = 300,
            .SizeMode = PictureBoxSizeMode.Zoom,
            .BorderStyle = BorderStyle.FixedSingle,
            .Anchor = AnchorStyles.None
        }

        ' === Chargement ROBUSTE du logo embarqué ===
        Dim asm = GetType(Form1).Assembly
        For Each resName In asm.GetManifestResourceNames()
            If resName.EndsWith(".logo.png", StringComparison.OrdinalIgnoreCase) Then
                Using s = asm.GetManifestResourceStream(resName)
                    If s IsNot Nothing Then
                        picLogo.Image = Image.FromStream(s)
                        Exit For
                    End If
                End Using
            End If
        Next

        imagesRow.Controls.Add(picThumbnail, 0, 0)
        imagesRow.Controls.Add(picLogo, 1, 0)
        root.Controls.Add(imagesRow, 0, 2)

        ' ===== CONTROLS =====
        Dim controlsPanel As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown,
            .AutoScroll = True
        }

        pnlStatus = New Panel With {
            .Width = 18,
            .Height = 18,
            .BackColor = Color.Red,
            .Margin = New Padding(0, 0, 0, 10)
        }

        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        btnStart = New Button With {.Text = "START", .Width = 160, .Height = 40, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 160, .Height = 40, .Font = fnt, .Visible = False}

        AddHandler btnStart.MouseEnter, Sub() btnStart.BackColor = Color.LightGreen
        AddHandler btnStart.MouseLeave, Sub() btnStart.BackColor = SystemColors.Control
        AddHandler btnStop.MouseEnter, Sub() btnStop.BackColor = Color.LightCoral
        AddHandler btnStop.MouseLeave, Sub() btnStop.BackColor = SystemColors.Control

        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        lblProgress = New Label With {.Height = 30, .Width = 520, .Font = fnt}
        lblClicks = New Label With {.Height = 30, .Width = 520, .Font = fnt}
        lblArticles = New Label With {.Height = 30, .Width = 520, .Font = fnt}
        lblDead = New Label With {.Height = 30, .Width = 520, .Font = fnt}
        lblTime = New Label With {.Height = 30, .Width = 520, .Font = fnt}

        controlsPanel.Controls.Add(pnlStatus)
        controlsPanel.Controls.Add(btnStart)
        controlsPanel.Controls.Add(btnStop)
        controlsPanel.Controls.Add(lblProgress)
        controlsPanel.Controls.Add(lblClicks)
        controlsPanel.Controls.Add(lblArticles)
        controlsPanel.Controls.Add(lblDead)
        controlsPanel.Controls.Add(lblTime)

        root.Controls.Add(controlsPanel, 0, 3)
        Me.Controls.Add(root)

        webPages = New WebView2 With {.Visible = False}
        webArticle = New WebView2 With {.Visible = False}
        Me.Controls.Add(webPages)
        Me.Controls.Add(webArticle)
    End Sub

    ' ================= START =================
    Private Async Sub StartAsync(sender As Object, e As EventArgs)

        Running = True
        btnStart.Visible = False
        btnStop.Visible = True
        pnlStatus.BackColor = Color.LimeGreen

        ArticlesUrl.Clear()
        TotalClicks = 0
        DeadLinks = 0
        LoopCount = 1

        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        For page = 1 To 20
            webPages.Source = New Uri($"https://www.etsy.com/fr/shop/mabrocanges?page={page}")
            Await Task.Delay(3000)

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

        While Running

            Dim url = ArticlesUrl(index)
            lblCurrentArticle.Text = "Article : " & url
            lblProgress.Text = $"Article {index + 1} / {ArticlesFound} (tour n°{LoopCount})"
            TotalClicks += 1

            Try
                webArticle.CoreWebView2.Navigate(url)
                Await Task.Delay(1500)

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
            If index >= ArticlesFound Then
                index = 0
                LoopCount += 1
            End If
        End While
    End Sub

    ' ================= STOP =================
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        btnStart.Visible = True
        btnStop.Visible = False
        pnlStatus.BackColor = Color.Red
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
