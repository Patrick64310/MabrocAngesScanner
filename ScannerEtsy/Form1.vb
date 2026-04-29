
Imports System.Windows.Forms
Imports System.Text.RegularExpressions Inherits FormImports System.Text.RegularExpressions

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
        Me.Text = "Mabroc'Anges – Scanner Etsy"
        Me.Width = 1150
        Me.Height = 720
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.AliceBlue

        InitializeUI()

        ' ✅ APPLIQUER EXPLICITEMENT L’ICÔNE À LA FENÊTRE
        ApplyWindowIcon()

        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        statusTimer = New Timer() With {.Interval = 60}
        AddHandler statusTimer.Tick, AddressOf AnimateStatus
    End Sub

    ' ================= ICÔNE FENÊTRE =================
    Private Sub ApplyWindowIcon()
        Try
            Dim asm = GetType(Form1).Assembly
            For Each resName In asm.GetManifestResourceNames()
                If resName.EndsWith(".etsy.ico", StringComparison.OrdinalIgnoreCase) _
                   OrElse resName.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) Then

                    Using s = asm.GetManifestResourceStream(resName)
                        If s IsNot Nothing Then
                            Me.Icon = New Icon(s)
                            Exit For
                        End If
                    End Using

                End If
            Next
        Catch
            ' Sécurité : ne rien faire si l’icône n’est pas trouvée
        End Try
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

        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))        ' Header
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 6))     ' Separator
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 320))   ' Images
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))    ' Bottom

        ' ===== HEADER =====
        lblCurrentArticle = New Label With {
            .AutoSize = False,
            .Height = 28,
            .Width = 1050,
            .Font = New Font("Arial", 10, FontStyle.Regular),
            .Text = "Article :",
            .TextAlign = ContentAlignment.MiddleLeft
        }

        lblArticleTitle = New Label With {
            .AutoSize = False,
            .Height = 56,
            .Width = 1050,
            .Font = New Font("Arial", 12, FontStyle.Regular),
            .Text = "Description :",
            .TextAlign = ContentAlignment.MiddleLeft
        }

        Dim header As New FlowLayoutPanel With {
            .AutoSize = True,
            .FlowDirection = FlowDirection.TopDown,
            .WrapContents = False
        }

        header.Controls.Add(lblCurrentArticle)
        header.Controls.Add(lblArticleTitle)
        root.Controls.Add(header, 0, 0)

        root.Controls.Add(New Panel With {
            .Dock = DockStyle.Fill,
            .Height = 2,
            .BackColor = Color.DarkGray
        }, 0, 1)

        ' ===== IMAGES =====
        Dim imagesRow As New TableLayoutPanel With {
            .ColumnCount = 2,
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

        ' Logo embarqué (chargement robuste)
        For Each r In GetType(Form1).Assembly.GetManifestResourceNames()
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

        ' ===== BAS : GAUCHE / DROITE =====
        Dim bottomPanel As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .Padding = New Padding(10)
        }

        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))

        ' --- Compteurs (gauche)
        Dim counters As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown
        }

        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        lblProgress = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblClicks = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblArticles = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblDead = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblTime = New Label With {.Width = 520, .Height = 30, .Font = fnt}

        counters.Controls.Add(lblProgress)
        counters.Controls.Add(lblClicks)
        counters.Controls.Add(lblArticles)
        counters.Controls.Add(lblDead)
        counters.Controls.Add(lblTime)

        bottomPanel.Controls.Add(counters, 0, 0)

        ' --- Actions (droite, bas)
        Dim actions As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Bottom,
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }

        pnlStatus = New Panel With {
            .Width = 18,
            .Height = 18,
            .BackColor = Color.Red,
            .Margin = New Padding(0, 0, 0, 10)
        }

        btnStart = New Button With {.Text = "START", .Width = 160, .Height = 40, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 160, .Height = 40, .Font = fnt, .Visible = False}

        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        actions.Controls.Add(pnlStatus)
        actions.Controls.Add(btnStart)
        actions.Controls.Add(btnStop)

        bottomPanel.Controls.Add(actions, 1, 0)
        root.Controls.Add(bottomPanel, 0, 3)

        Me.Controls.Add(root)

        webPages = New WebView2 With {.Visible = False}
        webArticle = New WebView2 With {.Visible = False}
        Me.Controls.Add(webPages)
        Me.Controls.Add(webArticle)
    End Sub

    ' ===== ANIMATION VOYANT (FADE VERT) =====
    Private Sub AnimateStatus(sender As Object, e As EventArgs)
        If Not Running Then Exit Sub

        fadeValue += fadeDir * 8
        If fadeValue >= 255 Then fadeValue = 255 : fadeDir = -1
        If fadeValue <= 80 Then fadeValue = 80 : fadeDir = 1

        pnlStatus.BackColor = Color.FromArgb(0, fadeValue, 0)
    End Sub

    ' ================= START =================
    Private Async Sub StartAsync(sender As Object, e As EventArgs)

        Running = True
        btnStart.Visible = False
        btnStop.Visible = True
        statusTimer.Start()

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
            Dim url = ArticlesUrl(i)
            lblCurrentArticle.Text = "Article : " & url
            lblProgress.Text = $"Article {i + 1} / {ArticlesFound} (tour n°{LoopCount})"
            TotalClicks += 1

            Try
                webArticle.CoreWebView2.Navigate(url)
                Await Task.Delay(1500)

                lblArticleTitle.Text =
                    (Await webArticle.ExecuteScriptAsync("document.title")).Replace("""", "")

                Dim img = Await webArticle.ExecuteScriptAsync(
                    "document.querySelector('meta[property=""og:image""]')?.content")
                img = img.Replace("""", "")
                If img.StartsWith("http") Then picThumbnail.LoadAsync(img)
            Catch
                DeadLinks += 1
            End Try

            Await Task.Delay(rnd.Next(3000, 9000))
            i = (i + 1) Mod ArticlesFound
            If i = 0 Then LoopCount += 1
        End While
    End Sub

    ' ================= STOP =================
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        btnStart.Visible = True
        btnStop.Visible = False
        statusTimer.Stop()
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
``
Imports System.IO
Imports System.Drawing
Imports Microsoft.Web.WebView2.WinForms

Public Class Form1
