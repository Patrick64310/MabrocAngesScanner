
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
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
        ApplyWindowIcon()

        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        statusTimer = New Timer() With {.Interval = 60}
        AddHandler statusTimer.Tick, AddressOf AnimateStatus
    End Sub

    ' ================= ICÔNE DE LA FENÊTRE =================
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
            ' Sécurité : ignorer si introuvable
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
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 320))   ' Images (fixe)
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))    ' Bas

        ' ===== HEADER =====
        lblCurrentArticle = New Label With {
            .AutoSize = False,
            .Height = 28,
            .Width = 1050,
            .Font = New Font("Arial", 10, FontStyle.Regular),
            .ForeColor = Color.DarkBlue,
            .Text = "",
            .TextAlign = ContentAlignment.MiddleLeft
        }

        lblArticleTitle = New Label With {
            .AutoSize = False,
            .Height = 56,
            .Width = 1050,
            .Font = New Font("Arial", 11, FontStyle.Regular),
            .ForeColor = Color.DarkGreen,
            .Text = "Cliquer sur START pour commencer",
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

        ' ===== IMAGES (MINIATURE | LOGO) =====
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

        ' Chargement robuste du logo embarqué
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

        ' ===== BAS : GAUCHE (COMPTEURS) / DROITE (ACTIONS) =====
        Dim bottomPanel As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .Padding = New Padding(10)
        }
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))

        ' --- Compteurs (gauche)
        Dim counters As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.TopDown
        }

        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        lblProgress = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblClicks = New Label With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.Red}
        lblArticles = New Label With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.Green}
        lblDead = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblTime = New Label With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.DarkBlue}

        counters.Controls.Add(lblTime)
		lblTime.Margin     = New Padding(0, 0, 150, 5)
        counters.Controls.Add(lblClicks)
		lblClicks.Margin   = New Padding(0, 0, 150, 5)		
        counters.Controls.Add(lblArticles)
		lblArticles.Margin = New Padding(0, 0, 150, 5)
        counters.Controls.Add(lblDead)
		lblDead.Margin     = New Padding(0, 0, 150, 5)
        counters.Controls.Add(lblProgress)
		counters.Padding = New Padding(0, 10, 150, 0)		
		
        bottomPanel.Controls.Add(counters, 0, 0)

        ' --- Actions (droite, bas)
        Dim actions As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Fill,
            .WrapContents = False,
            .AutoSize = False,
            .MinimumSize = New Size(160, 0),
            .Padding = New Padding(10)
        }


        pnlStatus = New Panel With {
            .Width = 100,
            .Height = 60,
            .BackColor = Color.Red,
            .Margin = New Padding(0, 0, 0, 40)
        }
        'AddHandler pnlStatus.Paint, AddressOf DrawStatusBorder

        btnStart = New Button With {.Text = "START", .Width = 100, .Height = 60, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 100, .Height = 60, .Font = fnt, .Visible = False}

        ' START = vert
        AddHandler btnStart.MouseEnter, Sub()
            btnStart.BackColor = Color.LightGreen
        End Sub
        AddHandler btnStart.MouseLeave, Sub()
            btnStart.BackColor = SystemColors.Control
        End Sub
        
        ' STOP = rouge
        AddHandler btnStop.MouseEnter, Sub()
            btnStop.BackColor = Color.LightCoral
        End Sub
        AddHandler btnStop.MouseLeave, Sub()
            btnStop.BackColor = SystemColors.Control
        End Sub
                                                                                    
        AddHandler btnStart.Click, AddressOf StartAsync
        AddHandler btnStop.Click, AddressOf StopProcess
		
		btnStart.Margin = New Padding(0, 0, 60, 0)
		btnStop.Margin  = New Padding(0, 0, 60, 0)

        actions.Controls.Add(pnlStatus)
        actions.Controls.Add(btnStart)
        actions.Controls.Add(btnStop)
        'actions.Padding = New Padding(0, 0, 70, 0)
		
		
        bottomPanel.Controls.Add(actions, 1, 0)
        root.Controls.Add(bottomPanel, 0, 3)

        Me.Controls.Add(root)

        webPages = New WebView2 With {.Visible = False}
        webArticle = New WebView2 With {.Visible = False}
        Me.Controls.Add(webPages)
        Me.Controls.Add(webArticle)
    End Sub

        ' Ajouter le cadre noir
        Private Sub DrawStatusBorder(sender As Object, e As PaintEventArgs)
            Using pen As New Pen(Color.Black, 1)
                e.Graphics.DrawRectangle(
                    pen,
                    3,
                    3,
                    pnlStatus.Width - 1,
                    pnlStatus.Height - 1
                )
            End Using
        End Sub
                                                                                        
    ' ===== ANIMATION DU VOYANT (FADE VERT) =====
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
        lblArticleTitle.Text = "Recherche en cours . . . Merci de patienter  . . . "
        ArticlesUrl.Clear()
        TotalClicks = 0
        DeadLinks = 0
        LoopCount = 1

        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        LoopStartTime = DateTime.Now
        uiTimer.Start()

        For page = 1 To 10
            webPages.Source = New Uri($"https://www.etsy.com/fr/shop/mabrocanges?page={page}")
            Await Task.Delay(2600)

            Dim html = Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")
            html = html.Replace("""", "")

			If html.Contains("Aucun article en vente pour le moment") Then Exit For

            For Each m As Match In ListingRegex.Matches(html)
                If m.Value.Length < 100 AndAlso Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
                End If
				lblArticleTitle.Text = $"Recherche en cours . . . Merci de patienter  . . .   {ArticlesUrl.Count} articles "
				if ArticlesUrl.Count > 0 then lblArticles.Text = $"Articles trouvés :    {ArticlesUrl.Count}"
            Next
        Next

        ArticlesFound = ArticlesUrl.Count


        Dim rnd As New Random()
        Dim i As Integer = 0

        While Running
            Dim url = ArticlesUrl(i)
            ' lblCurrentArticle.Text = "Lien de l'article : " & url

            If LoopCount = 1 Then
                lblProgress.Text = $"Article {i + 1} / {ArticlesFound}     (1er tour)"
            Else
                lblProgress.Text = $"Article {i + 1} / {ArticlesFound}     ({LoopCount}ème tour)"
            End If

            TotalClicks += 1

            Try
                webArticle.CoreWebView2.Navigate(url)
                Await Task.Delay(1500)

                lblArticleTitle.Text =
                    (Await webArticle.ExecuteScriptAsync("document.title")).Replace("""", "")
				lblCurrentArticle.Text = "Lien de l'article : " & url
                Dim img = Await webArticle.ExecuteScriptAsync(
                    "document.querySelector('meta[property=""og:image""]')?.content")
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

    ' ================= STOP =================
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

    ' ================= UI TIMER =================
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblClicks.Text = $"Clics cumulés :    {TotalClicks}"
        lblArticles.Text = $"Articles trouvés :    {ArticlesFound}"
        lblDead.Text = $"Liens morts :    {DeadLinks}"
        lblTime.Text = $"Temps activité :    {(DateTime.Now - LoopStartTime):hh\:mm\:ss}"
    End Sub

End Class
