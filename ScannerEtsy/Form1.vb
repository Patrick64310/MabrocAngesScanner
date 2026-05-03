
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports Microsoft.Web.WebView2.WinForms
Imports System.Runtime.InteropServices
Public Class Form1
    Inherits Form
    ' ========= DONNÉES =========
    Private ArticlesUrl As New List(Of String)
	Private trayIcon As NotifyIcon
    Private trayMenu As ContextMenuStrip
    Private Running As Boolean
    Private TotalClicks As Integer
    Private ArticlesFound As Integer
    Private DeadLinks As Integer
    Private LoopCount As Integer = 1
    Private LoopStartTime As DateTime
    Private uiTimer As Timer
    Private statusTimer As Timer
    Private fadeValue As Integer = 80
    Private fadeDir As Integer = 1
    Private lblCurrentArticle As Label
    Private lblArticleTitle As Label
    Private lblProgress As Label
    Private lblClicks As Label
    Private lblArticles As Label
    Private lblDead As Label
    Private lblTime As Label
    Private btnStart As Button
    Private btnStop As Button
    Private pnlStatus As LedPanel
    Private picThumbnail As PictureBox
    Private picLogo As PictureBox
    Private webPages As WebView2
    Private webArticle As WebView2
    Private trayIconRun As Icon
    Private trayIconStop As Icon

    ' ========= REGEX =========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

	Protected Overrides Sub OnShown(e As EventArgs)
    MyBase.OnShown(e)
    ' Démarrage automatique une fois le handle créé
    StartAsync(Me, EventArgs.Empty)

    ' Headless visuel après coup
    Me.WindowState = FormWindowState.Minimized
    Me.ShowInTaskbar = False
    Me.Hide()
	End Sub

	Private Function LoadEmbeddedIcon(endsWithName As String) As Icon
	    Dim asm = GetType(Form1).Assembly
	    For Each res In asm.GetManifestResourceNames()
	        If res.EndsWith(endsWithName, StringComparison.OrdinalIgnoreCase) Then
	            Using s = asm.GetManifestResourceStream(res)
	                Return New Icon(s)
	            End Using
	        End If
	    Next
	    Return Nothing
	End Function

	Private Sub TrayItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)
	    Select Case e.ClickedItem.Name
	        Case "Show"
	            Me.Show()
	            Me.WindowState = FormWindowState.Normal
	            Me.ShowInTaskbar = True
	        Case "Stop"
	            StopProcess(Me, EventArgs.Empty)
	        Case "Exit"
	            trayIcon.Visible = False
	            Application.Exit()
	    End Select
	End Sub

	Private Sub InitializeTrayIcon()
    trayIconRun = LoadEmbeddedIcon("TrayIconRUN.ico")
    trayIconStop = LoadEmbeddedIcon("TrayIconSTOP.ico")
    trayMenu = New ContextMenuStrip()
    trayMenu.Items.Add("Afficher").Name = "Show"
    trayMenu.Items.Add("Stop").Name = "Stop"
    trayMenu.Items.Add("Quitter").Name = "Exit"
    AddHandler trayMenu.ItemClicked, AddressOf TrayItemClicked

    trayIcon = New NotifyIcon With {
        .Icon = TrayIconSTOP, ' 🔴 arrêté par défaut
        .Text = "Scanner Etsy – Arrêté",
        .Visible = True,
        .ContextMenuStrip = trayMenu
    }
	End Sub

    Public Sub New()
		Me.ShowInTaskbar = False
		Me.WindowState = FormWindowState.Minimized
		Me.Visible = False		
        Me.Text = "Mabroc''Anges – Scanner Etsy"
        Me.Width = 1150
        Me.Height = 700
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.AliceBlue
		Me.FormBorderStyle = FormBorderStyle.None
        InitializeUI()
        ApplyWindowIcon()
        uiTimer = New Timer() With {.Interval = 1000}
        AddHandler uiTimer.Tick, AddressOf UpdateUI
        statusTimer = New Timer() With {.Interval = 60}
        'AddHandler statusTimer.Tick, AddressOf AnimateStatus
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

	Private Sub UpdateTrayTooltip()
	    If trayIcon Is Nothing Then Exit Sub
	    Dim stateText As String = If(Running, "EN COURS", "ARRETE")
	    Dim tooltip As String =
	        $"Etat : {stateText}" & vbCrLf &
	        $"Articles : {ArticlesFound}" & vbCrLf &
	        $"Clics : {TotalClicks}" & vbCrLf &
	        '$"Morts : {DeadLinks}" & vbCrLf &
	        $"Durée : {(DateTime.Now - LoopStartTime):hh\:mm\:ss}"
	    ' Windows limite la longueur, on sécurise
	    If tooltip.Length > 120 Then
	        tooltip = tooltip.Substring(0, 120)
	    End If
	    trayIcon.Text = tooltip
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
            .Height = 20,
            .Width = 1050,
            .Font = New Font("Arial", 4, FontStyle.Regular),
            .ForeColor = Color.DarkBlue,
            .Text = "",
            .TextAlign = ContentAlignment.MiddleLeft,
			.visible = false						
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
            .Height = 4,
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
        ' Chargement du logo embarqué
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
        'lblDead = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblTime = New Label With {.Width = 520, .Height = 30, .Font = fnt, .ForeColor = Color.DarkBlue}
        counters.Controls.Add(lblTime)
        counters.Controls.Add(lblClicks)	
        counters.Controls.Add(lblArticles)
        'counters.Controls.Add(lblDead)
        counters.Controls.Add(lblProgress)
		counters.Padding = New Padding(140, 10, 0, 0)		
        bottomPanel.Controls.Add(counters, 0, 0)
        ' --- Actions (droite, bas)
        Dim actions As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Fill,
            .WrapContents = False,
            .AutoSize = False,
            .MinimumSize = New Size(160, 0),
            .Padding = New Padding(230, 10, 10, 20)
        }
        ' Voyant
		pnlStatus = New LedPanel With {
		    .Margin = New Padding(20, 0, 0, 40)
		}
        btnStart = New Button With {.Text = "START", .Width = 120, .Height = 60, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 120, .Height = 60, .Font = fnt, .Visible = False}
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
		btnStart.Margin = New Padding(0, 0, 0, 0)
		btnStop.Margin  = New Padding(0, 0, 0, 0)
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
	    InitializeTrayIcon()	
		AddCustomTitleBar(Me)																								
    End Sub

    ' ================= START =================
    Private Async Sub StartAsync(sender As Object, e As EventArgs)
		'trayIcon.Icon = trayIconRun
		If trayIconRun IsNot Nothing Then
		    trayIcon.Icon = trayIconRun
		Else
		    trayIcon.Icon = Me.Icon   ' fallback sécurisé
		End If																		
		trayIcon.Text = "Scanner Etsy – En cours"																						
        Running = True
		UpdateTrayTooltip()																								
        btnStart.Visible = False
        btnStop.Visible = True
		pnlStatus.StartLed(Color.Green)																									
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
            Await Task.Delay(4500)
            Dim html = Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")
            html = html.Replace("""", "")
			If html.Contains("Aucun article en vente pour le moment") Then Exit For
            For Each m As Match In ListingRegex.Matches(html)
                If m.Value.Length < 100 AndAlso Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
					ArticlesFound = ArticlesUrl.Count	
						If ArticlesFound = 0 Then
						    lblArticleTitle.Text = "Aucun article trouvé - arrêt du processus"
						    pnlStatus.StopLed(Color.Red)
						    Running = False
						    btnStart.Visible = True
						    btnStop.Visible = False
						    uiTimer.Stop()
						    statusTimer.Stop()
						    Exit Sub
						End If																							
					lblArticles.Text = $"Articles trouvés :    {ArticlesFound}"																								
                End If
				lblArticleTitle.Text = $"Recherche en cours . . . Merci de patienter  . . .   {ArticlesFound} articles "
				lblArticles.Text = $"Articles trouvés :    {ArticlesFound}"
            Next
        Next
        ArticlesFound = ArticlesUrl.Count
        Dim rnd As New Random()
        Dim i As Integer = 0
        While Running
			If i < 0 OrElse i >= ArticlesUrl.Count Then Exit While
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
                'DeadLinks += 1
            End Try
            Await Task.Delay(rnd.Next(2000, 9000))
            i = (i + 1) Mod ArticlesFound
            If i = 0 Then LoopCount += 1
        End While
    End Sub

    ' ================= STOP =================
    Private Sub StopProcess(sender As Object, e As EventArgs)																												
		'trayIcon.Icon = trayIconStop
		If trayIconStop IsNot Nothing Then
		    trayIcon.Icon = trayIconStop
		Else
		    trayIcon.Icon = Me.Icon
		End If																						
		trayIcon.Text = "Scanner Etsy – Arrêté"
        Running = False
		UpdateTrayTooltip()																											
        btnStart.Visible = True
        btnStop.Visible = False
        statusTimer.Stop()
        pnlStatus.StopLed(Color.Red)
        uiTimer.Stop()
        picThumbnail.Image = Nothing
        lblArticleTitle.Text = ""   
        lblCurrentArticle.Text = ""                                                                                                
    End Sub

    ' ================= UI TIMER =================
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblClicks.Text = $"Clics cumulés :    {TotalClicks}"
        lblArticles.Text = $"Articles trouvés :    {ArticlesFound}"
        'lblDead.Text = $"Liens morts :    {DeadLinks}"
        lblTime.Text = $"Temps activité :    {(DateTime.Now - LoopStartTime):hh\:mm\:ss}"
		UpdateTrayTooltip()																											
    End Sub

																												
' ===================== BARRE DE TITRE CUSTOM =====================

Private Sub AddCustomTitleBar(parent As Control)

    Dim titleBar As New Panel With {
        .Height = 32,
        .Dock = DockStyle.Top,
        .BackColor = Color.FromArgb(245, 245, 245)
    }

    AddHandler titleBar.MouseDown, AddressOf DragWindow

    ' Bouton Minimize
    Dim btnMin As New Button With {
        .Text = "—",
        .Width = 44,
        .Dock = DockStyle.Right,
        .FlatStyle = FlatStyle.Flat
    }
    btnMin.FlatAppearance.BorderSize = 0
    AddHandler btnMin.Click, Sub()
        Me.WindowState = FormWindowState.Minimized
    End Sub

    ' Bouton HideToTray (NOUVELLE COMMANDE)
    Dim btnHide As New Button With {
        .Text = "◉",
        .Width = 44,
        .Dock = DockStyle.Right,
        .FlatStyle = FlatStyle.Flat
    }
    btnHide.FlatAppearance.BorderSize = 0
    AddHandler btnHide.Click, Sub()
        Me.Hide()
        Me.ShowInTaskbar = False
    End Sub

    ' Bouton Close
    Dim btnClose As New Button With {
        .Text = "✕",
        .Width = 44,
        .Dock = DockStyle.Right,
        .FlatStyle = FlatStyle.Flat
    }
    btnClose.FlatAppearance.BorderSize = 0
    AddHandler btnClose.Click, Sub()
        Application.Exit()
    End Sub

    titleBar.Controls.Add(btnClose)
    titleBar.Controls.Add(btnHide)
    titleBar.Controls.Add(btnMin)

    parent.Controls.Add(titleBar)
    titleBar.BringToFront()

End Sub

<DllImport("user32.dll")>
Private Shared Sub ReleaseCapture()
End Sub

<DllImport("user32.dll")>
Private Shared Function SendMessage(
    hWnd As IntPtr,
    msg As Integer,
    wParam As Integer,
    lParam As Integer
) As Integer
End Function

Private Sub DragWindow(sender As Object, e As MouseEventArgs)
    If e.Button = MouseButtons.Left Then
        ReleaseCapture()
        SendMessage(Me.Handle, &HA1, 2, 0)
    End If
End Sub

End Class
