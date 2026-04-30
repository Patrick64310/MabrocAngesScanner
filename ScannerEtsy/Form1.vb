
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

    ' ================= ICÔNE =================
    Private Sub ApplyWindowIcon()
        Try
            Dim asm = GetType(Form1).Assembly
            For Each resName In asm.GetManifestResourceNames()
                If resName.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) Then
                    Using s = asm.GetManifestResourceStream(resName)
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

    ' ================= UI =================
    Private Sub InitializeUI()

        ' ===== ROOT =====
        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(10)
        }

        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 6))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 320))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        ' ===== HEADER =====
        lblCurrentArticle = New Label With {
            .AutoSize = False,
            .Height = 28,
            .Width = 1050,
            .Font = New Font("Arial", 10),
            .ForeColor = Color.DarkBlue,
            .TextAlign = ContentAlignment.MiddleLeft
        }

        lblArticleTitle = New Label With {
            .AutoSize = False,
            .Height = 56,
            .Width = 1050,
            .Font = New Font("Arial", 11),
            .ForeColor = Color.DarkGreen,
            .Text = "Cliquer sur START pour commencer",
            .TextAlign = ContentAlignment.MiddleLeft
        }

        Dim header As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .WrapContents = False
        }

        header.Controls.Add(lblCurrentArticle)
        header.Controls.Add(lblArticleTitle)
        root.Controls.Add(header, 0, 0)

        root.Controls.Add(New Panel With {
            .Height = 2,
            .Dock = DockStyle.Fill,
            .BackColor = Color.DarkGray
        }, 0, 1)

        ' ===== IMAGES =====
        Dim imagesRow As New TableLayoutPanel With {
            .ColumnCount = 2,
            .Dock = DockStyle.Fill
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

        imagesRow.Controls.Add(picThumbnail, 0, 0)
        imagesRow.Controls.Add(picLogo, 1, 0)
        root.Controls.Add(imagesRow, 0, 2)

        ' ===== BAS =====
        Dim bottomPanel As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 2,
            .Padding = New Padding(10)
        }

        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60))
        bottomPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40))

        ' === Compteurs ===
        Dim fnt As New Font("Arial", 14, FontStyle.Bold)

        Dim counters As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Fill
        }

        lblTime = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblClicks = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblArticles = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblDead = New Label With {.Width = 520, .Height = 30, .Font = fnt}
        lblProgress = New Label With {.Width = 520, .Height = 30, .Font = fnt}

        counters.Controls.AddRange(
            {lblTime, lblClicks, lblArticles, lblDead, lblProgress}
        )

        bottomPanel.Controls.Add(counters, 0, 0)

        ' === ACTIONS ===
        Dim actions As New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown,
            .Dock = DockStyle.Fill,
            .WrapContents = False,
            .MinimumSize = New Size(160, 0),
            .Padding = New Padding(10)
        }

        ' Voyant
        pnlStatus = New Panel With {
            .Width = 100,
            .Height = 60,
            .BackColor = Color.Red,
            .Margin = New Padding(0, 0, 0, 100)
        }
        AddHandler pnlStatus.Paint, AddressOf DrawStatusBorder

        ' Boutons
        btnStart = New Button With {.Text = "START", .Width = 100, .Height = 60, .Font = fnt}
        btnStop = New Button With {.Text = "STOP", .Width = 100, .Height = 60, .Font = fnt, .Visible = False}

        ' Hover Start
        AddHandler btnStart.MouseEnter, Sub() btnStart.BackColor = Color.LightGreen
        AddHandler btnStart.MouseLeave, Sub() btnStart.BackColor = SystemColors.Control

        ' Hover Stop
        AddHandler btnStop.MouseEnter, Sub() btnStop.BackColor = Color.LightCoral
        AddHandler btnStop.MouseLeave, Sub() btnStop.BackColor = SystemColors.Control

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

    ' ===== CADRE NOIR 6px =====
    Private Sub DrawStatusBorder(sender As Object, e As PaintEventArgs)
        Using pen As New Pen(Color.Black, 6)
            e.Graphics.DrawRectangle(pen, 3, 3, pnlStatus.Width - 6, pnlStatus.Height - 6)
        End Using
    End Sub

    ' ===== ANIMATION FADE =====
    Private Sub AnimateStatus(sender As Object, e As EventArgs)
        If Not Running Then Exit Sub
        fadeValue += fadeDir * 8
        If fadeValue >= 255 Then fadeDir = -1
        If fadeValue <= 80 Then fadeDir = 1
        pnlStatus.BackColor = Color.FromArgb(0, fadeValue, 0)
        pnlStatus.Invalidate()
    End Sub

    ' ===== START =====
    Private Async Sub StartAsync(sender As Object, e As EventArgs)
        Running = True
        btnStart.Visible = False
        btnStop.Visible = True
        statusTimer.Start()
        LoopStartTime = DateTime.Now
        uiTimer.Start()
    End Sub

    ' ===== STOP =====
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        btnStart.Visible = True
        btnStop.Visible = False
        statusTimer.Stop()
        pnlStatus.BackColor = Color.Red
        pnlStatus.Invalidate()
    End Sub

    ' ===== UI TIMER =====
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblTime.Text = $"Temps activité : {(DateTime.Now - LoopStartTime):hh\:mm\:ss}"
    End Sub

End Class
