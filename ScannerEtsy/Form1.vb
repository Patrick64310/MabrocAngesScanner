
Imports System
Imports System.Net
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

Public Class Form1
    Inherits Form

    ' ================= ÉTAT =================
    Private TotalClicks As Integer = 0
    Private DeadLinks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private CurrentPage As Integer = 1

    Private StopRequested As Boolean = False
    Private DebugVisuel As Boolean = False

    ' ================= TEMPS =================
    Private StartTime As DateTime
    Private uiTimer As System.Windows.Forms.Timer

    ' ================= DONNÉES =================
    Private UrlList As New List(Of String)

    ' ================= CONSTANTES =================
    Private Const ShopBaseUrl As String =
        "https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={0}&sort_order=date_desc#items"

    Private ReadOnly ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    ' ================= UI =================
    Private lblTotal As Label
    Private lblFound As Label
    Private lblDead As Label
    Private lblPage As Label
    Private lblTime As Label

    Private btnStart As Button
    Private btnStop As Button
    Private btnReset As Button
    Private chkDebug As CheckBox
    Private picLogo As PictureBox

    ' ================= FORM =================
    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner Etsy"
        Me.Width = 720
        Me.Height = 440
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        StartTime = DateTime.Now
        uiTimer = New System.Windows.Forms.Timer()
        uiTimer.Interval = 500
        AddHandler uiTimer.Tick, AddressOf RefreshUI
        uiTimer.Start()
    End Sub

    ' ================= UI INIT =================
    Private Sub InitializeUI()

        lblTotal = New Label() With {.Left = 30, .Top = 140, .Width = 350}
        lblFound = New Label() With {.Left = 30, .Top = 170, .Width = 350}
        lblDead = New Label() With {.Left = 30, .Top = 200, .Width = 350}
        lblPage = New Label() With {.Left = 30, .Top = 230, .Width = 350}
        lblTime = New Label() With {.Left = 30, .Top = 260, .Width = 350}

        btnStart = New Button() With {.Text = "START", .Left = 430, .Top = 170, .Width = 130}
        btnStop = New Button() With {.Text = "STOP", .Left = 430, .Top = 210, .Width = 130}
        btnReset = New Button() With {.Text = "RESET", .Left = 430, .Top = 250, .Width = 130}

        chkDebug = New CheckBox() With {
            .Text = "Debug visuel",
            .Left = 430,
            .Top = 290,
            .Width = 160
        }

        AddHandler btnStart.Click, AddressOf StartScan
        AddHandler btnStop.Click, AddressOf StopScan
        AddHandler btnReset.Click, AddressOf ResetAll
        AddHandler chkDebug.CheckedChanged, Sub() DebugVisuel = chkDebug.Checked

        ' Logo
        picLogo = New PictureBox()
        picLogo.SetBounds(520, 20, 128, 128)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        If File.Exists("logo.png") Then
            picLogo.Image = Image.FromFile("logo.png")
        End If

        Me.Controls.AddRange(New Control() {
            lblTotal, lblFound, lblDead, lblPage, lblTime,
            btnStart, btnStop, btnReset, chkDebug, picLogo
        })

        RefreshUI(Nothing, Nothing)
    End Sub

    ' ================= UI UPDATE =================
    Private Sub RefreshUI(sender As Object, e As EventArgs)
        lblTotal.Text = $"Cumul Total clics : {TotalClicks}"
        lblFound.Text = $"Articles trouvés : {ArticlesFound}"
        lblDead.Text = $"Liens morts : {DeadLinks}"
        lblPage.Text = $"Page boutique en cours : {CurrentPage}"
        lblTime.Text = $"Temps utilisation : {(DateTime.Now - StartTime):hh\:mm\:ss}"
    End Sub

    ' ================= CONTROLES =================
    Private Sub StopScan(sender As Object, e As EventArgs)
        StopRequested = True
    End Sub

    Private Sub ResetAll(sender As Object, e As EventArgs)
        StopRequested = True
        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        CurrentPage = 1
        UrlList.Clear()
    End Sub

    ' ================= TELECHARGEMENT HTML (clé) =================
    Private Function DownloadHtml(url As String) As String

        ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

        Dim req = CType(WebRequest.Create(url), HttpWebRequest)

        req.Method = "GET"
        req.AllowAutoRedirect = True
        req.CookieContainer = New CookieContainer()
        req.UserAgent =
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " &
            "AppleWebKit/537.36 (KHTML, like Gecko) " &
            "Chrome/123.0.0.0 Safari/537.36"

        req.Headers.Add("Accept-Language", "fr-FR,fr;q=0.9")
        req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"

        Using resp = CType(req.GetResponse(), HttpWebResponse)
            Using sr As New StreamReader(resp.GetResponseStream())
                Return sr.ReadToEnd()
            End Using
        End Using

    End Function

    ' ================= BOUCLE SYNCHRONE (COMME EXCEL) =================
    Private Sub StartScan(sender As Object, e As EventArgs)

        StopRequested = False
        UrlList.Clear()
        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        CurrentPage = 1

        ' === 1. SCAN INITIAL (snapshot Excel) ===
        Dim html As String

        Try
            html = DownloadHtml(String.Format(ShopBaseUrl, CurrentPage))
        Catch
            MessageBox.Show("Impossible de charger la boutique Etsy.")
            Exit Sub
        End Try

        Dim matches = ListingRegex.Matches(html)
        For Each m As Match In matches
            UrlList.Add(m.Groups(1).Value)
        Next

        ArticlesFound = UrlList.Count
        RefreshUI(Nothing, Nothing)

        ' === 2. BOUCLE PRINCIPALE ===
        Dim index As Integer = 0

        Do While Not StopRequested AndAlso index < UrlList.Count

            Dim url = UrlList(index)
            index += 1

            TotalClicks += 1

            Try
                If DebugVisuel Then
                    Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
                    Threading.Thread.Sleep(2000)
                Else
                    Dim p As New Process()
                    p.StartInfo.FileName = url
                    p.StartInfo.UseShellExecute = True
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                    p.Start()
                    Threading.Thread.Sleep(1000)
                    If Not p.HasExited Then p.Kill()
                End If
            Catch
                DeadLinks += 1
            End Try

            RefreshUI(Nothing, Nothing)
            Threading.Thread.Sleep(5000)
        Loop
    End Sub

End Class


