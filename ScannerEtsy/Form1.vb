
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    ' ================= ÉTAT =================
    Private TotalClicks As Integer = 0
    Private DeadLinks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private CurrentPage As Integer = 1

    Private StopRequested As Boolean = False
    Private DebugVisuel As Boolean = False
    Private InitialScanDone As Boolean = False

    ' ================= DONNÉES =================
    Private UrlQueue As New Queue(Of String)

    ' ================= WEBVIEW =================
    Private WithEvents Web As WebView2

    ' ================= REGEX =================
    Private ReadOnly ListingRegex As New Regex(
        "https:\/\/www\.etsy\.com\/fr\/listing\/[^\?\""]+",
        RegexOptions.IgnoreCase)

    ' ================= URL =================
    Private Const ShopBaseUrl As String =
        "https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={0}&sort_order=date_desc#items"

    ' ================= UI =================
    Private lblTotal, lblFound, lblDead, lblPage, lblTime As Label
    Private btnStart, btnStop, btnReset As Button
    Private chkDebug As CheckBox
    Private picLogo As PictureBox
    Private uiTimer As Timer
    Private StartTime As DateTime

    ' ================= FORM =================
    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner Etsy (WebView2)"
        Me.Width = 760
        Me.Height = 440
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()
        InitializeWebView()

        StartTime = Date.Now
        uiTimer = New Timer() With {.Interval = 500}
        AddHandler uiTimer.Tick, AddressOf RefreshUI
        uiTimer.Start()
    End Sub

    ' ================= INIT UI =================
    Private Sub InitializeUI()

        lblTotal = New Label() With {.Left = 30, .Top = 150, .Width = 350}
        lblFound = New Label() With {.Left = 30, .Top = 180, .Width = 350}
        lblDead = New Label() With {.Left = 30, .Top = 210, .Width = 350}
        lblPage = New Label() With {.Left = 30, .Top = 240, .Width = 350}
        lblTime = New Label() With {.Left = 30, .Top = 270, .Width = 350}

        btnStart = New Button() With {.Text = "START", .Left = 450, .Top = 160, .Width = 130}
        btnStop = New Button() With {.Text = "STOP", .Left = 450, .Top = 200, .Width = 130}
        btnReset = New Button() With {.Text = "RESET", .Left = 450, .Top = 240, .Width = 130}

        chkDebug = New CheckBox() With {.Text = "Debug visuel", .Left = 450, .Top = 280}

        AddHandler btnStart.Click, AddressOf StartScan
        AddHandler btnStop.Click, Sub() StopRequested = True
        AddHandler btnReset.Click, AddressOf ResetAll
        AddHandler chkDebug.CheckedChanged, Sub() DebugVisuel = chkDebug.Checked

        picLogo = New PictureBox()
        picLogo.SetBounds(560, 20, 160, 120)
        picLogo.SizeMode = PictureBoxSizeMode.Zoom
        If IO.File.Exists("logo.png") Then
            picLogo.Image = Image.FromFile("logo.png")
        End If

        Me.Controls.AddRange({lblTotal, lblFound, lblDead, lblPage, lblTime,
                              btnStart, btnStop, btnReset, chkDebug, picLogo})
    End Sub

    ' ================= INIT WEBVIEW =================
    Private Async Sub InitializeWebView()
        Web = New WebView2()
        Web.Dock = DockStyle.None
        Web.Visible = DebugVisuel
        Web.Width = 1
        Web.Height = 1
        Me.Controls.Add(Web)
        Await Web.EnsureCoreWebView2Async()
    End Sub

    ' ================= START =================
    Private Sub StartScan(sender As Object, e As EventArgs)

        StopRequested = False
        InitialScanDone = False
        UrlQueue.Clear()

        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        CurrentPage = 1

        NavigateToPage()
    End Sub

    ' ================= RESET =================
    Private Sub ResetAll(sender As Object, e As EventArgs)

        StopRequested = True
        UrlQueue.Clear()

        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        CurrentPage = 1
        InitialScanDone = False
    End Sub

    ' ================= NAVIGATION =================
    Private Sub NavigateToPage()

        If StopRequested Then Exit Sub

        Web.Visible = DebugVisuel
        Web.Source = New Uri(String.Format(ShopBaseUrl, CurrentPage))
    End Sub

    ' ================= PAGE LOADED =================
    Private Async Sub Web_NavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles Web.NavigationCompleted

        If StopRequested Then Exit Sub

        Dim html As String =
            Await Web.ExecuteScriptAsync("document.documentElement.outerHTML")

        html = html.Replace("\""", """")

        Dim matches = ListingRegex.Matches(html)

        If matches.Count = 0 Then Exit Sub

        For Each m As Match In matches
            If Not UrlQueue.Contains(m.Value) Then
                UrlQueue.Enqueue(m.Value)
            End If
        Next

        If Not InitialScanDone Then
            ArticlesFound = UrlQueue.Count
            InitialScanDone = True
        End If

        ProcessQueue()

        CurrentPage += 1
        NavigateToPage()
    End Sub

    ' ================= OPEN URL =================
    Private Sub ProcessQueue()

        While UrlQueue.Count > 0 AndAlso Not StopRequested

            Dim url = UrlQueue.Dequeue()
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
        End While
    End Sub

    ' ================= UI REFRESH =================
    Private Sub RefreshUI(sender As Object, e As EventArgs)

        lblTotal.Text = "Cumul Total clics : " & TotalClicks
        lblFound.Text = "Articles trouvés : " & ArticlesFound
        lblDead.Text = "Liens morts : " & DeadLinks
        lblPage.Text = "Page boutique en cours : " & CurrentPage
        lblTime.Text = "Temps utilisation : " & (Date.Now - StartTime).ToString("hh\:mm\:ss")
    End Sub

End Class


