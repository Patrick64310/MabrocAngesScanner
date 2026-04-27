
Imports System
Imports System.Net
Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Form1
    Inherits Form

    ' ====== Comptage / état ======
    Private TotalClicks As Integer = 0
    Private DeadLinks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private CurrentPage As Integer = 1

    ' ====== Contrôle exécution ======
    Private cts As CancellationTokenSource
    Private sw As New Stopwatch()
    Private rnd As New Random()
    Private DebugVisuel As Boolean = False

    ' ====== Données ======
    Private UrlList As New List(Of String)

    ' Boutique Etsy (identique Excel)
    Private Const ShopBaseUrl As String =
        "https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&sort_order=date_desc&page={0}#items"

    ' Regex STRICTEMENT identique à Excel
    Private ReadOnly ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    ' ====== UI ======
    Private lblTotal As Label
    Private lblFound As Label
    Private lblDead As Label
    Private lblTime As Label
    Private lblPage As Label

    Private btnStart As Button
    Private btnStop As Button
    Private btnReset As Button
    Private chkDebug As CheckBox

    Private uiTimer As New System.Windows.Forms.Timer()

    ' ====== Constructeur ======
    Public Sub New()
        ' Fenêtre (MINIMUM OBLIGATOIRE)
        Me.Text = "Mabroc'Anges – Scanner Etsy"
        Me.Width = 640
        Me.Height = 360
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ShowInTaskbar = True

        ' Initialisation interface
        InitializeUI()

        ' Temps d'utilisation
        sw.Start()
        AddHandler uiTimer.Tick, AddressOf OnUiTimer
        uiTimer.Interval = 500
        uiTimer.Start()
    End Sub

    ' ====== Initialisation UI ======
    Private Sub InitializeUI()
        lblTotal = New Label() With {.Left = 20, .Top = 20, .Width = 300}
        lblFound = New Label() With {.Left = 20, .Top = 45, .Width = 300}
        lblDead = New Label() With {.Left = 20, .Top = 70, .Width = 300}
        lblPage = New Label() With {.Left = 20, .Top = 95, .Width = 300}
        lblTime = New Label() With {.Left = 20, .Top = 120, .Width = 300}

        btnStart = New Button() With {.Text = "START", .Left = 360, .Top = 40, .Width = 100}
        btnStop = New Button() With {.Text = "STOP", .Left = 360, .Top = 75, .Width = 100}
        btnReset = New Button() With {.Text = "RESET", .Left = 360, .Top = 110, .Width = 100}

        chkDebug = New CheckBox() With {
            .Text = "Debug visuel",
            .Left = 360,
            .Top = 150,
            .Width = 120
        }

        AddHandler btnStart.Click, AddressOf OnStart
        AddHandler btnStop.Click, AddressOf OnStop
        AddHandler btnReset.Click, AddressOf OnReset
        AddHandler chkDebug.CheckedChanged, Sub() DebugVisuel = chkDebug.Checked

        Me.Controls.AddRange(New Control() {
            lblTotal, lblFound, lblDead, lblPage, lblTime,
            btnStart, btnStop, btnReset, chkDebug
        })

        RefreshLabels()
    End Sub

    ' ====== UI Tick ======
    Private Sub OnUiTimer(sender As Object, e As EventArgs)
        RefreshLabels()
    End Sub

    Private Sub RefreshLabels()
        lblTotal.Text = $"Cumul Total clics : {TotalClicks}"
        lblFound.Text = $"Articles trouvés : {ArticlesFound}"
        lblDead.Text = $"Liens morts : {DeadLinks}"
        lblPage.Text = $"Page boutique en cours : {CurrentPage}"
        lblTime.Text = $"Temps utilisation : {sw.Elapsed:hh\:mm\:ss}"
    End Sub

    ' ====== Boutons ======
    Private Sub OnStart(sender As Object, e As EventArgs)
        btnStart.Enabled = False
        btnStop.Enabled = True

        cts = New CancellationTokenSource()

        ' Toujours recréer la liste (Excel-like)
        UrlList.Clear()
        ArticlesFound = 0
        CurrentPage = 1

        ' Lancer découverte + boucle ouverture
        Task.Run(Sub() DiscoverUrls(cts.Token))
        Task.Run(Sub() OpenLinksLoop(cts.Token))
    End Sub

    Private Sub OnStop(sender As Object, e As EventArgs)
        If cts IsNot Nothing Then cts.Cancel()
        btnStart.Enabled = True
        btnStop.Enabled = False
    End Sub

    Private Sub OnReset(sender As Object, e As EventArgs)
        OnStop(Nothing, Nothing)

        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        CurrentPage = 1
        UrlList.Clear()

        RefreshLabels()
    End Sub

    ' ====== Découverte URLs (infinie, STOP only) ======
    Private Sub DiscoverUrls(token As CancellationToken)
        Do While Not token.IsCancellationRequested
            Dim html As String = ""
            Dim pageUrl = String.Format(ShopBaseUrl, CurrentPage)

            Try
                Using wc As New WebClient()
                    wc.Headers.Add("User-Agent", "Mozilla/5.0")
                    html = wc.DownloadString(pageUrl)
                End Using
            Catch
                CurrentPage += 1
                Continue Do
            End Try

            Dim matches = ListingRegex.Matches(html)
            For Each m As Match In matches
                Dim u = m.Groups(1).Value
                SyncLock UrlList
                    If Not UrlList.Contains(u) Then
                        UrlList.Add(u)
                        ArticlesFound += 1
                    End If
                End SyncLock
            Next

            CurrentPage += 1
            Thread.Sleep(rnd.Next(3000, 6000))
        Loop
    End Sub

    ' ====== Ouverture URLs (minimisé, 1s, STOP only) ======
    Private Sub OpenLinksLoop(token As CancellationToken)
        Dim index As Integer = 0

        Do While Not token.IsCancellationRequested
            Dim url As String = Nothing

            SyncLock UrlList
                If index < UrlList.Count Then
                    url = UrlList(index)
                    index += 1
                End If
            End SyncLock

            If url Is Nothing Then
                Thread.Sleep(1500)
                Continue Do
            End If

            Try
                If DebugVisuel Then
                    Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
                    Thread.Sleep(2000)
                Else
                    Dim p As New Process()
                    p.StartInfo.FileName = url
                    p.StartInfo.UseShellExecute = True
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
                    p.Start()

                    Thread.Sleep(1000)

                    If Not p.HasExited Then p.Kill()
                End If

                TotalClicks += 1

            Catch
                DeadLinks += 1
                Thread.Sleep(30000)
            End Try

            Thread.Sleep(rnd.Next(4000, 8000))
        Loop
    End Sub

End Class
