
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO

Public Class Form1
    Inherits Form

    ' ================= LISTES =================
    Private PagesUrl As New List(Of String)
    Private ArticlesUrl As New List(Of String)

    ' ================= ÉTAT =================
    Private Running As Boolean = False
    Private LoopRunning As Boolean = False

    Private TotalClicks As Integer = 0
    Private DeadLinks As Integer = 0
    Private ArticlesFound As Integer = 0

    ' ================= TEMPS =================
    Private LoopStartTime As DateTime
    Private uiTimer As Timer

    ' ================= UI =================
    Private lblClicks As Label
    Private lblArticles As Label
    Private lblDead As Label
    Private lblTime As Label
    Private btnStart As Button
    Private btnStop As Button

    ' ================= LOG =================
    Private LogPath As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scanner_log.txt")

    ' ================= REGEX =================
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()

        Me.Text = "Mabroc'Anges – Scanner V6 (Excel strict)"
        Me.Width = 650
        Me.Height = 320
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI
        ' ⚠️ Le timer NE démarre PAS ici

        WriteLog("APPLICATION LANCÉE")
    End Sub

    ' ================= UI =================
    Private Sub InitializeUI()

        lblClicks = New Label() With {.Left = 20, .Top = 30, .Width = 600}
        lblArticles = New Label() With {.Left = 20, .Top = 60, .Width = 600}
        lblDead = New Label() With {.Left = 20, .Top = 90, .Width = 600}
        lblTime = New Label() With {.Left = 20, .Top = 120, .Width = 600}

        btnStart = New Button() With {.Text = "START", .Left = 20, .Top = 200, .Width = 120}
        btnStop = New Button() With {.Text = "STOP", .Left = 160, .Top = 200, .Width = 120}

        AddHandler btnStart.Click, AddressOf StartProcess
        AddHandler btnStop.Click, AddressOf StopProcess

        Controls.AddRange({lblClicks, lblArticles, lblDead, lblTime, btnStart, btnStop})

        UpdateUI(Nothing, Nothing)
    End Sub

    ' ================= LOG =================
    Private Sub WriteLog(msg As String)
        File.AppendAllText(
            LogPath,
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}" & Environment.NewLine
        )
    End Sub

    ' ================= START =================
    Private Sub StartProcess(sender As Object, e As EventArgs)

        If Running Then Exit Sub

        Running = True
        LoopRunning = False

        PagesUrl.Clear()
        ArticlesUrl.Clear()
        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0

        WriteLog("START")

        ' =====================================================
        ' ETAPE 1 — BOUCLE PAGES 1 À 20 (STRICT EXCEL)
        ' =====================================================
        For PageNumber As Integer = 1 To 20

            Dim pageUrl =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={PageNumber}&sort_order=date_desc#items"

            WriteLog("PAGE PARCOURUE : " & pageUrl)

            Dim html As String = LoadPageHtml(PageNumber)

            If html.Contains("Aucun article en vente pour le moment") Then
                WriteLog("ARRET BOUCLE PAGES : Aucun article en vente pour le moment")
                Exit For
            End If

            PagesUrl.Add(pageUrl)

        Next

        ' =====================================================
        ' ETAPE 2 — EXTRACTION DES ARTICLES
        ' =====================================================
        For i As Integer = 0 To PagesUrl.Count - 1

            Dim html As String = LoadPageHtml(i + 1)

            For Each m As Match In ListingRegex.Matches(html)
                ArticlesUrl.Add(m.Value)
                WriteLog("ARTICLE TROUVE : " & m.Value)
            Next

        Next

        ArticlesFound = ArticlesUrl.Count
        WriteLog("ARTICLES TROUVES TOTAL : " & ArticlesFound)

        ' =====================================================
        ' ETAPE 3 — NAVIGATION DES ARTICLES (TIMER ACTIF)
        ' =====================================================
        LoopRunning = True
        LoopStartTime = DateTime.Now
        uiTimer.Start()
        WriteLog("TIMER START")

        Dim rnd As New Random()

        For Each articleUrl In ArticlesUrl

            If Not Running Then Exit For

            Try
                WriteLog("CLIC ARTICLE : " & articleUrl)
                Process.Start(New ProcessStartInfo(articleUrl) With {.UseShellExecute = True})
                TotalClicks += 1
                Threading.Thread.Sleep(1000)
            Catch
                DeadLinks += 1
                WriteLog("LIEN MORT : " & articleUrl)
            End Try

            Threading.Thread.Sleep(rnd.Next(3000, 9000))
        Next

        StopTimerInternal()
        WriteLog("FIN")
    End Sub

    ' ================= STOP =================
    Private Sub StopProcess(sender As Object, e As EventArgs)
        Running = False
        StopTimerInternal()
        WriteLog("STOP")
    End Sub

    ' ================= TIMER STOP =================
    Private Sub StopTimerInternal()
        If LoopRunning Then
            uiTimer.Stop()
            LoopRunning = False
            WriteLog("TIMER STOP")
        End If
    End Sub

    ' ================= SOURCE HTML =================
    ' ATTEND des fichiers page1.html, page2.html, ... dans le dossier de l'EXE
    Private Function LoadPageHtml(pageNumber As Integer) As String

        Dim filePath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"page{pageNumber}.html")

        If File.Exists(filePath) Then
            Return File.ReadAllText(filePath)
        End If

        Return ""
    End Function

    ' ================= UI UPDATE =================
    Private Sub UpdateUI(sender As Object, e As EventArgs)

        lblClicks.Text = "Clics cumulés : " & TotalClicks
        lblArticles.Text = "Articles trouvés : " & ArticlesFound
        lblDead.Text = "Liens morts : " & DeadLinks

        If LoopRunning Then
            lblTime.Text = "Temps activité : " &
                (DateTime.Now - LoopStartTime).ToString("hh\:mm\:ss")
        Else
            lblTime.Text = "Temps activité : 00:00:00"
        End If

    End Sub

End Class
