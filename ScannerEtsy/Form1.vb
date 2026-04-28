
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class Form1
    Inherits Form

    ' ================= MÉMOIRE =================
    Private PagesHtml As New Dictionary(Of Integer, String)
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
    Private lblClicks, lblArticles, lblDead, lblTime As Label
    Private btnStart, btnStop As Button
    Private chkSaveHtml As CheckBox

    ' ================= LOG =================
    Private LogPath As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scanner_log.txt")

    ' ================= OPTION HTML =================
    Private PagesHtmlDirectory As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pages_html")

    ' ================= REGEX EXCEL =================
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    ' ================= WEBVIEW =================
    Private web As WebView2

    Public Sub New()

        Me.Text = "Mabroc'Anges – Scanner V6 (Excel strict • Mémoire)"
        Me.Width = 780
        Me.Height = 360
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        WriteLog("APPLICATION LANCÉE")
    End Sub

    ' ================= UI =================
    Private Sub InitializeUI()

        lblClicks = New Label() With {.Left = 20, .Top = 30, .Width = 720}
        lblArticles = New Label() With {.Left = 20, .Top = 60, .Width = 720}
        lblDead = New Label() With {.Left = 20, .Top = 90, .Width = 720}
        lblTime = New Label() With {.Left = 20, .Top = 120, .Width = 720}

        chkSaveHtml = New CheckBox() With {
            .Text = "Sauvegarder HTML sur disque (debug / audit)",
            .Left = 20,
            .Top = 160,
            .Width = 350
        }

        btnStart = New Button() With {.Text = "START", .Left = 420, .Top = 155, .Width = 120}
        btnStop = New Button() With {.Text = "STOP", .Left = 560, .Top = 155, .Width = 120}

        AddHandler btnStart.Click, AddressOf StartProcessAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        Controls.AddRange({
            lblClicks, lblArticles, lblDead, lblTime,
            chkSaveHtml, btnStart, btnStop
        })

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
    Private Async Sub StartProcessAsync(sender As Object, e As EventArgs)

        If Running Then Exit Sub
        Running = True

        PagesHtml.Clear()
        PagesUrl.Clear()
        ArticlesUrl.Clear()

        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0

        WriteLog("START")

        ' ========= ÉTAPE 0 — GÉNÉRATION HTML EN MÉMOIRE =========
        Await GenerateHtmlPagesAsync()

        ' ========= ÉTAPE 1 — CONSTRUCTION DES PAGES =========
        For page = 1 To 20

            Dim pageUrl =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items"

            WriteLog("PAGE PARCOURUE : " & pageUrl)

            If Not PagesHtml.ContainsKey(page) Then Continue For

            Dim html = PagesHtml(page)

            If html.Contains("Aucun article en vente pour le moment") Then
                WriteLog("ARRET BOUCLE PAGES : Aucun article")
                Exit For
            End If

            PagesUrl.Add(pageUrl)
        Next

        ' ========= ÉTAPE 2 — EXTRACTION DES ARTICLES =========
        For Each kvp In PagesHtml

            For Each m As Match In ListingRegex.Matches(kvp.Value)
                If Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
                    WriteLog("ARTICLE TROUVE : " & m.Value)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        WriteLog("ARTICLES TROUVES TOTAL : " & ArticlesFound)

        If ArticlesFound = 0 Then
            WriteLog("AUCUN ARTICLE — FIN")
            Running = False
            Exit Sub
        End If

        ' ========= ÉTAPE 3 — NAVIGATION ARTICLES =========
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

    Private Sub StopTimerInternal()
        If LoopRunning Then
            uiTimer.Stop()
            LoopRunning = False
            WriteLog("TIMER STOP")
        End If
    End Sub

    ' ================= GÉNÉRATION HTML (WEBVIEW → MÉMOIRE) =================
    Private Async Function GenerateHtmlPagesAsync() As Task

        If web Is Nothing Then
            web = New WebView2()
            web.Visible = False
            Controls.Add(web)
            Await web.EnsureCoreWebView2Async()
        End If

        If chkSaveHtml.Checked Then
            Directory.CreateDirectory(PagesHtmlDirectory)
        End If

        WriteLog("DEBUT GENERATION HTML (MEMOIRE)")

        For page = 1 To 20

            Dim url =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items"

            WriteLog("NAVIGATION WEBVIEW : " & url)
            web.Source = New Uri(url)

            Await Task.Delay(6000)

            Dim html =
                Await web.ExecuteScriptAsync("document.documentElement.outerHTML")

            html = html.Replace("\""", """")

            PagesHtml(page) = html
            WriteLog("HTML CHARGE EN MEMOIRE : page " & page)

            If chkSaveHtml.Checked Then
                Dim path = Path.Combine(PagesHtmlDirectory, $"page{page}.html")
                File.WriteAllText(path, html)
                WriteLog("HTML SAUVEGARDE SUR DISQUE : " & path)
            End If

            If html.Contains("Aucun article en vente pour le moment") Then
                WriteLog("ARRET GENERATION HTML : Aucun article")
                Exit For
            End If
        Next

        WriteLog("FIN GENERATION HTML (MEMOIRE)")
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
