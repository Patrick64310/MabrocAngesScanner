
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class Form1
    Inherits Form

    ' ================= HTML EN MÉMOIRE =================
    Private PagesHtml As New Dictionary(Of Integer, String)

    ' ================= LISTES MÉTIER =================
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
    Private chkSaveHtml As CheckBox
    Private chkVisualiserArticles As CheckBox

    ' ================= LOG =================
    Private LogPath As String =
        System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scanner_log.txt")

    ' ================= OPTION FICHIERS HTML =================
    Private PagesHtmlDirectory As String =
        System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pages_html")

    ' ================= REGEX EXCEL =================
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    ' ================= WEBVIEW =================
    Private web As WebView2

    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner V6 (Excel strict • Mémoire)"
        Me.Width = 860
        Me.Height = 420
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        WriteLog("APPLICATION LANCÉE")
    End Sub

    ' ================= UI =================
    Private Sub InitializeUI()

        lblClicks = New Label() With {.Left = 20, .Top = 30, .Width = 800}
        lblArticles = New Label() With {.Left = 20, .Top = 60, .Width = 800}
        lblDead = New Label() With {.Left = 20, .Top = 90, .Width = 800}
        lblTime = New Label() With {.Left = 20, .Top = 120, .Width = 800}

        chkSaveHtml = New CheckBox() With {
            .Text = "Sauvegarder HTML sur disque (debug / audit)",
            .Left = 20,
            .Top = 160,
            .Width = 350
        }

        chkVisualiserArticles = New CheckBox() With {
            .Text = "Visualiser les pages articles",
            .Left = 20,
            .Top = 190,
            .Width = 300,
            .Checked = True
        }

        btnStart = New Button() With {.Text = "START", .Left = 500, .Top = 185, .Width = 140}
        btnStop = New Button() With {.Text = "STOP", .Left = 660, .Top = 185, .Width = 140}

        AddHandler btnStart.Click, AddressOf StartProcessAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        Controls.AddRange({
            lblClicks, lblArticles, lblDead, lblTime,
            chkSaveHtml, chkVisualiserArticles,
            btnStart, btnStop
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

        Await GenerateHtmlPagesAsync()

        ' ===== BOUCLE PAGES =====
        For page As Integer = 1 To 20
            If Not PagesHtml.ContainsKey(page) Then Exit For

            Dim pageUrl =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items"

            WriteLog("PAGE PARCOURUE : " & pageUrl)

            Dim html = PagesHtml(page)
            If html.Contains("Aucun article en vente pour le moment") Then Exit For

            PagesUrl.Add(pageUrl)
        Next

        ' ===== EXTRACTION ARTICLES =====
        For Each kvp In PagesHtml
            For Each m As Match In ListingRegex.Matches(kvp.Value)
                If Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
                    WriteLog("ARTICLE TROUVE : " & m.Value)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        If ArticlesFound = 0 Then Running = False : Exit Sub

        ' ===== NAVIGATION ARTICLES =====
        LoopRunning = True
        LoopStartTime = DateTime.Now
        uiTimer.Start()

        Dim rnd As New Random()

        For Each articleUrl In ArticlesUrl
            If Not Running Then Exit For

            Try
                WriteLog("CLIC ARTICLE : " & articleUrl)
                TotalClicks += 1

                If chkVisualiserArticles.Checked Then
                    Dim psi As New ProcessStartInfo(articleUrl) With {.UseShellExecute = True}
                    Dim p = Process.Start(psi)
                    Threading.Thread.Sleep(1000)
                    If p IsNot Nothing AndAlso Not p.HasExited Then p.Kill()
                    WriteLog("ONGLET FERME : " & articleUrl)
                End If

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

    ' ================= GENERATION HTML =================
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

        For page = 1 To 20
            Dim url =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items"

            web.Source = New Uri(url)
            Await Task.Delay(6000)

            Dim html =
                Await web.ExecuteScriptAsync("document.documentElement.outerHTML")

            html = html.Replace("\""", """")

            PagesHtml(page) = html

            If chkSaveHtml.Checked Then
                Dim filePath =
                    System.IO.Path.Combine(PagesHtmlDirectory, $"page{page}.html")
                File.WriteAllText(filePath, html)
            End If

            If html.Contains("Aucun article en vente pour le moment") Then Exit For
        Next
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

