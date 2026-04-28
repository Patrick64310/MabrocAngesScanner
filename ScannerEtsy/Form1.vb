
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class Form1
    Inherits Form

    ' ========== HTML EN MÉMOIRE ==========
    Private PagesHtml As New Dictionary(Of Integer, String)

    ' ========== LISTES MÉTIER ==========
    Private PagesUrl As New List(Of String)
    Private ArticlesUrl As New List(Of String)

    ' ========== ÉTAT ==========
    Private Running As Boolean = False
    Private LoopRunning As Boolean = False
    Private TotalClicks As Integer = 0
    Private DeadLinks As Integer = 0
    Private ArticlesFound As Integer = 0

    ' ========== TEMPS ==========
    Private LoopStartTime As DateTime
    Private uiTimer As Timer

    ' ========== UI ==========
    Private lblClicks As Label
    Private lblArticles As Label
    Private lblDead As Label
    Private lblTime As Label
    Private btnStart As Button
    Private btnStop As Button
    Private chkSaveHtml As CheckBox
    Private chkVisualiserArticles As CheckBox

    ' ========== WEBVIEW ==========
    Private webPages As WebView2        ' génération HTML pages boutique
    Private webArticle As WebView2      ' affichage article 1 seconde

    ' ========== LOG ==========
    Private LogPath As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scanner_log.txt")

    Private PagesHtmlDirectory As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pages_html")

    ' ========== REGEX EXCEL ==========
    Private ListingRegex As New Regex(
        "(https:\/\/www\.etsy\.com\/fr\/listing\/[^\?]+)",
        RegexOptions.IgnoreCase)

    Public Sub New()

        Me.Text = "Mabroc'Anges – Scanner V6 (Excel strict / WebView2)"
        Me.Width = 900
        Me.Height = 520
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI

        WriteLog("APPLICATION LANCÉE")
    End Sub

    ' ========== UI ==========
    Private Sub InitializeUI()

        lblClicks = New Label() With {.Left = 20, .Top = 20, .Width = 850}
        lblArticles = New Label() With {.Left = 20, .Top = 45, .Width = 850}
        lblDead = New Label() With {.Left = 20, .Top = 70, .Width = 850}
        lblTime = New Label() With {.Left = 20, .Top = 95, .Width = 850}

        chkSaveHtml = New CheckBox() With {
            .Text = "Sauvegarder HTML (debug / audit)",
            .Left = 20,
            .Top = 130,
            .Width = 320
        }

        chkVisualiserArticles = New CheckBox() With {
            .Text = "Visualiser les pages articles (1 seconde)",
            .Left = 20,
            .Top = 160,
            .Width = 360,
            .Checked = True
        }

        btnStart = New Button() With {.Text = "START", .Left = 450, .Top = 145, .Width = 140}
        btnStop = New Button() With {.Text = "STOP", .Left = 620, .Top = 145, .Width = 140}

        AddHandler btnStart.Click, AddressOf StartProcessAsync
        AddHandler btnStop.Click, AddressOf StopProcess

        ' WebView2 invisible (pages boutique)
        webPages = New WebView2() With {.Visible = False}
        Controls.Add(webPages)

        ' WebView2 pour affichage articles
        webArticle = New WebView2() With {
            .Left = 20,
            .Top = 200,
            .Width = 850,
            .Height = 260,
            .Visible = False
        }
        Controls.Add(webArticle)

        Controls.AddRange({
            lblClicks, lblArticles, lblDead, lblTime,
            chkSaveHtml, chkVisualiserArticles,
            btnStart, btnStop
        })

        UpdateUI(Nothing, Nothing)
    End Sub

    ' ========== LOG ==========
    Private Sub WriteLog(msg As String)
        File.AppendAllText(
            LogPath,
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}" & Environment.NewLine)
    End Sub

    ' ========== START ==========
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

        ' Initialisation WebView2
        Await webPages.EnsureCoreWebView2Async()
        Await webArticle.EnsureCoreWebView2Async()

        ' ====== GÉNÉRATION HTML PAGES ======
        If chkSaveHtml.Checked Then
            Directory.CreateDirectory(PagesHtmlDirectory)
        End If

        For page = 1 To 20
            Dim url =
                $"https://www.etsy.com/fr/shop/mabrocanges?ref=items-pagination&page={page}&sort_order=date_desc#items"

            WriteLog("NAVIGATION PAGE : " & url)
            webPages.Source = New Uri(url)
            Await Task.Delay(6000)

            Dim html As String =
                Await webPages.ExecuteScriptAsync("document.documentElement.outerHTML")

            html = html.Replace("\""", """")

            PagesHtml(page) = html

            If chkSaveHtml.Checked Then
                Dim filePath As String =
                    Path.Combine(PagesHtmlDirectory, $"page{page}.html")
                File.WriteAllText(filePath, html)
            End If

            If html.Contains("Aucun article en vente pour le moment") Then
                Exit For
            End If
        Next

        ' ====== EXTRACTION ARTICLES ======
        For Each kvp In PagesHtml
            For Each m As Match In ListingRegex.Matches(kvp.Value)
                If Not ArticlesUrl.Contains(m.Value) Then
                    ArticlesUrl.Add(m.Value)
                    WriteLog("ARTICLE TROUVÉ : " & m.Value)
                End If
            Next
        Next

        ArticlesFound = ArticlesUrl.Count
        If ArticlesFound = 0 Then Running = False : Exit Sub

        ' ====== NAVIGATION ARTICLES ======
        LoopRunning = True
        LoopStartTime = DateTime.Now
        uiTimer.Start()

        Dim rnd As New Random()

        For Each articleUrl In ArticlesUrl
            If Not Running Then Exit For

            WriteLog("CLIC ARTICLE : " & articleUrl)
            TotalClicks += 1

            If chkVisualiserArticles.Checked Then
                webArticle.Visible = True
                webArticle.Source = New Uri(articleUrl)

                Await Task.Delay(1000)

                webArticle.Source = New Uri("about:blank")
                webArticle.Visible = False
            End If

            Await Task.Delay(rnd.Next(3000, 9000))
        Next

        StopTimerInternal()
        WriteLog("FIN")
    End Sub

    ' ========== STOP ==========
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

    ' ========== UI UPDATE ==========
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

