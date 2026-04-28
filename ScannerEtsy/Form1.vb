
Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks

Public Class Form1

    ' ===== CONFIG =====
    Private Const SHOP_URL As String =
        "https://www.etsy.com/fr/shop/mabrocanges?page={0}"

    Private Const MAX_PAGES As Integer = 20

    ' ===== ÉTAT =====
    Private TotalPages As Integer = 0
    Private ArticlesFound As Integer = 0
    Private DeadLinks As Integer = 0
    Private StopRequested As Boolean = False
    Private StartTime As DateTime

    ' ===== HTTP =====
    Private ReadOnly Client As HttpClient =
        New HttpClient With {
            .Timeout = TimeSpan.FromSeconds(20)
        }

    ' ===== REGEX =====
    Private ReadOnly ListingRegex As New Regex(
        "https:\/\/www\.etsy\.com\/listing\/\d+",
        RegexOptions.IgnoreCase)

    ' ===== UI =====
    Private WithEvents btnStart As Button
    Private WithEvents btnStop As Button
    Private lblPages, lblFound, lblDead, lblTime As Label
    Private logBox As ListBox
    Private uiTimer As Timer

    ' ================= FORM =================
    Public Sub New()

        Me.Text = "Etsy Scanner – Headless HTTP"
        Me.Width = 820
        Me.Height = 520
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()
        ConfigureHttpClient()

        StartTime = Date.Now
        uiTimer = New Timer With {.Interval = 500}
        AddHandler uiTimer.Tick, AddressOf RefreshUI
        uiTimer.Start()
    End Sub

    ' ================= INIT UI =================
    Private Sub InitializeUI()

        btnStart = New Button With {.Text = "START", .Left = 30, .Top = 30, .Width = 100}
        btnStop = New Button With {.Text = "STOP", .Left = 140, .Top = 30, .Width = 100}

        lblPages = New Label With {.Left = 270, .Top = 35, .Width = 200}
        lblFound = New Label With {.Left = 270, .Top = 65, .Width = 200}
        lblDead = New Label With {.Left = 270, .Top = 95, .Width = 200}
        lblTime = New Label With {.Left = 270, .Top = 125, .Width = 200}

        logBox = New ListBox With {
            .Left = 30,
            .Top = 170,
            .Width = 740,
            .Height = 280
        }

        Controls.AddRange({btnStart, btnStop, lblPages, lblFound, lblDead, lblTime, logBox})
    End Sub

    ' ================= HTTP CONFIG =================
    Private Sub ConfigureHttpClient()
        Client.DefaultRequestHeaders.UserAgent.ParseAdd(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        Client.DefaultRequestHeaders.Accept.ParseAdd("text/html")
    End Sub

    ' ================= START =================
    Private Async Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        StopRequested = False
        TotalPages = 0
        ArticlesFound = 0
        DeadLinks = 0
        logBox.Items.Clear()

        Log("🚀 Scan démarré")

        Await ScanShopAsync()

        Log("✅ Scan terminé")
    End Sub

    ' ================= STOP =================
    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        StopRequested = True
        Log("⛔ Arrêt demandé")
    End Sub

    ' ================= SCAN SHOP =================
    Private Async Function ScanShopAsync() As Task

        Dim seen As New HashSet(Of String)

        For page = 1 To MAX_PAGES

            If StopRequested Then Exit For

            Dim url = String.Format(SHOP_URL, page)
            Log($"📄 Page {page}")

            Dim html As String = Await GetHtmlAsync(url)
            If String.IsNullOrEmpty(html) Then Exit For

            Dim matches = ListingRegex.Matches(html)
            If matches.Count = 0 Then Exit For

            For Each m As Match In matches
                If seen.Add(m.Value) Then
                    ArticlesFound += 1
                    Await CheckLinkAsync(m.Value)
                End If
            Next

            TotalPages += 1
            Await Task.Delay(800)
        Next

    End Function

    ' ================= HTTP HTML =================
    Private Async Function GetHtmlAsync(url As String) As Task(Of String)
        Try
            Return Await Client.GetStringAsync(url)
        Catch
            Log("❌ Erreur chargement page")
            Return Nothing
        End Try
    End Function

    ' ================= CHECK LINK =================
    Private Async Function CheckLinkAsync(url As String) As Task
        Try
            Dim response = Await Client.SendAsync(
                New HttpRequestMessage(HttpMethod.Head, url))

            If Not response.IsSuccessStatusCode Then
                DeadLinks += 1
                Log("☠ Lien mort : " & url)
            Else
                Log("✅ " & url)
            End If
        Catch
            DeadLinks += 1
            Log("☠ Lien mort : " & url)
        End Try
    End Function

    ' ================= LOG =================
    Private Sub Log(text As String)
        If InvokeRequired Then
            Invoke(Sub() Log(text))
        Else
            logBox.Items.Insert(0, $"{Date.Now:HH:mm:ss} {text}")
        End If
    End Sub

    ' ================= UI REFRESH =================
    Private Sub RefreshUI(sender As Object, e As EventArgs)
        lblPages.Text = $"Pages scannées : {TotalPages}"
        lblFound.Text = $"Articles trouvés : {ArticlesFound}"
        lblDead.Text = $"Liens morts : {DeadLinks}"
        lblTime.Text = $"Temps : {(Date.Now - StartTime):hh\:mm\:ss}"
    End Sub

End Class
