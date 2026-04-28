
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO

Public Class Form1
    Inherits Form

    ' ===== Compteurs =====
    Private TotalClicks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private DeadLinks As Integer = 0
    Private ElapsedStart As DateTime

    ' ===== Contrôle =====
    Private Running As Boolean = False

    ' ===== Log =====
    Private LogFilePath As String =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scanner_log.txt")

    ' ===== Simulation URLs =====
    Private SimulatedUrls As List(Of String)
    Private UrlIndex As Integer = 0

    ' ===== UI =====
    Private lblTotal As Label
    Private lblFound As Label
    Private lblDead As Label
    Private lblTime As Label
    Private btnStart As Button
    Private btnStop As Button
    Private btnReset As Button
    Private uiTimer As Timer

    Public Sub New()
        Me.Text = "Mabroc'Anges – Scanner V6 (Stable + Logs)"
        Me.Width = 520
        Me.Height = 300
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        ElapsedStart = DateTime.Now
        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI
        uiTimer.Start()

        WriteLog("APPLICATION LANCÉE")

        ' URLs simulées réalistes
        SimulatedUrls = New List(Of String) From {
            "https://www.etsy.com/fr/listing/100000001",
            "https://www.etsy.com/fr/listing/100000002",
            "https://www.etsy.com/fr/listing/100000003",
            "https://www.etsy.com/fr/listing/100000004",
            "https://www.etsy.com/fr/listing/100000005"
        }
    End Sub

    ' ===== UI =====
    Private Sub InitializeUI()

        lblTotal = New Label() With {.Left = 20, .Top = 30, .Width = 400}
        lblFound = New Label() With {.Left = 20, .Top = 60, .Width = 400}
        lblDead = New Label() With {.Left = 20, .Top = 90, .Width = 400}
        lblTime = New Label() With {.Left = 20, .Top = 120, .Width = 400}

        btnStart = New Button() With {.Text = "START", .Left = 20, .Top = 170, .Width = 120}
        btnStop = New Button() With {.Text = "STOP", .Left = 160, .Top = 170, .Width = 120}
        btnReset = New Button() With {.Text = "RESET", .Left = 300, .Top = 170, .Width = 120}

        AddHandler btnStart.Click, AddressOf StartSimulation
        AddHandler btnStop.Click, AddressOf StopSimulation
        AddHandler btnReset.Click, AddressOf ResetAll

        Me.Controls.AddRange(
            {lblTotal, lblFound, lblDead, lblTime, btnStart, btnStop, btnReset})

        UpdateUI(Nothing, Nothing)
    End Sub

    ' ===== LOG =====
    Private Sub WriteLog(message As String)
        Try
            Dim line As String =
                "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "] " & message
            File.AppendAllText(LogFilePath, line & Environment.NewLine)
        Catch
        End Try
    End Sub

    ' ===== Simulation Excel-like =====
    Private Sub StartSimulation(sender As Object, e As EventArgs)
        If Running Then Exit Sub

        Running = True
        ArticlesFound = SimulatedUrls.Count
        UrlIndex = 0

        WriteLog("START")
        WriteLog("PAGE PARCOURUE : https://www.etsy.com/fr/shop/mabrocanges?page=1")
        WriteLog("ARTICLES TROUVÉS : " & ArticlesFound)

        SimulateLoop()
    End Sub

    Private Sub StopSimulation(sender As Object, e As EventArgs)
        Running = False
        WriteLog("STOP")
    End Sub

    Private Sub ResetAll(sender As Object, e As EventArgs)
        Running = False
        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        UrlIndex = 0
        ElapsedStart = DateTime.Now
        WriteLog("RESET")
    End Sub

    ' ===== Boucle simulée =====
    Private Sub SimulateLoop()
        Dim t As New Threading.Thread(
            Sub()
                While Running AndAlso UrlIndex < SimulatedUrls.Count
                    Dim url = SimulatedUrls(UrlIndex)
                    UrlIndex += 1

                    TotalClicks += 1
                    WriteLog("CLIC URL : " & url)

                    ' Simuler un lien mort tous les 3
                    If TotalClicks Mod 3 = 0 Then
                        DeadLinks += 1
                        WriteLog("LIEN MORT : " & url)
                    End If

                    Threading.Thread.Sleep(1500)
                End While
            End Sub)
        t.IsBackground = True
        t.Start()
    End Sub

    ' ===== UI Refresh =====
    Private Sub UpdateUI(sender As Object, e As EventArgs)
        lblTotal.Text = "Cumul Total clics : " & TotalClicks
        lblFound.Text = "Articles trouvés : " & ArticlesFound
        lblDead.Text = "Liens morts : " & DeadLinks
        lblTime.Text = "Temps utilisation : " &
                       (DateTime.Now - ElapsedStart).ToString("hh\:mm\:ss")
    End Sub

End Class
