
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Drawing

Public Class Form1
    Inherits Form

    ' ===== Compteurs =====
    Private TotalClicks As Integer = 0
    Private ArticlesFound As Integer = 0
    Private DeadLinks As Integer = 0
    Private ElapsedStart As DateTime

    ' ===== Contrôle =====
    Private Running As Boolean = False

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
        Me.Text = "Mabroc'Anges – Scanner V6 (Stable)"
        Me.Width = 520
        Me.Height = 300
        Me.StartPosition = FormStartPosition.CenterScreen

        InitializeUI()

        ElapsedStart = DateTime.Now
        uiTimer = New Timer()
        uiTimer.Interval = 1000
        AddHandler uiTimer.Tick, AddressOf UpdateUI
        uiTimer.Start()
    End Sub

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

        Me.Controls.AddRange({lblTotal, lblFound, lblDead, lblTime, btnStart, btnStop, btnReset})
        UpdateUI(Nothing, Nothing)
    End Sub

    ' ===== Simulation Excel-like =====
    Private Sub StartSimulation(sender As Object, e As EventArgs)
        Running = True
        ArticlesFound = 42   ' valeur simulée stable
        SimulateLoop()
    End Sub

    Private Sub StopSimulation(sender As Object, e As EventArgs)
        Running = False
    End Sub

    Private Sub ResetAll(sender As Object, e As EventArgs)
        Running = False
        TotalClicks = 0
        DeadLinks = 0
        ArticlesFound = 0
        ElapsedStart = DateTime.Now
    End Sub

    Private Sub SimulateLoop()
        Dim t As New Threading.Thread(
            Sub()
                While Running
                    TotalClicks += 1
                    If TotalClicks Mod 7 = 0 Then DeadLinks += 1
                    Threading.Thread.Sleep(1200)
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

