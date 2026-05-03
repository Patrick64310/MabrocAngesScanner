
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class LedPanel
    Inherits Panel

    Private fadeTimer As Timer
    Private intensity As Integer = 80          ' luminosité initiale (0–255)
    Private fadeDirection As Integer = 1       ' 1 = fade in, -1 = fade out
    Private ledColor As Color = Color.Green
    Private isBlinking As Boolean = False

    Public Sub New()
        Me.Width = 60
        Me.Height = 60
        Me.DoubleBuffered = True

        fadeTimer = New Timer With {.Interval = 40} ' ≈25 FPS
        AddHandler fadeTimer.Tick, AddressOf FadeTick
    End Sub

    ' ================= API PUBLIQUE =================

    ' LED clignotante avec fade
    Public Sub StartLed(color As Color)
        ledColor = color
        intensity = 80
        fadeDirection = 1
        isBlinking = True
        fadeTimer.Start()
        Me.Invalidate()
    End Sub

    ' LED fixe (sans clignotement)
    Public Sub StopLed(color As Color)
        fadeTimer.Stop()
        isBlinking = False
        ledColor = color
        intensity = 255
        Me.Invalidate()
    End Sub

    ' ================= ANIMATION =================

    Private Sub FadeTick(sender As Object, e As EventArgs)
        If Not isBlinking Then Return

        intensity += fadeDirection * 6

        If intensity >= 255 Then
            intensity = 255
            fadeDirection = -1
        ElseIf intensity <= 60 Then
            intensity = 60
            fadeDirection = 1
        End If

        Me.Invalidate()
    End Sub

    ' ================= DESSIN =================

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

        Dim rect As New Rectangle(2, 2, Me.Width - 4, Me.Height - 4)

        Using br As New SolidBrush(Color.FromArgb(intensity, ledColor))
            e.Graphics.FillEllipse(br, rect)
        End Using

        Using pen As New Pen(Color.FromArgb(120, Color.Black), 2)
            e.Graphics.DrawEllipse(pen, rect)
        End Using
    End Sub

End Class
