
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class LedPanel
    Inherits Panel

    Private pulseValue As Integer = 0
    Private pulseDir As Integer = 1
    Private WithEvents pulseTimer As New Timer With {.Interval = 40}

    Public Property LedColor As Color = Color.Green


    Public Sub New()
        Me.Size = New Size(70, 70)
        Me.BackColor = Color.Black
        Me.DoubleBuffered = True
    
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or _
                    ControlStyles.OptimizedDoubleBuffer Or _
                    ControlStyles.UserPaint, True)
    
        LedColor = Color.Red      ' 🔴 LED rouge par défaut
        pulseTimer.Stop()         ' ❌ pas d’animation au démarrage
    
        UpdateRegion()
    End Sub


    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        UpdateRegion()
    End Sub

    Private Sub UpdateRegion()
        Dim path As New GraphicsPath()
        path.AddEllipse(0, 0, Me.Width, Me.Height)
        Me.Region = New Region(path)
    End Sub

    Private Sub PulseTick(sender As Object, e As EventArgs) Handles pulseTimer.Tick
        pulseValue += pulseDir * 5
        If pulseValue >= 60 Then pulseDir = -1
        If pulseValue <= 0 Then pulseDir = 1
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias

        Dim center = New Point(Me.Width \ 2, Me.Height \ 2)

        ' ===== HALO LUMINEUX EXTERNE =====
        For i = 1 To 6
            Dim alpha = Math.Max(0, 35 - i * 5 + pulseValue \ 4)
            Using halo As New SolidBrush(Color.FromArgb(alpha, LedColor))
                g.FillEllipse(
                    halo,
                    center.X - (30 + i),
                    center.Y - (30 + i),
                    (60 + i * 2),
                    (60 + i * 2)
                )
            End Using
        Next

        ' ===== ANNEAU MÉTAL INDUSTRIEL =====
        Dim metalRect As New Rectangle(2, 2, Me.Width - 4, Me.Height - 4)
        Using metalBrush As New LinearGradientBrush(
            metalRect,
            Color.FromArgb(180, 180, 180),
            Color.FromArgb(80, 80, 80),
            LinearGradientMode.Vertical)
            g.FillEllipse(metalBrush, metalRect)
        End Using

        ' ===== LED CENTRALE =====
        Dim ledRect As New Rectangle(8, 8, Me.Width - 16, Me.Height - 16)
        Using ledBrush As New LinearGradientBrush(
            ledRect,
            ControlPaint.Light(LedColor),
            ControlPaint.Dark(LedColor),
            LinearGradientMode.Vertical)
            g.FillEllipse(ledBrush, ledRect)
        End Using

        ' ===== REFLET BRILLANT =====
        Using highlight As New SolidBrush(Color.FromArgb(70, Color.White))
            g.FillEllipse(
                highlight,
                ledRect.X + 6,
                ledRect.Y + 5,
                ledRect.Width \ 2,
                ledRect.Height \ 3
            )
        End Using

        ' ===== CONTOUR MÉTAL =====
        Using pen As New Pen(Color.FromArgb(140, 0, 0, 0), 1)
            g.DrawEllipse(pen, metalRect)
        End Using
    End Sub


Public Sub StartLed(color As Color)
    LedColor = color
    pulseTimer.Start()
    Me.Invalidate()
End Sub

Public Sub StopLed(color As Color)
    LedColor = color
    pulseTimer.Stop()
    Me.Invalidate()
End Sub      
                
End Class
