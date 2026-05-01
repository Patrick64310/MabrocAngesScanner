
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class LedPanel
    Inherits Panel

    Public Sub New()
        Me.Size = New Size(60, 60)
        Me.BackColor = Color.Red
        Me.DoubleBuffered = True
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or _
                    ControlStyles.OptimizedDoubleBuffer Or _
                    ControlStyles.UserPaint, True)

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

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias

        Dim r As New Rectangle(2, 2, Me.Width - 4, Me.Height - 4)

        ' LED principale
        Using brush As New SolidBrush(Me.BackColor)
            g.FillEllipse(brush, r)
        End Using

        ' Reflet LED (effet brillant)
        Using highlight As New SolidBrush(Color.FromArgb(70, 255, 255, 255))
            g.FillEllipse(
                highlight,
                r.X + 6,
                r.Y + 6,
                r.Width \ 3,
                r.Height \ 3
            )
        End Using

        ' Bord sombre (anneau extérieur)
        Using pen As New Pen(Color.FromArgb(120, Color.Black), 1)
            g.DrawEllipse(pen, r)
        End Using
    End Sub

End Class
