
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class LedPanel
    Inherits Panel

    Public Sub New()
        Me.Width = 60
        Me.Height = 60
        Me.BackColor = Color.Red
        Me.DoubleBuffered = True
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias

        Dim r As New Rectangle(2, 2, Me.Width - 4, Me.Height - 4)

        ' Ombre externe
        Using shadow As New SolidBrush(Color.FromArgb(40, Color.Black))
            g.FillEllipse(shadow, r.X - 2, r.Y - 2, r.Width + 4, r.Height + 4)
        End Using

        ' LED principale
        Using brush As New SolidBrush(Me.BackColor)
            g.FillEllipse(brush, r)
        End Using

        ' Reflet
        Using highlight As New SolidBrush(Color.FromArgb(60, Color.White))
            g.FillEllipse(
                highlight,
                r.X + 6,
                r.Y + 6,
                r.Width \ 3,
                r.Height \ 3
            )
        End Using
    End Sub
End Class

