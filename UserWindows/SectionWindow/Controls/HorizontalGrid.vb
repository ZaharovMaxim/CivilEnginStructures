Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
<System.Reflection.ObfuscationAttribute(Exclude:=True, StripAfterObfuscation:=True)>
Partial Class HorizontalGrid
    Inherits UserControl
    Private m_SectionLayer As CadViewLayer

    Public Sub New()
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.Opaque Or ControlStyles.OptimizedDoubleBuffer, True)
        MyBase.DoubleBuffered = True
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Public Property SectionLayer As CadViewLayer
        Get
            Return m_SectionLayer
        End Get
        Set(ByVal value As CadViewLayer)
            m_SectionLayer = value
        End Set
    End Property

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.Clear(BackColor)

        If SectionLayer IsNot Nothing AndAlso SectionLayer.CadView IsNot Nothing Then
            Dim cadView = SectionLayer.CadView
            Dim offset = cadView.Left - Me.Left
            Dim viewBounds = cadView.ViewBounds
            Dim color As Color

            If BackColor.GetBrightness() >= 0.5 Then
                color = Color.Black
            Else
                color = Color.White
            End If

            Dim size = 1

            While size / cadView.CurrentScale < 25

                size *= 10

            End While

            If size > 1000000 Then
                Return
            End If

            Dim left = CInt((viewBounds.Left / size)) - 1

            Using pen As Pen = New Pen(color)

                Using brush As Brush = New SolidBrush(color)

                    While left * size < viewBounds.Right
                        Dim x As Integer = cadView.ProjectPoint(New Vector2D(left * size, 0)).Point.X + offset
                        Dim s = (left * size).ToString()
                        Dim stringSize = e.Graphics.MeasureString(s, Font)
                        Dim pos = New Point(CInt((x - stringSize.Width / 2)), CInt((ClientSize.Height - Font.Height)))

                        If Me.ClientRectangle.Contains(pos) Then
                            e.Graphics.DrawString(s, Font, brush, pos)
                        End If

                        If (x >= 0) AndAlso (x < ClientSize.Width) Then
                            e.Graphics.DrawLine(pen, x, ClientSize.Height - Font.Height, x, ClientSize.Height - Font.Height - 10)
                        End If

                        If 0.1 * size / cadView.CurrentScale > 3 Then

                            For i As Integer = 0 To 10 - 1
                                Dim xValue = x + CSng((i * 0.1 * size / cadView.CurrentScale))

                                If (xValue >= 0) AndAlso (xValue < ClientSize.Width) Then
                                    e.Graphics.DrawLine(pen, xValue, ClientSize.Height - Font.Height - 3, xValue, ClientSize.Height - Font.Height - 7)
                                End If
                            Next
                        End If

                        left += 1
                    End While

                    Dim y1 = ClientSize.Height - Font.Height - 15
                    Dim y2 = ClientSize.Height - 2 * Font.Height - 15
                    e.Graphics.DrawLine(pen, 0, y1, ClientSize.Width, y1)
                    e.Graphics.DrawLine(pen, 0, y2, ClientSize.Width, y2)
                End Using
            End Using
        End If
    End Sub

    Public Sub DoPaint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Refresh()
    End Sub

    Private Sub HorizontalGrid_Load(sender As Object, e As EventArgs) Handles Me.Load

    End Sub
End Class
