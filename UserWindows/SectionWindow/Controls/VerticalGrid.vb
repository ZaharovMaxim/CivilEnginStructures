Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
<System.Reflection.ObfuscationAttribute(Exclude:=True, StripAfterObfuscation:=True)>
Partial Class VerticalGrid
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
            Dim viewBounds = cadView.ViewBounds
            Dim color As Color

            If BackColor.GetBrightness() >= 0.5 Then
                color = Color.Black
            Else
                color = Color.White
            End If

            Dim size = 1

            While size / (cadView.CurrentScale / cadView.ScreenRatio) < 25
                '
                size *= 10

            End While

            If size > 1000000 Then
                Return
            End If

            Dim bottom = CInt((viewBounds.Bottom / size)) - 1
            Dim maxWidth As Integer = 0

            Using pen As Pen = New Pen(color)

                Using brush As Brush = New SolidBrush(color)

                    While bottom * size < viewBounds.Top + size
                        Dim y As Integer = cadView.ProjectPoint(New Vector2D(0, bottom * size)).Point.Y
                        Dim s = (bottom * size).ToString()
                        Dim stringSize = e.Graphics.MeasureString(s, Font, 1000, StringFormat.GenericTypographic)
                        Dim pos1 = New Point(CInt((ClientSize.Width - 12 - stringSize.Width)), CInt((y - stringSize.Height / 2)))

                        If Me.ClientRectangle.Contains(pos1) Then
                            e.Graphics.DrawString(s, Font, brush, pos1, StringFormat.GenericTypographic)
                        End If

                        If CInt(stringSize.Width) > maxWidth Then
                            maxWidth = CInt(stringSize.Width)
                        End If

                        If (y >= 0) AndAlso (y < ClientSize.Height) Then
                            e.Graphics.DrawLine(pen, ClientSize.Width, y, ClientSize.Width - 10, y)
                        End If

                        If 0.1 * size / (cadView.CurrentScale / cadView.ScreenRatio) > 3 Then

                            For i As Integer = 0 To 10 - 1
                                Dim yValue = y + CSng((i * 0.1 * size / (cadView.CurrentScale / cadView.ScreenRatio)))

                                If (yValue >= 0) AndAlso (yValue < ClientSize.Height) Then
                                    e.Graphics.DrawLine(pen, ClientSize.Width - 8, yValue, ClientSize.Width - 4, yValue)
                                End If
                            Next
                        End If

                        bottom += 1
                    End While
                End Using
            End Using

            If Me.Width <> maxWidth + 12 Then
                Dim invoke As MethodInvoker = Function()
                                                  Me.Width = maxWidth + 12
                                              End Function

                Parent.BeginInvoke(invoke)
            End If
        End If
    End Sub

    Public Sub DoPaint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Refresh()
    End Sub

    Private Sub VerticalGrid_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
