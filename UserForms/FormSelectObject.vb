Imports System.Windows.Forms
Public Class FormSelectObject
    Public boolShow As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ' Добавить код инициализации после вызова InitializeComponent().
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub FormSelectObject_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim userSize As Drawing.Size = Me.Size
        GroupBox1.Size = New Drawing.Size(userSize.Width - 50, userSize.Height - 160)
        ListBox1.Size = New Drawing.Size(userSize.Width - 70, userSize.Height - 170)
        Dim srButton As Double = userSize.Width / 2
        Button1.Location = New Drawing.Point(srButton - 260, userSize.Height - 120)
        Button2.Location = New Drawing.Point(srButton - 30, userSize.Height - 120)
    End Sub

    Private Sub FormSelectObject_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class