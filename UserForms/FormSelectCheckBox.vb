Imports System.Windows.Forms

Public Class FormSelectCheckBox
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Close()
    End Sub

    Private Sub FormSelectCheckBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ВыбратьВсеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьВсеToolStripMenuItem.Click
        If CheckedListBox1.Items.Count > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        End If
    End Sub

    Private Sub ОтменитьВыборToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтменитьВыборToolStripMenuItem.Click
        If CheckedListBox1.Items.Count > 0 Then
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub
End Class