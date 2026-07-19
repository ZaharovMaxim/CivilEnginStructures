Imports System.Text
Imports System.Windows.Forms


Public Class FormRegistration
    Public boolShow As Boolean = False
    Public hasp As String = ""
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        MaskedTextBox1.Mask = "AAAA-AAAA-AAAA-AAAA"
        ' Добавить код инициализации после вызова InitializeComponent().

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim key As String = MaskedTextBox1.Text
        Dim userName As String = TextBox1.Text
        Dim userRng As String = MaskedTextBox1.Tag
        If IsNothing(userRng) = False Then
            userName = userName & userRng
        End If
        Using md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
            Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(userName)
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
            Dim sb As StringBuilder = New StringBuilder()
            If IsArray(hashBytes) = True Then
                For i As Integer = 0 To hashBytes.Length - 1
                    sb.Append(hashBytes(i).ToString("x2"))
                Next i
            End If
            hasp = sb.ToString
        End Using
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Hide()
    End Sub

    Private Sub FormRegistration_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class