Imports System.IO

Public Class FormUserTemptate
    Public boolWriteFile As Boolean = False
    Public templateDir As String
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ApplyModernAppearance()
    End Sub

    Private Sub ListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedValueChanged
        Dim nameSheme As String = ListBox1.SelectedItem
        If IsNothing(nameSheme) = False Then
            If nameSheme.Trim.Length > 0 Then
                TextBox1.Text = nameSheme.Trim
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'записать файл
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim boolWrite As Boolean = True
        Dim nameFile As String = TextBox1.Text
        If IsNothing(nameFile) = False Then
            If nameFile.Trim.Length > 0 Then
                For i As Integer = 0 To ListBox1.Items.Count - 1
                    If ListBox1.Items.Item(i).ToString.Trim Like nameFile.Trim Then
                        Dim rez As MsgBoxResult = MsgBox("Такое имя уже существует. Перезаписать схему?", vbOKCancel, "Имя схемы раскладки балок")
                        If rez = MsgBoxResult.Cancel Then
                            boolWrite = False
                            Exit For
                        End If
                    End If
                Next i
            Else
                boolWrite = False
            End If
        Else
            boolWrite = False
        End If

        If boolWrite = True Then
            boolWriteFile = True
            Me.Hide()
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'удалить файл
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim templateDir As String = Label3.Text
        If Directory.Exists(templateDir) = True Then
            Dim allFiles As String() = Directory.GetFiles(templateDir)
            Dim arrayFilesTemplate As String() = Nothing
            Dim count As Integer = 0
            For i As Integer = 0 To allFiles.Length - 1
                Dim nameFile As String = Path.GetFileNameWithoutExtension(allFiles(i))
                If nameFile.Trim Like TextBox1.Text.Trim() Then
                    If File.Exists(allFiles(i)) = True Then
                        Dim rezMsg As MsgBoxResult = MsgBox("Вы уверены что хотите удалить файл: " & nameFile, MsgBoxStyle.OkCancel, "Удаление схемы раскладки балок")
                        If rezMsg = MsgBoxResult.Ok Then
                            File.Delete(allFiles(i))
                        End If
                    End If
                End If
            Next i

            allFiles = Directory.GetFiles(templateDir)
            For i As Integer = 0 To allFiles.Length - 1
                Dim nameFile As String = Path.GetFileNameWithoutExtension(allFiles(i))
                ReDim Preserve arrayFilesTemplate(count)
                arrayFilesTemplate(count) = nameFile
                count += 1
            Next

            If IsArray(arrayFilesTemplate) = True Then
                ListBox1.Items.Clear()
                ListBox1.Items.AddRange(arrayFilesTemplate)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolWriteFile = False
        Me.Hide()
    End Sub
End Class