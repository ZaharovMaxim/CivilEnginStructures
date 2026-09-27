Imports System.IO

Public Class FormSelectBeams
    Public boolShow As Boolean = False
    Public boolTrack As Boolean = False
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayPref As String() = {"", "а", "б", "в", "г", "д"}
        CB_ListNumberProlet.DataSource = arrayPref
    End Sub
    Private Sub FormSelectBeams_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim nameAlbum As String = ComboBox1.Text
        If nameAlbum.Trim.Length > 0 Then
            Dim nameFolderBearm As String = ComboBox1.Tag & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
            If Directory.Exists(nameFolderBearm) = True Then
                Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
                'находим файлы xls в директории альбома
                Dim allFilesXML As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
                If allFilesXML.Length > 0 Then
                    'загружаем из файла xls балки
                    Dim arrayModel As String() = Nothing
                    For i As Integer = 0 To allFilesXML.Length - 1
                        Dim boolFindBeams As Boolean = FuncXML.readNamesBeamsFromXML(allFilesXML(i), arrayModel)
                    Next
                    If IsArray(arrayModel) = True Then
                        ComboBox2.DataSource = arrayModel
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_ListNumberProlet.SelectedIndexChanged

    End Sub
    'указать траекторию
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolTrack = True
        Me.Hide()
    End Sub
End Class