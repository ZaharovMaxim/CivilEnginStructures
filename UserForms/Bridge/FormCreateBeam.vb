Imports System.IO
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Entities

Public Class FormCreateBeam
    Public CadViewPanel As CadView = Nothing
    Public ActivDocumentPanel As Topomatic.Dwg.Drawing = Nothing
    Public boolCancel As Boolean = False
    Public boolOk As Boolean = False
    Public boolElev1 As Boolean = False
    Public boolElev2 As Boolean = False
    Public startPoint As Vector3D = New Vector3D()
    Public endPoint As Vector3D = New Vector3D()
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayPref As String() = {"", "а", "б", "в", "г", "д"}
        ComboBox4.DataSource = arrayPref

    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolCancel = True
        boolOk = False
        boolElev1 = False
        boolElev2 = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolOk = True
        boolCancel = False
        boolElev1 = False
        boolElev2 = False
        Me.Hide()
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отметка начала
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        boolOk = False
        boolCancel = False
        boolElev1 = True
        boolElev2 = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отметка конца
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        boolOk = False
        boolCancel = False
        boolElev1 = False
        boolElev2 = True
        Me.Hide()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub FormCreateBeam_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class