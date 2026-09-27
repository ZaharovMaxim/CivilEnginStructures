Imports System.IO
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Entities

Public Class FormCreateBeam
    Public CadViewPanel As CadView = Nothing
    Public ActivDocumentPanel As Topomatic.Dwg.Drawing = Nothing
    Public boolCancel As Boolean = False
    Public generalDir As String = ""
    Public dictionaryAlbums As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolCancel = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolCancel = False
        Me.Hide()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles CB_NameAlbums.SelectedValueChanged
        Dim nameAlbum As String = CB_NameAlbums.Text
        If nameAlbum.Trim.Length > 0 Then
            Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
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
                        CB_NameBeams.DataSource = arrayModel
                    End If
                End If
            End If
        End If
    End Sub

End Class