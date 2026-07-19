Imports System.IO
Imports System.Windows.Forms
Imports Topomatic.Alg
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Arrangements
Imports Topomatic.FoundationClasses
Imports Topomatic.Sfc

Public Class FormCreateMonolitSitesBeams
    Public boolShowDlg As Boolean = False
    Public ActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public civilStructuresProject As ProjectCivilStructures = Nothing
    Public civilBridgeProject As ProjectBridge = Nothing
    Public dictNamesProjectBridge As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public dictionaryBridge As Dictionary(Of String, StructureElement) = Nothing 'словарь с мостами активного чертежа
    Public countProlet As Integer = 0
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
    End Sub

    'поиск сооружений в выбранном подобъекте
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesArrProject.SelectedIndexChanged
        Dim nameBridgeProject As String = CBox_ListNamesArrProject.Text
        If nameBridgeProject.Trim.Length = 0 Then Exit Sub
        If IsNothing(ActivDocument) = True Then
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()
            Dim surf As Surface = Nothing
            Dim userAlign As Alignment = Nothing
            '// Получение модели проекта
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameArrx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameArrx.Trim Like nameBridgeProject.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                        userSubObjectBearm = child.Model
                        ActivDocument = userSubObjectBearm.Drawing
                        Exit For
                    End If
                End If
            Next
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получаем сооружение
        Dim arrayNameBridge As String() = {""}
        Dim countArrayNameBridge As Integer = 1
        Dim brigeGeneralAxisDictionary As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        Dim boolNameBridge As Boolean = Bridges.getBridgeObject(ActivDocument, brigeGeneralAxisDictionary)
        If brigeGeneralAxisDictionary.Count > 0 Then
            For i As Integer = 0 To brigeGeneralAxisDictionary.Count - 1
                Dim arrayDataBridge As String() = brigeGeneralAxisDictionary.ElementAt(i).Value
                Dim keyParam As String = arrayDataBridge(1)
                Dim userTempBridge As Bridges = Nothing
                Try
                    userTempBridge = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Bridges)(keyParam)
                Catch ex As Newtonsoft.Json.JsonException
                End Try
                If IsNothing(userTempBridge) = False Then
                    Dim nameTempBridge As String = userTempBridge.NameBridge
                    ReDim Preserve arrayNameBridge(countArrayNameBridge)
                    arrayNameBridge(countArrayNameBridge) = nameTempBridge
                    countArrayNameBridge += 1
                End If
            Next
        End If
        CBox_ListNamesBridge.DataSource = arrayNameBridge
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Расчет числа пролетов
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesBridge.SelectedIndexChanged
        Dim nameBridgeProject As String = CBox_ListNamesArrProject.Text
        If nameBridgeProject.Trim.Length = 0 Then Exit Sub
        If IsNothing(ActivDocument) = True Then
            Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
            Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()
            Dim surf As Surface = Nothing
            Dim userAlign As Alignment = Nothing
            '// Получение модели проекта
            Dim childs As IProjectModel() = userProject.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".arrx" Then
                    Dim fileNameArrx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If fileNameArrx.Trim Like nameBridgeProject.Trim Then
                        ApplicationHost.Current.Plugins.Execute("activate", New Object() {child})
                        userSubObjectBearm = child.Model
                        ActivDocument = userSubObjectBearm.Drawing
                        Exit For
                    End If
                End If
            Next
        End If
        Dim brigeGeneralAxisDictionary As Dictionary(Of String, String()) = New Dictionary(Of String, String())
        Dim boolNameBridge As Boolean = Bridges.getBridgeObject(ActivDocument, brigeGeneralAxisDictionary)
        If brigeGeneralAxisDictionary.Count > 0 Then
            For i As Integer = 0 To brigeGeneralAxisDictionary.Count - 1
                Dim arrayDataBridge As String() = brigeGeneralAxisDictionary.ElementAt(i).Value
                Dim keyParam As String = arrayDataBridge(1)
                Dim userTempBridge As Bridges = Nothing
                Try
                    userTempBridge = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Bridges)(keyParam)
                Catch ex As Newtonsoft.Json.JsonException
                End Try
                If IsNothing(userTempBridge) = False Then
                    Dim nameTempBridge As String = userTempBridge.NameBridge
                    If nameTempBridge Like CBox_ListNamesBridge.Text.Trim Then
                        countProlet = userTempBridge.ProletCount
                        If countProlet > 0 Then
                            Dim arraProlet As String() = {}
                            For j As Integer = 0 To countProlet - 1
                                ReDim Preserve arraProlet(j)
                                arraProlet(j) = j + 1
                            Next j
                            CBox_ListProlet.DataSource = arraProlet
                            CBox_ListNamesBridge.Tag = arrayDataBridge(3)
                        End If
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub
    'отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShowDlg = False
        Me.Close()
    End Sub
    'ок
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShowDlg = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор объекта
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim arrayProlet As String() = {}
        For j As Integer = 0 To countProlet - 1
            ReDim Preserve arrayProlet(j)
            arrayProlet(j) = j + 1
        Next j
        CBox_ListProlet.DataSource = arrayProlet
    End Sub

End Class