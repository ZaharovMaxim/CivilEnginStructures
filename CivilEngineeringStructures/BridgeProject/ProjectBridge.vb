Imports System.Threading.Tasks
Imports CivilEnginStructures.StructureElement
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports Topomatic.Alg
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Runtime
Imports Drawing = Topomatic.Dwg.Drawing
Public Class ProjectBridge
    Inherits ProjectCivilStructures
    Private _name As String
    Private _arrangementModel As ArrangementModel
    Private _projectSurface As Surface
    Private _egSurface As Surface
    Private _alignment As Alignment
    Private _drawing As Drawing
    Private _dictionaryBridge As Dictionary(Of String, StructureElement)

    Public Sub New()
        _name = "Новое искусственное сооружение"
        _arrangementModel = Nothing
        _projectSurface = Nothing
        _egSurface = Nothing
        _alignment = Nothing
        _drawing = Nothing
        _dictionaryBridge = New Dictionary(Of String, StructureElement)
    End Sub
    ' Свойство Name
    Public Property NameArrangementModel As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Property BridgeModel As ArrangementModel
        Get
            Return _arrangementModel
        End Get
        Set(value As ArrangementModel)
            _arrangementModel = value
        End Set
    End Property

    Public Property ProjectSurface As Surface
        Get
            Return _projectSurface
        End Get
        Set(value As Surface)
            _projectSurface = value
        End Set
    End Property

    Public Property EgSurface As Surface
        Get
            Return _egSurface
        End Get
        Set(value As Surface)
            _egSurface = value
        End Set
    End Property

    Public Property ProjectAlignment As Alignment
        Get
            Return _alignment
        End Get
        Set(value As Alignment)
            _alignment = value
        End Set
    End Property

    Public Property ActivDocument As Drawing
        Get
            Return _drawing
        End Get
        Set(value As Drawing)
            _drawing = value
        End Set
    End Property

    Public Property ListBridges As Dictionary(Of String, StructureElement)
        Get
            Return _dictionaryBridge
        End Get
        Set(value As Dictionary(Of String, StructureElement))
            _dictionaryBridge = value
        End Set
    End Property

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'получить все мосты в проекте arr (массив нужен как список имен для comboBox
    Public Sub getBridges(Optional createNewBridge As Boolean = True)
        If ListBridges.Count > 0 Then
            ListBridges.Clear()
        End If
        'первым идет новое пустое сооружение
        If createNewBridge = True Then
            Dim dataBridge As StructureElement = Bridges.createBridge()
            ListBridges.Add(dataBridge.IdStructure, dataBridge)
        End If
        'далее заполням список существующими сооружениями
        Try
            If IsNothing(BridgeModel) = False Then
                Dim drawModel As Drawing = BridgeModel.Drawing
                If IsNothing(drawModel) = False Then
                    For Each entity As DwgEntity In drawModel.ActiveSpace.Entities
                        ' Пропускаем нерелевантные объекты
                        If Not IsRelevantEntity(entity) Then Continue For
                        Dim xdataElement As StructureElement = New StructureElement
                        ' Читаем XData
                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(entity, xdataElement)
                        If boolReadData = True Then
                            If xdataElement.Name = StructureElement.typeObject.axisBridge Then
                                If xdataElement.IdStructure.Trim.Length > 0 Then
                                    If ListBridges.ContainsKey(xdataElement.IdStructure) = False Then
                                        ListBridges.Add(xdataElement.IdStructure, xdataElement)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'получить все мосты в проекте arr (массив нужен как список имен для comboBox
    Public Shared Function getBridges(ByRef arrProject As ArrangementModel) As List(Of StructureElement)
        Dim result As New List(Of StructureElement)
        'далее заполням список существующими сооружениями
        If IsNothing(arrProject) = False Then
            Try
                Dim drawModel As Drawing = arrProject.Drawing
                If IsNothing(drawModel) = False Then
                    For Each entity As DwgEntity In drawModel.ActiveSpace.Entities
                        Dim xdataElement As StructureElement = New StructureElement
                        Dim boolReadData As Boolean = FuncXRecords.getXRecords(entity, xdataElement)
                        If boolReadData = True Then
                            If xdataElement.Name = StructureElement.typeObject.axisBridge Then
                                If xdataElement.IdStructure.Trim.Length > 0 Then
                                    result.Add(xdataElement)
                                End If
                            End If
                        End If
                    Next
                End If

            Catch ex As System.Exception
            End Try
        End If
        Return result
    End Function

    Private Function IsRelevantEntity(entity As DwgEntity) As Boolean
        IsRelevantEntity = False
        If TypeOf entity Is DwgPolyline Then
            Return True
        End If
    End Function
    'получить все имена мостовых сооружений в виде списка
    Public Function getNamesBridges() As List(Of String)
        Dim result As List(Of String) = New List(Of String)
        If ListBridges.Count > 0 Then
            For i As Integer = 0 To ListBridges.Count - 1
                Dim dataList As StructureElement = ListBridges.ElementAt(i).Value
                If IsNothing(dataList) = False Then
                    Dim keyParamBridge As String = dataList.KeyParameter
                    Dim jsonObj As JObject = JObject.Parse(keyParamBridge)
                    Dim nameBridge As String = jsonObj("NameBridge").Value(Of String)()
                    If nameBridge.Trim.Length = 0 Then
                        nameBridge = "Новое искусственное сооружение"
                    End If
                    result.Add(nameBridge)
                End If
            Next i
        End If
        Return result
    End Function
    'получить все имена мостовых сооружений в виде словаря (ID, Имя сооружения)
    Public Function getDictionaryNamesBridge() As Dictionary(Of String, String)
        Dim result As Dictionary(Of String, String) = New Dictionary(Of String, String)
        If ListBridges.Count > 0 Then
            For i As Integer = 0 To ListBridges.Count - 1
                Dim dataList As StructureElement = ListBridges.ElementAt(i).Value
                If IsNothing(dataList) = False Then
                    Dim keyParamBridge As String = dataList.KeyParameter
                    Dim jsonObj As JObject = JObject.Parse(keyParamBridge)
                    Dim nameBridge As String = jsonObj("NameBridge").Value(Of String)()
                    If nameBridge.Trim.Length = 0 Then
                        nameBridge = "Новое искусственное сооружение"
                    End If
                    If result.ContainsKey(dataList.IdStructure) = False Then
                        result.Add(dataList.IdStructure, nameBridge)
                    End If
                End If
            Next i
        End If
        Return result
    End Function
    'получить характеристики мостового сооружения по его ID
    Public Function getBridgeByID(ByVal idBridge As String) As Bridges
        Dim result As Bridges = Nothing
        If ListBridges.Count > 0 Then
            For i As Integer = 0 To ListBridges.Count - 1
                Dim dataList As StructureElement = ListBridges.ElementAt(i).Value
                If IsNothing(dataList) = False Then
                    If idBridge Like dataList.IdStructure Then
                        result = dataList.getBridge()
                        Exit For
                    End If
                End If
            Next
        End If
        Return result
    End Function
    Public Function getDataBridgeByID(ByVal idBridge As String) As StructureElement
        Dim result As StructureElement = Nothing
        If ListBridges.Count > 0 Then
            For i As Integer = 0 To ListBridges.Count - 1
                Dim dataList As StructureElement = ListBridges.ElementAt(i).Value
                If IsNothing(dataList) = False Then
                    If idBridge Like dataList.IdStructure Then
                        result = dataList
                        Exit For
                    End If
                End If
            Next
        End If
        Return result
    End Function
    Public Function getIdStructureByIndex(ByVal index As Integer) As String
        If index = -1 Then Return Nothing
        If ListBridges.Count > 0 Then
            If index < ListBridges.Count Then
                Return ListBridges.ElementAt(index).Key
            End If
        End If
        Return Nothing
    End Function
    Public Function getBridgeByIndex(ByVal index As Integer) As Bridges
        Dim result As Integer = 0
        If index = -1 Then index = 0
        If ListBridges.Count > 0 Then
            If index < ListBridges.Count Then
                If IsNothing(ListBridges.ElementAt(index).Value) = False Then
                    Return ListBridges.ElementAt(index).Value.getBridge
                End If
            End If
        End If
        Return Nothing
    End Function
    Public Function getStructureElementBridgeByIndex(ByVal index As Integer) As StructureElement
        Dim result As Integer = 0
        If index = -1 Then index = 0
        If ListBridges.Count > 0 Then
            If index < ListBridges.Count Then
                If IsNothing(ListBridges.ElementAt(index).Value) = False Then
                    Return ListBridges.ElementAt(index).Value
                End If
            End If
        End If
        Return Nothing
    End Function
    'получить определенную модель проекта arrx
    Public Shared Function getIArrangementModel(ByVal nameArrangementModel As String, Optional ByVal nameFolder As String = "/Модели/ИССО") As ArrangementModel
        Dim arrModel As ArrangementModel = Nothing
        Try
            Dim modelProject As ModelProject = ApplicationHost.Current.ActiveProject
            Dim modelProjectChilds As IProjectModel() = modelProject.Model.GetChilds()
            For Each modelProjectChild As IProjectModel In modelProjectChilds
                If TypeOf modelProjectChild Is ArrangementModel Then
                    Dim modelProjectUri As Topomatic.FoundationClasses.URI = modelProjectChild.Uri
                    Dim fileNameModelProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                    Dim boolFindFolder As Boolean = True
                    If nameFolder.Trim.Length > 0 Then
                        boolFindFolder = FuncFiles.IsPathContainsFolder(modelProjectUri.AsFilePath, nameFolder)
                    End If
                    If boolFindFolder = True Then
                        If fileNameModelProject Like nameArrangementModel & ".arrx" Then
                            arrModel = modelProjectChild.Model
                            Return arrModel
                        End If
                    End If
                End If
            Next
        Catch ex As System.Exception
        End Try
        Return arrModel
    End Function
    'получить все модели проектов arrx из папки ИССО
    Public Shared Function getIArrangementModels(Optional ByVal nameFolder As String = "/Модели/ИССО") As Dictionary(Of String, ArrangementModel)
        Dim result As New Dictionary(Of String, ArrangementModel)
        Try
            Dim modelProject As ModelProject = ApplicationHost.Current.ActiveProject
            Dim modelProjectChilds As IProjectModel() = modelProject.Model.GetChilds()
            For Each modelProjectChild As IProjectModel In modelProjectChilds
                If TypeOf modelProjectChild Is ArrangementModel Then
                    Dim modelProjectUri As Topomatic.FoundationClasses.URI = modelProjectChild.Uri
                    Dim fileNameModelProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                    Dim boolFindFolder As Boolean = True
                    If nameFolder.Trim.Length > 0 Then
                        boolFindFolder = FuncFiles.IsPathContainsFolder(modelProjectUri.AsFilePath, nameFolder)
                    End If
                    result.Add(modelProjectUri.LastPathComponent, modelProjectChild.Model)
                End If
            Next
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'раскладка балок мостового сооружения
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Sub PlacementBeams(ByVal arrayLastAxisBeams As String(,), ByVal arrayMiddleAxisBeams As String(,), ByVal dictionaryBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByRef dictionaryPillars As Dictionary(Of Integer, List(Of StructureElement)), ByRef dictionaryObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByRef alignStructure As DwgPolyline, ByVal putchAlbumBeams As String, ByVal templateXML As String)
        If dictionaryBeams.Count = 0 Then
            MsgBox("Словарь с марками балок пуст. Сооружение не построено!!")
            Exit Sub
        End If
        '1. Оформление
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim deltaAxisPillar As Double = 2 'величина выпуска осей опоры за границы сооружения
        Dim categoryTables As String = "Искусственные сооружения"
        Dim nameTableBridge As String = "Мостовое сооружение"
        Dim funcBridges As Bridges = New Bridges
        'ось балки
        Dim drawingPlacementBeams As Drawing = ActivDocument
        If IsNothing(drawingPlacementBeams) = True Then
            MsgBox("Активный проект для раскладки балок не найден")
            Exit Sub
        End If
        Dim styleAxisBeam As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleAxisBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось балки")
        'балка
        Dim styleModelBeam As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleModelBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Балка (модель)")
        'Верх ребра балки
        Dim styleTopBeam As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleTopBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Верх ребра плиты балки")
        'Низ ребра балки
        Dim styleBottomBeam As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleBottomBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Низ ребра балки")
        'Ось опоры
        Dim styleAxisPillar As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleAxisPillar.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опоры")
        'Ось опирания балок
        Dim styleAxisBeamsPillar As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleAxisBeamsPillar.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опирания балок")
        'главная ось путепровода
        Dim styleAxisBridge As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleAxisBridge.setObjectStyle(templateXML, categoryTables, nameTableBridge, ProjectCivilStructuresStyle.typeEntity.Полилиния, "Главная ось сооружения")
        'граница путепровода
        Dim styleBoundBridge As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingPlacementBeams)
        styleBoundBridge.setObjectStyle(templateXML, categoryTables, nameTableBridge, ProjectCivilStructuresStyle.typeEntity.Полилиния, "Граница сооружения")
        '================================================================================================
        If IsNothing(alignStructure) = True Then
            MsgBox("Ось сооружения не найдена.")
            Exit Sub
        End If
        Dim dataAlign As StructureElement = New StructureElement
        Dim boolFindData As Boolean = FuncXRecords.getXRecords(alignStructure, dataAlign)
        If boolFindData = False Then
            MsgBox("Оси сооружения не присвоены данные.")
            Exit Sub
        End If
        Dim userBridge As Bridges = dataAlign.getBridge()
        Dim idBridge As String = dataAlign.IdStructure
        '=================================================================================================
        'читаем свойства сооружения
        If userBridge.LeftStructureWidth = 0 And userBridge.RightStructureWidth = 0 Then
            MsgBox("Ширина сооружения не задана. Сооружение не построено!!!!")
            Exit Sub
        End If
        'имя трассы
        Dim align As Alignment = ProjectAlignment
        If IsNothing(align) = True Then
            MsgBox("Трасса не найдена")
            Exit Sub
        End If
        Dim surf As Surface = ProjectSurface
        If IsNothing(surf) = True Then
            MsgBox("Проектная поверхность не найдена")
            Exit Sub
        End If
        'граница габарита моста слева
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, -1 * userBridge.LeftStructureWidth + userBridge.TransverseOffset)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)
        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPlineDirect) = True Then
            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPlineDirect)
        End If
        'граница габарита моста справа
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, -1 * userBridge.RightStructureWidth + userBridge.TransverseOffset, True)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)
        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPlineReverse) = True Then
            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPlineReverse)
        End If

        Dim axisCentrePline As DwgPolyline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, userBridge.TransverseOffset)
        Dim axisCentrePline3D As IPolyline3D = New Polyline3D()
        axisCentrePline.GetPolyline(axisCentrePline3D)
        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisCentrePline) = True Then
            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisCentrePline)
        End If
        Try
            '==================================================================================================
            drawingPlacementBeams.BeginUpdate()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim StartPositionBridge As Double = 0
            Dim endPositionBridge As Double = 0
            Dim numberProlet As Integer = userBridge.ProletCount
            Dim numberRows As Integer = 0
            If IsArray(arrayLastAxisBeams) = True Then
                numberRows = arrayLastAxisBeams.GetUpperBound(1) + 1
            End If
            If IsArray(arrayMiddleAxisBeams) = True Then
                numberRows = numberRows + arrayMiddleAxisBeams.GetUpperBound(1)
            End If
            'заполняем начальный словарь с балками
            Dim dictBeams As Dictionary(Of Integer, DwgLine()) = New Dictionary(Of Integer, DwgLine())
            For i As Integer = 1 To numberProlet
                Dim arrayRowBeams As DwgLine() = Nothing
                ReDim arrayRowBeams(numberRows)
                dictBeams.Add(i, arrayRowBeams)
            Next i
            'словарь содержит элементы 2 крайних балок первой и последней опоры и служит для рисования габарита сооружения
            Dim dictionaryLastElementBeam As Dictionary(Of Integer, DwgLine()) = New Dictionary(Of Integer, DwgLine())
            'массив с пролетами общий
            Dim arrayPr0 As String(,) = Nothing
            Dim countArrayPr0 As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'расчетный блок для крайних рядов
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '=============================================================================================================
            'раскладка фиксированными балками
            '=============================================================================================================
            Dim numberDefinedAxisPillars As Integer = 0
            If IsNothing(dictionaryPillars) = True Then
                MsgBox("Массив с осями опор сооружения не найден. Сооружение не построено!!!")
                Exit Sub
            End If
            If userBridge.TypeBridge = 0 Then
                Dim startSection As DwgLine = Nothing 'находим определяющую опору,от которой будем производить построения
                'определяющая опора
                If dictionaryPillars.Count < 2 Then
                    MsgBox("Число опор сооружения меньше двух. Сооружение не построено!!!")
                    Exit Sub
                End If
                'перебираем массив с пролетами в прямом направлении
                Dim arrayPr1 As String(,) = Nothing
                Dim countArrayPr1 As Integer = 0
                Dim boolWriteArray As Boolean = False
                For i As Integer = 0 To dictionaryPillars.Count - 1
                    Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key
                    Dim listDataPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value
                    If IsNothing(listDataPillar) = True Then Continue For
                    If listDataPillar.Count > 2 Then
                        'получаем доступ к опоре
                        Dim userDataPillar As StructureElement = listDataPillar.Item(1)
                        If IsNothing(userDataPillar) = False Then
                            Dim userAxisPillar As Pillar = userDataPillar.getPillar()
                            If IsNothing(userAxisPillar) = False Then
                                Dim boolCheck As Boolean = userAxisPillar.Defining
                                If boolCheck = True Then
                                    boolWriteArray = True
                                    numberDefinedAxisPillars = userAxisPillar.Number
                                    startSection = userDataPillar.DWGEntity
                                    'делаем смещение определяющей опоры
                                    If userBridge.HorizontalOffset <> 0 Then
                                        Dim deltaMoveBridge As Double = Math.Round((userBridge.startPlacementPosition - userBridge.HorizontalOffset), 3)
                                        If deltaMoveBridge <> 0 Then
                                            'делаем смещение определяющей опоры
                                            Dim boolMovePillar As Boolean = Pillar.moveAxisPillar(axisCentrePline3D, startSection, deltaMoveBridge)
                                            If boolMovePillar = False Then
                                                MsgBox("Не удалось переместить определяющую опору №" & userAxisPillar.Number & ". Возможно она за пределами вытранной трассы!!! Сооружение не построено.")
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                                If boolWriteArray = True Then
                                    ReDim Preserve arrayPr1(3, countArrayPr1)
                                    arrayPr1(0, countArrayPr1) = numberPillar 'номер опоры
                                    arrayPr1(1, countArrayPr1) = Val(userAxisPillar.Clearence) 'левый зазор
                                    arrayPr1(2, countArrayPr1) = Val(userAxisPillar.RightClearence) 'правый зазор
                                    arrayPr1(3, countArrayPr1) = Val(userAxisPillar.SiteMonolit) 'участок омоличивания балки
                                    countArrayPr1 += 1
                                End If
                            End If
                        End If
                    End If
                Next i
                'массив с пролетами обратный
                boolWriteArray = False
                Dim arrayPr2 As String(,) = Nothing
                Dim countArrayPr2 As Integer = 0
                For i As Integer = dictionaryPillars.Count - 1 To 0 Step -1
                    Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key
                    Dim listDataPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value
                    If IsNothing(listDataPillar) = True Then Continue For
                    If listDataPillar.Count > 2 Then
                        Dim userDataPillar As StructureElement = listDataPillar.Item(1)
                        If IsNothing(userDataPillar) = False Then
                            Dim userAxisPillar As Pillar = userDataPillar.getPillar()
                            If IsNothing(userAxisPillar) = False Then
                                Dim boolCheck As Boolean = userAxisPillar.Defining
                                If boolCheck = True Then
                                    boolWriteArray = True
                                    numberDefinedAxisPillars = userAxisPillar.Number
                                    startSection = userDataPillar.DWGEntity
                                    'делаем смещение определяющей опоры
                                    If userBridge.HorizontalOffset <> 0 Then
                                        Dim deltaMoveBridge As Double = Math.Round((userBridge.startPlacementPosition - userBridge.HorizontalOffset), 3)
                                        If deltaMoveBridge <> 0 Then
                                            'делаем смещение определяющей опоры
                                            Dim boolMovePillar As Boolean = Pillar.moveAxisPillar(axisCentrePline3D, startSection, deltaMoveBridge)
                                        End If
                                    End If
                                End If
                                If boolWriteArray = True Then
                                    ReDim Preserve arrayPr2(3, countArrayPr2)
                                    arrayPr2(0, countArrayPr2) = numberPillar 'номер опоры
                                    arrayPr2(1, countArrayPr2) = Val(userAxisPillar.Clearence) 'левый зазор
                                    arrayPr2(2, countArrayPr2) = Val(userAxisPillar.RightClearence) 'правый зазор
                                    arrayPr2(3, countArrayPr2) = Val(userAxisPillar.SiteMonolit) 'участок омоличивания балки
                                    countArrayPr2 += 1
                                End If
                            End If
                        End If
                    End If
                Next i
                'если определяющая опора не выбрана, назначаем перевую опору в качестве определяющей
                If numberDefinedAxisPillars = 0 Then
                    numberDefinedAxisPillars = 1
                    Dim listDataPillar As List(Of StructureElement) = dictionaryPillars.Item(numberDefinedAxisPillars)
                    If listDataPillar.Count > 2 Then
                        Dim userDataPillar As StructureElement = listDataPillar.Item(1)
                        If IsNothing(userDataPillar) = False Then
                            Dim userAxisPillar As Pillar = userDataPillar.getPillar()
                            If IsNothing(userAxisPillar) Then
                                userAxisPillar.Defining = True
                            End If
                            listDataPillar.Item(1) = userDataPillar
                            Dim axisLine As DwgLine = userDataPillar.DWGEntity
                            If IsNothing(axisLine) = True Then
                                axisLine = New DwgLine
                            End If
                            Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar)
                            userDataPillar.KeyParameter = keyParam
                            Dim boolRec As Boolean = FuncXRecords.setXRecords(axisLine, StructureElement.tableXRecords.PROJECT_STRUCTURES, userDataPillar)
                            startSection = axisLine
                        End If
                        listDataPillar.Item(1) = userDataPillar
                    End If
                End If
                '=============================================================================================
                'проверка исходных данных
                If IsNothing(startSection) = True Then
                    MsgBox("Не выбрана определяющая опора начала раскладки. Сооружение не построено!!!")
                    Exit Sub
                End If
                If startSection.Length = 0 Then
                    MsgBox("Не выбрана определяющая опора начала раскладки. Сооружение не построено!!!")
                    Exit Sub
                End If
                'если в массиве только один элемент - это крйняя последняя опора очищаем массв
                If IsNothing(arrayPr1) = False Then
                    If arrayPr1.GetUpperBound(1) = 0 Then
                        Erase arrayPr1
                    End If
                End If
                If IsNothing(arrayPr2) = False Then
                    If arrayPr2.GetUpperBound(1) = 0 Then
                        Erase arrayPr2
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '1.Расставляем крайние ряды
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If IsArray(arrayPr1) = True Then
                    For i As Integer = 0 To arrayLastAxisBeams.GetUpperBound(1)
                        'первый и второй ряды считаем по зазорам, остальные так кладем
                        Dim numberRowBeam As Integer = CInt(arrayLastAxisBeams(0, i)) 'номер ряда
                        Dim offsetPlaneRowAxis As Double = Val(arrayLastAxisBeams(1, i)) 'горизонтальное смещение
                        Dim offsetElevRowAxis As Double = Val(arrayLastAxisBeams(2, i)) 'вертикальное смещение
                        Dim axisPline As DwgPolyline = Nothing
                        Dim axisPline3D As IPolyline3D = New Polyline3D()
                        Dim hgTraskObj As String = arrayLastAxisBeams(3, i)
                        If IsNothing(hgTraskObj) = False Then
                            Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                            If IsNothing(dataPlacementBeams) = False Then
                                If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                    axisPline = dataPlacementBeams.DWGEntity
                                    If axisPline.Length > 0 Then
                                        Dim stPoint As Vector2D = axisPline.Item(0).Vertex
                                        Dim enPoint As Vector2D = axisPline.Item(axisPline.Count - 1).Vertex
                                        Dim startPK As Double = 0
                                        Dim off As Double = 0
                                        Dim endPk As Double = 0
                                        Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                        Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                        If boolStartPk = True And boolEndPk = True Then
                                            If startPK > endPk Then
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, True)
                                            Else
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, False)
                                            End If
                                        End If
                                        axisPline.GetPolyline(axisPline3D)
                                    Else
                                        axisPline = Nothing
                                    End If
                                End If
                            End If
                        End If
                        If IsNothing(axisPline) = True Then
                            'делаем смещение оси трассы
                            axisPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis)
                            axisPline.GetPolyline(axisPline3D)
                            'удаляем вспомогательную линию
                            If IsNothing(axisPline) = False Then
                                drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                            End If
                        End If
                        'переменная для запоминания предыдущей балки
                        Dim prevLineShortBeam As DwgLine = Nothing
                        Dim prevUserBeam As BeamI = Nothing
                        'строим балки по крайним рядам,
                        For j As Integer = 0 To arrayPr1.GetUpperBound(1) - 1
                            'номер пролета
                            numberProlet = Val(arrayPr1(0, j))
                            'берем зазор
                            Dim zazor As Double = 0
                            If i = 0 Then
                                zazor = Val(arrayPr1(1, j))
                            ElseIf i = 1 Then
                                zazor = Val(arrayPr1(2, j))
                                If zazor = 0 Then
                                    zazor = Val(arrayPr1(1, j))
                                End If
                            End If
                            Dim monolitBeam1 As Double = Val(arrayPr1(3, j))
                            Dim monolitBeam2 As Double = Val(arrayPr1(3, j + 1))
                            'читаем марку балки
                            Dim nameAlbumBeam As String = ""
                            Dim modelBeam As String = ""
                            Dim userBeam As BeamI = Nothing
                            Dim dataBeam As StructureElement = Nothing
                            If dictionaryBeams.ContainsKey(numberProlet) = True Then
                                Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                If IsNothing(listBeamsProlet) = False Then
                                    If listBeamsProlet.Count > 0 Then
                                        dataBeam = listBeamsProlet.Item(numberRowBeam)
                                        If IsNothing(dataBeam) = False Then
                                            userBeam = dataBeam.getBeamI()
                                        End If
                                    End If
                                    If IsNothing(userBeam) = True Then
                                        MsgBox("Марка балки в крайнем ряду №" & numberRowBeam & ", пролет № " & numberProlet & ", не найдена. Сооружение не построено!!!")
                                        Exit Sub
                                    End If
                                    nameAlbumBeam = userBeam.nameAlbum
                                    modelBeam = userBeam.model
                                    'записываем высоту балки над поверхностью
                                    userBeam.offsetSurface = offsetElevRowAxis
                                    userBeam.clearence = Math.Round(zazor, 3)
                                    userBeam.startLenghtMonolith = Math.Round(monolitBeam1, 3)
                                    userBeam.endLenghtMonolith = Math.Round(monolitBeam2, 3)
                                    'номер ряда для записи в массив
                                    '=============================================================================================
                                    If j = 0 Then
                                        '==========================================================================================
                                        'первая балка
                                        Dim startAlignPoint As Vector2D = New Vector2D(0, 0) 'начальная точка раскладки балок
                                        Dim startAlignPoint3d As Vector3D = New Vector3D(0, 0, 0)
                                        Dim pkStartPoint As Double = 0 'пикет начала раскладки
                                        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, startSection.StartPoint.Pos, startSection.EndPoint.Pos)
                                        If pointIntersectCollection.Count = 0 Then
                                            'удлинним второй аргумент (отрезок)
                                            Dim newStartPoint As Vector3D = startSection.StartPoint
                                            Dim newEndPoint As Vector3D = startSection.EndPoint
                                            Dim boolExtendLine As Boolean = MathFunction.FuncExtendPos(newStartPoint, newEndPoint, startSection.Length, startSection.Length)
                                            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPline3D, newStartPoint.Pos, newEndPoint.Pos)
                                        End If
                                        If pointIntersectCollection.Count > 0 Then
                                            startAlignPoint = pointIntersectCollection.ElementAt(0)
                                            'высота начальной точки раскладки по низу балки
                                            Dim elevST As Double = 0
                                            Try
                                                elevST = surf.GetElevation(startAlignPoint)
                                                elevST = elevST - userBeam.height - userBeam.offsetSurface
                                                startAlignPoint3d = New Vector3D(startAlignPoint, elevST)
                                            Catch ex As System.NullReferenceException
                                                MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                                Exit Sub
                                            Catch ex As System.InvalidOperationException
                                                MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                                Exit Sub
                                            End Try
                                            Dim off As Double = 0
                                            Dim distStartPointPr As Double = 0
                                            If IsArray(arrayPr2) = True Then
                                                'это раскладка из середины
                                                distStartPointPr = userBeam.a + zazor / 2
                                                startAlignPoint3d = CalculationBeams.correctionLenghtBeamToElevation(axisPline, startAlignPoint3d, distStartPointPr, surf, userBeam.height, userBeam.offsetSurface)
                                                If startAlignPoint3d.X = -1 And startAlignPoint3d.Y = -1 And startAlignPoint3d.Z = -1 Then
                                                    MsgBox("Ошибка в коррекции балки по высоте. Сооркжение не построено!!!")
                                                    Exit Sub
                                                End If
                                                startAlignPoint = startAlignPoint3d.Pos
                                            End If
                                            'вычисляем пикет начала раскладки балки
                                            Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startAlignPoint, pkStartPoint, off)
                                            'вычисляем второй конец раскладки балки
                                            Dim distEndPointPr As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPline, startAlignPoint3d, distEndPointPr, surf, userBeam.height, userBeam.offsetSurface)
                                            'создаем ось балки или корректируем старую
                                            Dim LineShortBeam As DwgLine = New DwgLine()
                                            If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                                                Dim listBridgeBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                                                If listBridgeBeams.Count > 0 Then
                                                    For k As Integer = 0 To listBridgeBeams.Count - 1
                                                        Dim tempData As StructureElement = listBridgeBeams.Item(k)
                                                        If IsNothing(tempData) = False Then
                                                            Dim tempBeam As BeamI = tempData.getBeamI
                                                            If tempBeam.numberProlet = numberProlet And tempBeam.numberRow = numberRowBeam Then
                                                                LineShortBeam = tempData.DWGEntity
                                                                listBridgeBeams.Item(k) = Nothing
                                                            End If
                                                            dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam) = listBridgeBeams
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                            LineShortBeam.StartPoint = startAlignPoint3d
                                            LineShortBeam.EndPoint = endAlignPoint3d
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(LineShortBeam) = False Then
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(LineShortBeam)
                                                dataBeam.DWGEntity = LineShortBeam
                                            End If
                                            'делаем коррекцию балки в плане
                                            Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPline3D, theoryShortLineBeam)
                                            '==============================================================================================================
                                            'оформляем балку
                                            userBeam.numberProlet = numberProlet
                                            userBeam.numberRow = numberRowBeam
                                            userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                            userBeam.model = modelBeam
                                            userBeam.nameAlbum = nameAlbumBeam
                                            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            Dim listProjectBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                            Dim dataBeamStructure As StructureElement = listProjectBeams.Item(numberRowBeam)
                                            dataBeamStructure.DWGEntity = LineShortBeam
                                            'опускаем балку на нужную высоту
                                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                            'возвращаем значение в словарь
                                            userBeam._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                            userBeam._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                            strGSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            dataBeamStructure.KeyParameter = strGSON
                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                            listProjectBeams.Item(numberRowBeam) = dataBeamStructure
                                            dictionaryBeams.Item(numberProlet) = listProjectBeams
                                            prevLineShortBeam = LineShortBeam
                                            prevUserBeam = userBeam
                                        Else
                                            MsgBox("Пересечение оси раскладки балок с осью опоры сооружения не найдено. Сооружение не построено!!!")
                                            Exit Sub
                                        End If
                                    Else
                                        'остальные балки в ряду
                                        Dim startAlignPoint3d As Vector3D = prevLineShortBeam.EndPoint
                                        Dim pkStartPoint As Double = 0
                                        Dim off As Double = 0
                                        Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startAlignPoint3d.Pos, pkStartPoint, off)
                                        If boolStartPoint = True Then
                                            pkStartPoint = pkStartPoint + prevUserBeam.b + zazor + userBeam.a
                                            Dim startAlignPoint As Vector2D = axisPline3D.StaOffsetToPos(pkStartPoint, 0)
                                            Dim elevST As Double = 0
                                            Try
                                                elevST = surf.GetElevation(startAlignPoint)
                                                elevST = elevST - userBeam.height - userBeam.offsetSurface
                                                startAlignPoint3d = New Vector3D(startAlignPoint, elevST)
                                            Catch ex As System.Exception
                                                MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооружение не построено!!!")
                                                Exit Sub
                                            End Try
                                            Dim distEndPointPr As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPline, startAlignPoint3d, distEndPointPr, surf, userBeam.height, userBeam.offsetSurface)

                                            Dim LineShortBeam As DwgLine = New DwgLine()
                                            If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                                                Dim listBridgeBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                                                If listBridgeBeams.Count > 0 Then
                                                    For k As Integer = 0 To listBridgeBeams.Count - 1
                                                        Dim tempData As StructureElement = listBridgeBeams.Item(k)
                                                        If IsNothing(tempData) = False Then
                                                            Dim tempBeam As BeamI = tempData.getBeamI
                                                            If tempBeam.numberProlet = numberProlet And tempBeam.numberRow = numberRowBeam Then
                                                                LineShortBeam = tempData.DWGEntity
                                                                listBridgeBeams.Item(k) = Nothing
                                                            End If
                                                            dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam) = listBridgeBeams
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(LineShortBeam) = False Then
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(LineShortBeam)
                                                dataBeam.DWGEntity = LineShortBeam
                                            End If
                                            LineShortBeam.StartPoint = startAlignPoint3d
                                            LineShortBeam.EndPoint = endAlignPoint3d
                                            '==============================================================================================================
                                            'оформляем балку
                                            userBeam.numberProlet = numberProlet
                                            userBeam.numberRow = numberRowBeam
                                            userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                            userBeam.model = modelBeam
                                            userBeam.nameAlbum = nameAlbumBeam
                                            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            Dim listProjectBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                            Dim dataBeamStructure As StructureElement = listProjectBeams.Item(numberRowBeam)
                                            dataBeamStructure.DWGEntity = LineShortBeam
                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                            For k As Integer = 0 To 2
                                                Dim userListClearence As List(Of Double) = New List(Of Double)
                                                Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, LineShortBeam, userListClearence)
                                                Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                                If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                                    Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(LineShortBeam, deltaTrimBeam)
                                                End If
                                                'корректируем балку по высоте
                                                Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                                'делаем коррекцию балки в плане
                                                Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                                Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPline3D, theoryShortLineBeam)
                                            Next k
                                            Dim userListClearenceControl As List(Of Double) = New List(Of Double)
                                            Dim minZazorControl As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, LineShortBeam, userListClearenceControl)
                                            'возвращаем значение в словарь
                                            userBeam._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                            userBeam._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                            strGSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            dataBeamStructure.KeyParameter = strGSON
                                            boolRecData = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                            listProjectBeams.Item(numberRowBeam) = dataBeamStructure
                                            dictionaryBeams.Item(numberProlet) = listProjectBeams
                                            prevLineShortBeam = LineShortBeam
                                        Else
                                            MsgBox("Не удалось определить начальный пикет раскладки балки. Сооружение не построено!!!")
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        Next j
                        If IsNothing(axisPline) = False Then
                            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                        End If
                    Next i
                End If
                'строим балки в обратном направлении
                If IsArray(arrayPr2) = True Then
                    For i As Integer = 0 To arrayLastAxisBeams.GetUpperBound(1)
                        'первый и второй ряды считаем по зазорам, остальные так кладем
                        Dim numberRowBeam As Integer = CInt(arrayLastAxisBeams(0, i))
                        Dim offsetPlaneRowAxis As Double = Val(arrayLastAxisBeams(1, i)) * -1
                        Dim offsetElevRowAxis As Double = Val(arrayLastAxisBeams(2, i))
                        Dim axisPline3D As IPolyline3D = New Polyline3D()
                        Dim axisPline As DwgPolyline = Nothing
                        Dim hgTraskObj As String = arrayLastAxisBeams(3, i)
                        If IsNothing(hgTraskObj) = False Then
                            Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                            If IsNothing(dataPlacementBeams) = False Then
                                If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                    axisPline = dataPlacementBeams.DWGEntity
                                    If axisPline.Length > 0 Then
                                        Dim stPoint As Vector2D = axisPline.Item(0).Vertex
                                        Dim enPoint As Vector2D = axisPline.Item(axisPline.Count - 1).Vertex
                                        Dim startPK As Double = 0
                                        Dim off As Double = 0
                                        Dim endPk As Double = 0
                                        Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                        Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                        If boolStartPk = True And boolEndPk = True Then
                                            If startPK > endPk Then
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, False)
                                            Else
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, True)
                                            End If
                                        End If
                                        axisPline.GetPolyline(axisPline3D)
                                    Else
                                        axisPline = Nothing
                                    End If
                                End If
                            End If
                        End If
                        If IsNothing(axisPline) = True Then
                            axisPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis, True)
                            axisPline.GetPolyline(axisPline3D)
                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPline) = True Then
                                drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                            End If
                        End If
                        'номер ряда для записи в массив созданных балок (0-первый ряд,numberRow-последний ряд)
                        Dim numberWriteRowBeam As Integer = 0
                        If i <> 0 Then
                            numberWriteRowBeam = numberRows
                        End If
                        'переменная для запоминания предыдущей балки
                        Dim prevLineShortBeam As DwgLine = Nothing
                        Dim prevUserBeam As BeamI = Nothing
                        'строим балки по крайним рядам,
                        For j As Integer = 0 To arrayPr2.GetUpperBound(1) - 1
                            'пролеты
                            numberProlet = Val(arrayPr2(0, j)) - 1
                            Dim zazor As Double = 0
                            If i = 0 Then
                                zazor = Val(arrayPr2(1, j))
                            ElseIf i = 1 Then
                                zazor = Val(arrayPr2(2, j))
                                If zazor = 0 Then
                                    zazor = Val(arrayPr2(1, j))
                                End If
                            End If
                            Dim monolitBeam1 As Double = Val(arrayPr2(3, j))
                            Dim monolitBeam2 As Double = Val(arrayPr2(3, j + 1))
                            'читаем марку балки
                            Dim nameAlbumBeam As String = ""
                            Dim modelBeam As String = ""
                            Dim userBeam As BeamI = Nothing
                            Dim dataBeam As StructureElement = Nothing
                            If dictionaryBeams.ContainsKey(numberProlet) = True Then
                                Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                If IsNothing(listBeamsProlet) = False Then
                                    If listBeamsProlet.Count > 0 Then
                                        dataBeam = listBeamsProlet.Item(numberRowBeam)
                                        If IsNothing(dataBeam) = False Then
                                            userBeam = dataBeam.getBeamI()
                                        End If
                                    End If
                                    If IsNothing(userBeam) = True Then
                                        MsgBox("Марка балки в крайнем ряду №" & numberRowBeam & ", пролет № " & numberProlet & ", не найдена. Сооружение не построено!!!")
                                        Exit Sub
                                    End If
                                    nameAlbumBeam = userBeam.nameAlbum
                                    modelBeam = userBeam.model
                                    userBeam.offsetSurface = offsetElevRowAxis
                                    userBeam.clearence = Math.Round(zazor, 3)
                                    userBeam.startLenghtMonolith = Math.Round(monolitBeam2, 3)
                                    userBeam.endLenghtMonolith = Math.Round(monolitBeam1, 3)
                                    'получаем ось опоры и ось опирания балок
                                    Dim listPillars As List(Of StructureElement) = New List(Of StructureElement)
                                    If dictionaryPillars.ContainsKey(numberProlet) = True Then
                                        listPillars = dictionaryPillars.Item(numberProlet)
                                    End If
                                    '======================================================================================
                                    If IsArray(arrayPr1) = False And j = 0 Then 'прямой массив пустой, раскладка с конца сооружения
                                        'раскладка с конца моста
                                        Dim startAlignPoint As Vector2D = New Vector2D(0, 0) 'начальная точка раскладки балок
                                        Dim startAlignPoint3d As Vector3D = New Vector3D(0, 0, 0)
                                        Dim pkStartPoint As Double = 0 'пикет начала раскладки
                                        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, startSection.StartPoint.Pos, startSection.EndPoint.Pos)
                                        If pointIntersectCollection.Count > 0 Then
                                            startAlignPoint = pointIntersectCollection.ElementAt(0)
                                            'высота начальной точки раскладки по низу балки
                                            Dim elevST As Double = 0
                                            Try
                                                elevST = surf.GetElevation(startAlignPoint)
                                                elevST = elevST - userBeam.height - userBeam.offsetSurface
                                                startAlignPoint3d = New Vector3D(startAlignPoint, elevST)
                                            Catch ex As System.Exception
                                                MsgBox("Ошибка в определении высоты начальной точки опирания балки.")
                                                Exit Sub
                                            End Try
                                            Dim off As Double = 0
                                            Dim distStartPointPr As Double = 0
                                            'вычисляем второй конец раскладки балки
                                            Dim distEndPointPr As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPline, startAlignPoint3d, distEndPointPr, surf, userBeam.height, userBeam.offsetSurface)
                                            'новая ось балки
                                            Dim LineShortBeam As DwgLine = New DwgLine()
                                            If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                                                Dim listBridgeBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                                                If listBridgeBeams.Count > 0 Then
                                                    For k As Integer = 0 To listBridgeBeams.Count - 1
                                                        Dim tempData As StructureElement = listBridgeBeams.Item(k)
                                                        If IsNothing(tempData) = False Then
                                                            Dim tempBeam As BeamI = tempData.getBeamI
                                                            If tempBeam.numberProlet = numberProlet And tempBeam.numberRow = numberRowBeam Then
                                                                LineShortBeam = tempData.DWGEntity
                                                                listBridgeBeams.Item(k) = Nothing
                                                            End If
                                                            dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam) = listBridgeBeams
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(LineShortBeam) = False Then
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(LineShortBeam)
                                                dataBeam.DWGEntity = LineShortBeam
                                            End If
                                            LineShortBeam.StartPoint = startAlignPoint3d
                                            LineShortBeam.EndPoint = endAlignPoint3d
                                            'оформляем балку
                                            userBeam.numberProlet = numberProlet
                                            userBeam.numberRow = numberRowBeam
                                            userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                            userBeam.model = modelBeam
                                            userBeam.nameAlbum = nameAlbumBeam
                                            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            Dim listProjectBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                            Dim dataBeamStructure As StructureElement = listProjectBeams.Item(numberRowBeam)
                                            dataBeamStructure.DWGEntity = LineShortBeam
                                            'корректируем балку по высоте
                                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                            'делаем коррекцию балки в плане
                                            Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPline3D, theoryShortLineBeam)
                                            'возвращаем значение в словарь
                                            userBeam._elementBridgePoint.StartAxisPoint = LineShortBeam.EndPoint
                                            userBeam._elementBridgePoint.EndAxisPoint = LineShortBeam.StartPoint
                                            strGSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            dataBeamStructure.KeyParameter = strGSON
                                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                            listProjectBeams.Item(numberRowBeam) = dataBeamStructure
                                            dictionaryBeams.Item(numberProlet) = listProjectBeams
                                            prevLineShortBeam = LineShortBeam
                                        Else
                                            MsgBox("Не удалось определить начальный пикет раскладки балки. Сооружение не построено!!!")
                                            Exit Sub
                                        End If
                                    Else
                                        'раскладка из середины
                                        If j = 0 Then
                                            'первая балка раскладка из середины
                                            If i = 0 Then
                                                If dictionaryBeams.ContainsKey(numberProlet + 1) = True Then
                                                    Dim listFirstBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet + 1)
                                                    Dim dataFirstBeam As StructureElement = listFirstBeams.Item(numberRowBeam)
                                                    prevUserBeam = dataFirstBeam.getBeamI()
                                                    prevLineShortBeam = dataFirstBeam.DWGEntity
                                                    Dim tempStartPoint As Vector3D = prevLineShortBeam.StartPoint
                                                    prevLineShortBeam.StartPoint = prevLineShortBeam.EndPoint
                                                    prevLineShortBeam.EndPoint = tempStartPoint
                                                Else
                                                    MsgBox("Балка с ловаре не найдена.")
                                                    Exit Sub
                                                End If
                                            Else
                                                If dictionaryBeams.ContainsKey(numberProlet + 1) = True Then
                                                    Dim listFirstBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet + 1)
                                                    Dim dataFirstBeam As StructureElement = listFirstBeams.Item(numberRowBeam)
                                                    prevUserBeam = dataFirstBeam.getBeamI()
                                                    prevLineShortBeam = dataFirstBeam.DWGEntity
                                                    Dim tempStartPoint As Vector3D = prevLineShortBeam.StartPoint
                                                    prevLineShortBeam.StartPoint = prevLineShortBeam.EndPoint
                                                    prevLineShortBeam.EndPoint = tempStartPoint
                                                Else
                                                    MsgBox("Балка с ловаре не найдена.")
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                        Dim startAlignPoint3d As Vector3D = prevLineShortBeam.EndPoint
                                        Dim pkStartPoint As Double = 0
                                        Dim off As Double = 0
                                        Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startAlignPoint3d.Pos, pkStartPoint, off)
                                        pkStartPoint = pkStartPoint + prevUserBeam.b + zazor + userBeam.a
                                        Dim startAlignPoint As Vector2D = axisPline3D.StaOffsetToPos(pkStartPoint, 0)
                                        Dim elevST As Double = 0
                                        Try
                                            elevST = surf.GetElevation(startAlignPoint)
                                            elevST = elevST - userBeam.height - userBeam.offsetSurface
                                            startAlignPoint3d = New Vector3D(startAlignPoint, elevST)
                                        Catch ex As System.Exception
                                            MsgBox("Ошибка в определении высоты начальной точки опирания балки")
                                            Exit Sub
                                        End Try
                                        Dim distEndPointPr As Double = userBeam.lenght - userBeam.a - userBeam.b
                                        Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPline, startAlignPoint3d, distEndPointPr, surf, userBeam.height, userBeam.offsetSurface)
                                        Dim LineShortBeam As DwgLine = New DwgLine()
                                        If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                                            Dim listBridgeBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                                            If listBridgeBeams.Count > 0 Then
                                                For k As Integer = 0 To listBridgeBeams.Count - 1
                                                    Dim tempData As StructureElement = listBridgeBeams.Item(k)
                                                    If IsNothing(tempData) = False Then
                                                        Dim tempBeam As BeamI = tempData.getBeamI
                                                        If tempBeam.numberProlet = numberProlet And tempBeam.numberRow = numberRowBeam Then
                                                            LineShortBeam = tempData.DWGEntity
                                                            listBridgeBeams.Item(k) = Nothing
                                                        End If
                                                        dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam) = listBridgeBeams
                                                    End If
                                                Next k
                                            End If
                                        End If
                                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(LineShortBeam) = False Then
                                            drawingPlacementBeams.ActiveSpace.Entities.Add(LineShortBeam)
                                            dataBeam.DWGEntity = LineShortBeam
                                        End If
                                        LineShortBeam.StartPoint = startAlignPoint3d
                                        LineShortBeam.EndPoint = endAlignPoint3d
                                        'находим зазор и обрезаем балку
                                        userBeam.numberProlet = numberProlet
                                        userBeam.numberRow = numberRowBeam
                                        userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                        userBeam.model = modelBeam
                                        userBeam.nameAlbum = nameAlbumBeam
                                        Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                        Dim listProjectBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                        Dim dataBeamStructure As StructureElement = listProjectBeams.Item(numberRowBeam)
                                        dataBeamStructure.DWGEntity = LineShortBeam
                                        Dim boolRecData As Boolean = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                        For k As Integer = 0 To 2
                                            Dim userListClearence As List(Of Double) = New List(Of Double)
                                            Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, LineShortBeam, userListClearence)
                                            Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                            If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                                Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(LineShortBeam, deltaTrimBeam)
                                            End If
                                            'корректируем балку по высоте
                                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                            'делаем коррекцию балки в плане
                                            Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPline3D, theoryShortLineBeam)
                                        Next k
                                        'возвращаем значение в словарь
                                        userBeam._elementBridgePoint.StartAxisPoint = LineShortBeam.EndPoint
                                        userBeam._elementBridgePoint.EndAxisPoint = LineShortBeam.StartPoint
                                        strGSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                        dataBeamStructure.KeyParameter = strGSON
                                        boolRecData = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                        listProjectBeams.Item(numberRowBeam) = dataBeamStructure
                                        dictionaryBeams.Item(numberProlet) = listProjectBeams
                                        prevUserBeam = userBeam
                                        prevLineShortBeam = LineShortBeam
                                    End If
                                End If
                            End If
                        Next j
                        If IsNothing(axisPline) = False Then
                            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                        End If
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'раскладка произвольными балками
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ElseIf userBridge.TypeBridge = Bridges.typePlacementBeam.float Then
                Dim arrayPr1 As String(,) = Nothing
                Dim countArrayPr1 As Integer = 0
                Dim boolWriteArray As Boolean = False
                If IsNothing(dictionaryPillars) = False Then
                    If dictionaryPillars.Count > 0 Then
                        'делаем смещение всех опор моста (если такое смещение задано
                        For i As Integer = 0 To dictionaryPillars.Count - 1
                            Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key
                            Dim listPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value
                            If IsNothing(listPillar) = False Then
                                If listPillar.Count > 0 Then
                                    Dim dataPillarStructure As StructureElement = listPillar.Item(1)
                                    If IsNothing(dataPillarStructure) = False Then
                                        Dim acLineAxisPillar As DwgLine = dataPillarStructure.DWGEntity
                                        If IsNothing(acLineAxisPillar) = False Then
                                            If acLineAxisPillar.Length > 0 Then
                                                Dim userPillar As Pillar = dataPillarStructure.getPillar()
                                                If IsNothing(userPillar) = False Then
                                                    ReDim Preserve arrayPr1(3, countArrayPr1)
                                                    arrayPr1(0, countArrayPr1) = numberPillar 'номер опоры
                                                    arrayPr1(1, countArrayPr1) = Val(userPillar.Clearence) 'левый зазор
                                                    arrayPr1(2, countArrayPr1) = Val(userPillar.RightClearence) 'правый зазор
                                                    If Val(userPillar.RightClearence) = 0 Then
                                                        arrayPr1(2, countArrayPr1) = Val(userPillar.Clearence) 'правый зазор
                                                    End If
                                                    arrayPr1(3, countArrayPr1) = Val(userPillar.SiteMonolit) 'участок омоличивания балки
                                                    countArrayPr1 += 1
                                                    'делаем смещение оси опроры
                                                    If userBridge.HorizontalOffset <> 0 Then
                                                        Dim deltaMoveBridge As Double = Math.Round((userBridge.startPlacementPosition + userBridge.HorizontalOffset), 3)
                                                        If deltaMoveBridge <> 0 Then
                                                            'делаем смещение оси опоры
                                                            Dim boolMovePillar As Boolean = Pillar.moveAxisPillar(axisCentrePline3D, acLineAxisPillar, deltaMoveBridge)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If

                If IsArray(arrayPr1) = True Then
                    Dim boolLeftAxis As Boolean = True
                    For i As Integer = 0 To arrayLastAxisBeams.GetUpperBound(1)
                        'первый и второй ряды считаем по зазорам, остальные так кладем
                        Dim numberRowBeam As Integer = CInt(arrayLastAxisBeams(0, i))
                        Dim offsetPlaneRowAxis As Double = Val(arrayLastAxisBeams(1, i))
                        Dim offsetElevRowAxis As Double = Val(arrayLastAxisBeams(2, i))
                        Dim axisPline As DwgPolyline = Nothing
                        Dim hgTraskObj As String = arrayLastAxisBeams(3, i)
                        If i > 0 Then
                            boolLeftAxis = False
                        End If
                        If IsNothing(hgTraskObj) = False Then
                            Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                            If IsNothing(dataPlacementBeams) = False Then
                                If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                    axisPline = dataPlacementBeams.DWGEntity
                                    If axisPline.Length > 0 Then
                                        Dim stPoint As Vector2D = axisPline.Item(0).Vertex
                                        Dim enPoint As Vector2D = axisPline.Item(axisPline.Count - 1).Vertex
                                        Dim startPK As Double = 0
                                        Dim off As Double = 0
                                        Dim endPk As Double = 0
                                        Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                        Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                        If boolStartPk = True And boolEndPk = True Then
                                            If startPK > endPk Then
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, True)
                                            Else
                                                axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, False)
                                            End If
                                        End If
                                    Else
                                        axisPline = Nothing
                                    End If
                                End If
                            End If
                        End If
                        If IsNothing(axisPline) = True Then
                            axisPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis)
                        End If
                        Dim axisPline3D As IPolyline3D = New Polyline3D()
                        axisPline.GetPolyline(axisPline3D)
                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPline) = True Then
                            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                        End If
                        'переменная для запоминания предыдущей балки
                        Dim prevLineShortBeam As DwgLine = Nothing
                        Dim prevUserBeam As BeamI = Nothing
                        'строим балки по крайним рядам,
                        For j As Integer = 0 To arrayPr1.GetUpperBound(1) - 1
                            numberProlet = Val(arrayPr1(0, j)) 'номер пролета
                            'берем зазор
                            Dim zazor As Double = 0
                            If i = 0 Then
                                zazor = Val(arrayPr1(1, j))
                            ElseIf i = 1 Then
                                zazor = Val(arrayPr1(2, j))
                                If zazor = 0 Then
                                    zazor = Val(arrayPr1(1, j))
                                End If
                            End If
                            Dim monolitBeam1 As Double = Val(arrayPr1(3, j))
                            Dim monolitBeam2 As Double = Val(arrayPr1(3, j + 1))
                            'читаем марку балки
                            Dim nameAlbumBeam As String = ""
                            Dim modelBeam As String = ""
                            Dim userBeam As BeamI = Nothing
                            Dim dataBeam As StructureElement = Nothing
                            If dictionaryBeams.ContainsKey(numberProlet) = True Then
                                Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                If IsNothing(listBeamsProlet) = False Then
                                    If listBeamsProlet.Count > 0 Then
                                        dataBeam = listBeamsProlet.Item(numberRowBeam)
                                        If IsNothing(dataBeam) = False Then
                                            userBeam = dataBeam.getBeamI()
                                        End If
                                    End If
                                    If IsNothing(userBeam) = True Then
                                        MsgBox("Не удалос прочитать параметры крайней балки в пролете №" & numberProlet & " , ряде №" & numberRowBeam)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    'записываем высоту балки над поверхностью
                                    userBeam.offsetSurface = offsetElevRowAxis
                                    userBeam.startLenghtMonolith = Math.Round(monolitBeam1, 3)
                                    userBeam.endLenghtMonolith = Math.Round(monolitBeam2, 3)
                                    userBeam.clearence = zazor
                                    'находим начальную и конечную ось опоры в пролете
                                    Dim acLineAxisPillar1 As DwgLine = Nothing 'ось опоры первая
                                    Dim acLineAxisPillar2 As DwgLine = Nothing 'ось опоры вторая
                                    Dim acLineAxisPillarBeam1 As DwgLine = New DwgLine() 'оси опирания 1
                                    Dim acLineAxisPillarBeam2 As DwgLine = New DwgLine() 'оси опирания 2
                                    If dictionaryPillars.ContainsKey(numberProlet) = True Then
                                        Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet)
                                        If IsNothing(dataListPillars) = False Then
                                            If dataListPillars.Count > 2 Then
                                                acLineAxisPillar1 = dataListPillars.Item(1).DWGEntity
                                                If IsNothing(dataListPillars.Item(2).DWGEntity) = False Then
                                                    acLineAxisPillarBeam1 = dataListPillars.Item(2).DWGEntity
                                                End If
                                            End If
                                        End If
                                    End If
                                    If dictionaryPillars.ContainsKey(numberProlet + 1) = True Then
                                        Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet + 1)
                                        If IsNothing(dataListPillars) = False Then
                                            If dataListPillars.Count > 2 Then
                                                acLineAxisPillar2 = dataListPillars.Item(1).DWGEntity
                                                If IsNothing(dataListPillars.Item(0).DWGEntity) = False Then
                                                    acLineAxisPillarBeam2 = dataListPillars.Item(0).DWGEntity
                                                End If
                                            End If
                                        End If
                                    End If

                                    If IsNothing(acLineAxisPillar1) = True Then Continue For
                                    If IsNothing(acLineAxisPillar2) = True Then Continue For
                                    Dim acLineShortBeam As DwgLine = New DwgLine()
                                    'создаем ось балки или корректируем старую
                                    Dim dataStrucrureAxisBeam As StructureElement = CalculationBeams.getAxisBeam(dictionaryObjectBridge, numberProlet, numberRowBeam)
                                    If IsNothing(dataStrucrureAxisBeam) = False Then
                                        If IsNothing(dataStrucrureAxisBeam.DWGEntity) = False Then
                                            acLineShortBeam = dataStrucrureAxisBeam.DWGEntity
                                        End If
                                    End If
                                    If drawingPlacementBeams.ActiveSpace.Entities.Contains(acLineShortBeam) = False Then
                                        drawingPlacementBeams.ActiveSpace.Entities.Add(acLineShortBeam)
                                        dataBeam.DWGEntity = acLineShortBeam
                                    End If
                                    If drawingPlacementBeams.ActiveSpace.Entities.Contains(acLineAxisPillarBeam1) = False Then
                                        drawingPlacementBeams.ActiveSpace.Entities.Add(acLineAxisPillarBeam1)
                                    End If
                                    If drawingPlacementBeams.ActiveSpace.Entities.Contains(acLineAxisPillarBeam2) = False Then
                                        drawingPlacementBeams.ActiveSpace.Entities.Add(acLineAxisPillarBeam2)
                                    End If
                                    'начало ряда
                                    If j = 0 Then
                                        Dim boolCorrDirection1 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar1, align)
                                        Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                        'первый пролет на крайних балках, начальная ось опирания совпадает с осью опоры
                                        Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, 0, axisPline3D, boolLeftAxis, 0, Nothing, acLineAxisPillar2, surf, userBeam)
                                        'корректируем первую ось
                                        If acLineShortBeam.Length > 0 Then
                                            userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                        End If
                                        Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                        Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                        Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisPline3D, theoryShortLineBeam)
                                        userBeam.clearence = Math.Round(zazor, 3)
                                        'последний пролет
                                    ElseIf j = arrayPr1.GetUpperBound(1) - 1 Then
                                        Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                        Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisPline3D, boolLeftAxis, 0, acLineAxisPillar2, Nothing, surf, userBeam)
                                        If acLineShortBeam.Length > 0 Then
                                            userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                        End If
                                        'находим зазор и обрезаем балку
                                        For k As Integer = 0 To 2
                                            Dim userListClearence As List(Of Double) = New List(Of Double)
                                            Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence, userBeam)
                                            Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                            If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                                Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(acLineShortBeam, deltaTrimBeam)
                                            End If
                                            'корректируем балку по высоте
                                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                            'делаем коррекцию балки в плане
                                            Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisPline3D, theoryShortLineBeam)
                                            'userBeam.clearence = Math.Round(minZazor, 3)
                                        Next k
                                    Else
                                        Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                        Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisPline3D, boolLeftAxis, 0, Nothing, acLineAxisPillar2, surf, userBeam)
                                        If acLineShortBeam.Length > 0 Then
                                            userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                        End If
                                        'находим зазор и обрезаем балку
                                        For k As Integer = 0 To 2
                                            Dim userListClearence As List(Of Double) = New List(Of Double)
                                            Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence, userBeam)
                                            Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                            If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                                Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(acLineShortBeam, deltaTrimBeam)
                                            End If
                                            'корректируем балку по высоте
                                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                            'делаем коррекцию балки в плане
                                            Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisPline3D, theoryShortLineBeam)
                                            'userBeam.clearence = Math.Round(minZazor, 3)
                                        Next k
                                    End If
                                    '==============================================================================================================
                                    'оформляем балку
                                    userBeam.numberProlet = numberProlet
                                    userBeam.numberRow = numberRowBeam
                                    Dim fullLenghtBeam As Double = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                    userBeam.lenght = fullLenghtBeam
                                    userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                    'userBeam.model = modelBeam
                                    'userBeam.nameAlbum = nameAlbumBeam
                                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                    Dim listProjectBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                                    Dim dataBeamStructure As StructureElement = listProjectBeams.Item(numberRowBeam)
                                    dataBeamStructure.DWGEntity = acLineShortBeam
                                    listProjectBeams.Item(numberRowBeam) = dataBeamStructure
                                    dictionaryBeams.Item(numberProlet) = listProjectBeams
                                    'возвращаем значение в словарь
                                    userBeam._elementBridgePoint.StartAxisPoint = acLineShortBeam.StartPoint
                                    userBeam._elementBridgePoint.EndAxisPoint = acLineShortBeam.EndPoint
                                    strGSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                    dataBeamStructure.KeyParameter = strGSON
                                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(acLineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeamStructure)
                                    prevLineShortBeam = acLineShortBeam
                                End If
                            End If
                        Next j
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'раскладка произвольными балками с автоматическим вычислением максимального зазора
            ElseIf userBridge.TypeBridge = Bridges.typePlacementBeam.maxClearence Then
                Dim arrayPr1 As String(,) = {}
                Dim countArrayPr1 As Integer = 0
                If IsNothing(dictionaryPillars) = False Then
                    If dictionaryPillars.Count > 1 Then
                        For i As Integer = 0 To dictionaryPillars.Count - 1
                            Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key
                            Dim listPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value
                            If IsNothing(listPillar) = False Then
                                If listPillar.Count > 0 Then
                                    Dim dataPillarStructure As StructureElement = listPillar.Item(1)
                                    If IsNothing(dataPillarStructure) = False Then
                                        If IsNothing(dataPillarStructure.DWGEntity) = False Then
                                            Dim acLineAxisPillar As DwgLine = dataPillarStructure.DWGEntity
                                            If acLineAxisPillar.Length > 0 Then
                                                Dim userPillar As Pillar = dataPillarStructure.getPillar()
                                                If IsNothing(userPillar) = False Then
                                                    ReDim Preserve arrayPr1(4, countArrayPr1)
                                                    arrayPr1(0, countArrayPr1) = numberPillar 'номер опоры
                                                    arrayPr1(1, countArrayPr1) = Val(userPillar.Clearence) 'левый зазор
                                                    arrayPr1(2, countArrayPr1) = Val(userPillar.SiteMonolit) 'участок омоноличивания
                                                    arrayPr1(3, countArrayPr1) = 9999999999 'минимальная длина балки в этом пролете
                                                    arrayPr1(4, countArrayPr1) = i
                                                    countArrayPr1 += 1
                                                    'делаем смещение осей всех опор на заданную величину
                                                    If userBridge.HorizontalOffset <> 0 Then
                                                        Dim deltaMoveBridge As Double = Math.Round((userBridge.startPlacementPosition + userBridge.HorizontalOffset), 3)
                                                        If deltaMoveBridge <> 0 Then
                                                            'делаем смещение оси опоры
                                                            Dim boolMovePillar As Boolean = Pillar.moveAxisPillar(axisCentrePline3D, acLineAxisPillar, deltaMoveBridge)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                Dim tempArrayBeam As DwgLine() = {}
                Dim countTempArrayBeam As Integer = 0
                For i As Integer = 0 To 1
                    'первый и второй ряды считаем по зазорам, остальные так кладем
                    Dim numberRowBeam As Integer = CInt(arrayLastAxisBeams(0, i)) 'номер ряда
                    Dim offsetPlaneRowAxis As Double = Val(arrayLastAxisBeams(1, i)) 'горизонтальное смещение ряда относительно оси трассы
                    Dim offsetElevRowAxis As Double = Val(arrayLastAxisBeams(2, i)) 'вертикальное смещение ряда относительно оси трассы
                    Dim axisAlignPline As DwgPolyline = Nothing
                    Dim hgTraskObj As String = arrayLastAxisBeams(3, i) 'идентификатор траектории раскладки балок
                    If IsNothing(hgTraskObj) = False Then
                        Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                        If IsNothing(dataPlacementBeams) = False Then
                            If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                axisAlignPline = dataPlacementBeams.DWGEntity
                                If axisAlignPline.Length > 0 Then
                                    Dim stPoint As Vector2D = axisAlignPline.Item(0).Vertex
                                    Dim enPoint As Vector2D = axisAlignPline.Item(axisAlignPline.Count - 1).Vertex
                                    Dim startPK As Double = 0
                                    Dim off As Double = 0
                                    Dim endPk As Double = 0
                                    Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                    Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                    If boolStartPk = True And boolEndPk = True Then
                                        If startPK > endPk Then
                                            axisAlignPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisAlignPline, offsetPlaneRowAxis, True)
                                        Else
                                            axisAlignPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisAlignPline, offsetPlaneRowAxis, False)
                                        End If
                                    End If
                                Else
                                    axisAlignPline = Nothing
                                End If
                            End If
                        End If
                    End If
                    'траектории не найдены, распараллеливанм ось трассы
                    If IsNothing(axisAlignPline) = True Then
                        axisAlignPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis)
                    End If
                    'конвертируем полилинию для вычислений
                    Dim axisAlignPline3D As IPolyline3D = New Polyline3D()
                    axisAlignPline.GetPolyline(axisAlignPline3D)
                    If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisAlignPline) = True Then
                        drawingPlacementBeams.ActiveSpace.Entities.Remove(axisAlignPline)
                    End If
                    'переменная для запоминания предыдущей балки
                    Dim prevLineShortBeam As DwgLine = Nothing
                    Dim prevUserBeam As BeamI = Nothing
                    'строим балки по крайним рядам,
                    For j As Integer = 0 To arrayPr1.GetUpperBound(1) - 1
                        'номер пролета
                        numberProlet = Val(arrayPr1(0, j))
                        'берем зазор
                        Dim zazor As Double = Val(arrayPr1(1, j))
                        Dim monolitBeam1 As Double = Val(arrayPr1(2, j))
                        Dim monolitBeam2 As Double = Val(arrayPr1(2, j + 1))
                        Dim userBeam As BeamI = Nothing
                        Dim dataBeam As StructureElement = Nothing
                        If dictionaryBeams.ContainsKey(numberProlet) = True Then
                            'забираем все балки в пролете 
                            Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                            'забираем балку согласно номеру ряда
                            If IsNothing(listBeamsProlet) = False Then
                                If listBeamsProlet.Count > 0 Then
                                    dataBeam = listBeamsProlet.Item(numberRowBeam)
                                    If IsNothing(dataBeam) = False Then
                                        userBeam = dataBeam.getBeamI()
                                    End If
                                End If
                                If IsNothing(userBeam) = True Then
                                    MsgBox("Не удалос прочитать параметры крайней балки в пролете №" & numberProlet & " , ряде №" & numberRowBeam)
                                    drawingPlacementBeams.EndUpdate()
                                    Exit Sub
                                End If
                                'записываем высоту балки над поверхностью
                                userBeam.offsetSurface = offsetElevRowAxis
                                userBeam.startLenghtMonolith = Math.Round(monolitBeam1, 3)
                                userBeam.endLenghtMonolith = Math.Round(monolitBeam2, 3)
                                userBeam.clearence = zazor
                                'находим начальную и конечную ось опоры в пролете
                                Dim acLineAxisPillar1 As DwgLine = Nothing 'ось опоры первая
                                Dim acLineAxisPillar2 As DwgLine = Nothing 'ось опоры вторая
                                If dictionaryPillars.ContainsKey(numberProlet) = True Then
                                    Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet)
                                    If IsNothing(dataListPillars) = False Then
                                        If dataListPillars.Count > 2 Then
                                            acLineAxisPillar1 = dataListPillars.Item(1).DWGEntity
                                        End If
                                    End If
                                End If
                                If dictionaryPillars.ContainsKey(numberProlet + 1) = True Then
                                    Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet + 1)
                                    If IsNothing(dataListPillars) = False Then
                                        If dataListPillars.Count > 2 Then
                                            acLineAxisPillar2 = dataListPillars.Item(1).DWGEntity
                                        End If
                                    End If
                                End If
                                If IsNothing(acLineAxisPillar1) = True Then
                                    MsgBox("Не удалос прочитать параметры 1 опоры в пролете №" & numberProlet)
                                    drawingPlacementBeams.EndUpdate()
                                    Exit Sub
                                End If
                                If IsNothing(acLineAxisPillar2) = True Then
                                    MsgBox("Не удалос прочитать параметры 2 опоры в пролете №" & numberProlet)
                                    drawingPlacementBeams.EndUpdate()
                                    Exit Sub
                                End If
                                'новая ось балки
                                Dim acLineShortBeam As DwgLine = New DwgLine()
                                'создаем ось балки или корректируем старую
                                Dim dataStrucrureAxisBeam As StructureElement = CalculationBeams.getAxisBeam(dictionaryObjectBridge, numberProlet, numberRowBeam)
                                If IsNothing(dataStrucrureAxisBeam) = False Then
                                    If IsNothing(dataStrucrureAxisBeam.DWGEntity) = False Then
                                        acLineShortBeam = dataStrucrureAxisBeam.DWGEntity
                                    End If
                                End If
                                'добавляем балку в чертеж
                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(acLineShortBeam) = False Then
                                    drawingPlacementBeams.ActiveSpace.Entities.Add(acLineShortBeam)
                                    dataBeam.DWGEntity = acLineShortBeam
                                End If
                                'срабатывает на первой балке пролета или на балках которые расположены между крайними балками для которых не надо высчитывать зазор
                                If j = 0 Then
                                    'коррекция направлений осей опоры
                                    Dim boolCorrDirection1 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar1, align)
                                    Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                    'первый пролет на крайних балках, начальная ось опирания совпадает с осью опоры
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, 0, axisAlignPline3D, True, 0, Nothing, acLineAxisPillar2, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        'добавляем длину только к началу
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = acLineShortBeam.Length 'userBeam.lenght - userBeam.a - userBeam.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                    Dim BeamLenght As Double = Val(arrayPr1(3, j))
                                    If BeamLenght > acLineShortBeam.Length Then
                                        arrayPr1(3, j) = Math.Round(acLineShortBeam.Length, 2)
                                        arrayPr1(4, j) = i
                                    End If
                                    'последний пролет
                                ElseIf j = arrayPr1.GetUpperBound(1) - 1 Then
                                    Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisAlignPline3D, True, 0, acLineAxisPillar2, Nothing, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        'добавляем только к концу
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.b, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    'находим зазор и обрезаем балку
                                    For k As Integer = 0 To 2
                                        Dim userListClearence As List(Of Double) = New List(Of Double)
                                        Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence)
                                        Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                        If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                            Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(acLineShortBeam, deltaTrimBeam)
                                        End If
                                        'корректируем балку по высоте
                                        Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                        'делаем коррекцию балки в плане
                                        Dim theoryShortLineBeam As Double = acLineShortBeam.Length ' userBeam.lenght - userBeam.a - userBeam.b
                                        Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                    Next k
                                    Dim BeamLenght As Double = Val(arrayPr1(3, j))
                                    If BeamLenght > acLineShortBeam.Length Then
                                        arrayPr1(3, j) = Math.Round(acLineShortBeam.Length, 2)
                                        arrayPr1(4, j) = i
                                    End If
                                Else
                                    Dim boolCorrDirection2 As Boolean = Pillar.correctDirectionAxisPillar(acLineAxisPillar2, align)
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisAlignPline3D, True, 0, Nothing, acLineAxisPillar2, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    For k As Integer = 0 To 2
                                        Dim userListClearence As List(Of Double) = New List(Of Double)
                                        Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence)
                                        Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                        If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                            Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(acLineShortBeam, deltaTrimBeam)
                                        End If
                                        'корректируем балку по высоте
                                        Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                        'делаем коррекцию балки в плане
                                        Dim theoryShortLineBeam As Double = acLineShortBeam.Length ' userBeam.lenght - userBeam.a - userBeam.b
                                        Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                    Next k
                                    Dim BeamLenght As Double = Val(arrayPr1(3, j))
                                    If BeamLenght > acLineShortBeam.Length Then
                                        arrayPr1(3, j) = Math.Round(acLineShortBeam.Length, 2)
                                        arrayPr1(4, j) = i
                                    End If
                                End If
                                'помещаем балку во временный массив, далее на этой итерации ее удаляем
                                ReDim Preserve tempArrayBeam(countTempArrayBeam)
                                tempArrayBeam(countTempArrayBeam) = acLineShortBeam
                                countTempArrayBeam += 1
                                prevLineShortBeam = acLineShortBeam
                            End If
                        End If
                    Next j
                Next i
                'удаляем предварительно построенные балки
                If IsArray(tempArrayBeam) = True Then
                    For i As Integer = 0 To tempArrayBeam.Length - 1
                        Dim lineDelete As DwgLine = tempArrayBeam(i)
                        drawingPlacementBeams.ActiveSpace.Entities.Remove(lineDelete)
                    Next i
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'минимальная длина вычислена, заново раскладываем балки
                For i As Integer = 0 To arrayLastAxisBeams.GetUpperBound(1)
                    Dim numberRowBeam As Integer = CInt(arrayLastAxisBeams(0, i))
                    Dim offsetPlaneRowAxis As Double = Val(arrayLastAxisBeams(1, i))
                    Dim offsetElevRowAxis As Double = Val(arrayLastAxisBeams(2, i))
                    Dim axisAlignPline As DwgPolyline = Nothing
                    Dim hgTraskObj As String = arrayLastAxisBeams(3, i)
                    Dim boolLeftBeam As Boolean = True
                    If i > 0 Then boolLeftBeam = False
                    If IsNothing(hgTraskObj) = False Then
                        Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                        If IsNothing(dataPlacementBeams) = False Then
                            If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                axisAlignPline = dataPlacementBeams.DWGEntity
                                If axisAlignPline.Length > 0 Then
                                    Dim stPoint As Vector2D = axisAlignPline.Item(0).Vertex
                                    Dim enPoint As Vector2D = axisAlignPline.Item(axisAlignPline.Count - 1).Vertex
                                    Dim startPK As Double = 0
                                    Dim off As Double = 0
                                    Dim endPk As Double = 0
                                    Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                    Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                    If boolStartPk = True And boolEndPk = True Then
                                        If startPK > endPk Then
                                            axisAlignPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisAlignPline, offsetPlaneRowAxis, True)
                                        Else
                                            axisAlignPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisAlignPline, offsetPlaneRowAxis, False)
                                        End If
                                    End If
                                Else
                                    axisAlignPline = Nothing
                                End If
                            End If
                        End If
                    End If
                    If IsNothing(axisAlignPline) = True Then
                        axisAlignPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis)
                    End If
                    Dim axisAlignPline3D As IPolyline3D = New Polyline3D()
                    axisAlignPline.GetPolyline(axisAlignPline3D)
                    If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisAlignPline) = True Then
                        drawingPlacementBeams.ActiveSpace.Entities.Remove(axisAlignPline)
                    End If
                    'строим балки по крайним рядам,
                    Dim prevLineShortBeam As DwgLine = Nothing
                    For j As Integer = 0 To arrayPr1.GetUpperBound(1) - 1
                        'номер пролета
                        numberProlet = Val(arrayPr1(0, j))
                        'берем зазор
                        Dim zazor As Double = Val(arrayPr1(1, j))
                        Dim monolitBeam1 As Double = Val(arrayPr1(2, j))
                        Dim monolitBeam2 As Double = Val(arrayPr1(2, j + 1))
                        Dim lenghtBeam As Double = Val(arrayPr1(3, j)) 'длина балки
                        Dim nClerence As Double = Val(arrayPr1(4, j)) 'ряд
                        'читаем марку балки
                        Dim userBeam As BeamI = Nothing
                        Dim dataBeam As StructureElement = Nothing
                        If dictionaryBeams.ContainsKey(numberProlet) = True Then
                            Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberProlet)
                            If IsNothing(listBeamsProlet) = False Then
                                If listBeamsProlet.Count > 0 Then
                                    dataBeam = listBeamsProlet.Item(numberRowBeam)
                                    If IsNothing(dataBeam) = False Then
                                        userBeam = dataBeam.getBeamI()
                                    End If
                                End If
                                If IsNothing(userBeam) = True Then
                                    MsgBox("Не удалос прочитать параметры крайней балки в пролете №" & numberProlet & " , ряде №" & numberRowBeam)
                                    drawingPlacementBeams.EndUpdate()
                                    'если не удалось прочитать свойства крайней балки, это ошибка выходим из программы
                                    Exit Sub
                                End If
                                'записываем высоту балки над поверхностью
                                userBeam.offsetSurface = offsetElevRowAxis
                                userBeam.clearence = Math.Round(zazor, 3)
                                userBeam.startLenghtMonolith = Math.Round(monolitBeam1, 3)
                                userBeam.endLenghtMonolith = Math.Round(monolitBeam2, 3)
                                'находим оси опор
                                Dim acLineAxisPillar1 As DwgLine = Nothing 'ось опоры первая
                                Dim acLineAxisPillar2 As DwgLine = Nothing 'ось опоры вторая
                                If dictionaryPillars.ContainsKey(numberProlet) = True Then
                                    Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet)
                                    If IsNothing(dataListPillars) = False Then
                                        If dataListPillars.Count > 2 Then
                                            acLineAxisPillar1 = dataListPillars.Item(1).DWGEntity
                                        End If
                                    End If
                                End If
                                If dictionaryPillars.ContainsKey(numberProlet + 1) = True Then
                                    Dim dataListPillars As List(Of StructureElement) = dictionaryPillars.Item(numberProlet + 1)
                                    If IsNothing(dataListPillars) = False Then
                                        If dataListPillars.Count > 2 Then
                                            acLineAxisPillar2 = dataListPillars.Item(1).DWGEntity
                                        End If
                                    End If
                                End If
                                Dim acLineShortBeam As DwgLine = New DwgLine()
                                If IsNothing(dataBeam.DWGEntity) = False Then
                                    acLineShortBeam = dataBeam.DWGEntity
                                End If
                                'находим старую ось балки и забираем из нее линию
                                Dim dataStrucrureAxisBeam As StructureElement = CalculationBeams.getAxisBeam(dictionaryObjectBridge, numberProlet, numberRowBeam)
                                If IsNothing(dataStrucrureAxisBeam) = False Then
                                    If IsNothing(dataStrucrureAxisBeam.DWGEntity) = False Then
                                        acLineShortBeam = dataStrucrureAxisBeam.DWGEntity
                                    End If
                                End If
                                'добавляем балку в чертеж
                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(acLineShortBeam) = False Then
                                    drawingPlacementBeams.ActiveSpace.Entities.Add(acLineShortBeam)
                                    dataBeam.DWGEntity = acLineShortBeam
                                End If
                                'срабатывает на первой балке пролета или на балках которые расположены между крайними балками для которых не надо высчитывать зазор
                                If j = 0 Then 'первый пролет на крайних балках, начальная ось опирания совпадает с осью опоры
                                    'строим ось балки
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, 0, axisAlignPline3D, boolLeftBeam, lenghtBeam, Nothing, Nothing, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                    'теоретическая длина балки
                                    Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                    'делаем коррекцию балки в плане
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                    'последний пролет
                                ElseIf j = arrayPr1.GetUpperBound(1) - 1 Then
                                    'строим ось балки
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisAlignPline3D, boolLeftBeam, lenghtBeam, Nothing, Nothing, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    'находим зазор и обрезаем балку
                                    Dim userListClearence As List(Of Double) = New List(Of Double)
                                    Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence, userBeam)
                                    If i = nClerence Then
                                        Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                        If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                            Dim boolTrimBeam As Boolean = BridgeGeometry.extendBeam(acLineShortBeam, -1 * deltaTrimBeam, 0, 3)
                                        End If
                                    End If
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = userBeam.lenght - userBeam.a - userBeam.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                Else
                                    Dim boolCreateBeam As Boolean = CalculationBeams.createMiddleAxisBeam(acLineShortBeam, acLineAxisPillar1, userBeam.a, axisAlignPline3D, boolLeftBeam, lenghtBeam, Nothing, Nothing, surf, userBeam)
                                    If acLineShortBeam.Length > 0 Then
                                        userBeam.lenght = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                    Else
                                        MsgBox("Не удалось вычислить длину балки №" & userBeam.numberRow)
                                        drawingPlacementBeams.EndUpdate()
                                        Exit Sub
                                    End If
                                    Dim userListClearence As List(Of Double) = New List(Of Double)
                                    Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, acLineShortBeam, userListClearence, userBeam)
                                    If i = nClerence Then
                                        Dim deltaTrimBeam As Double = Math.Round(userBeam.clearence - minZazor, 3)
                                        If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                            Dim boolTrimBeam As Boolean = BridgeGeometry.extendBeam(acLineShortBeam, -1 * deltaTrimBeam, 0, 3)
                                        End If
                                    End If
                                    userBeam.clearence = Math.Round(minZazor, 3)
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(acLineShortBeam, userBeam, surf, userBeam.offsetSurface)
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = acLineShortBeam.Length ' userBeam.lenght - userBeam.a - userBeam.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(acLineShortBeam, axisAlignPline3D, theoryShortLineBeam)
                                End If
                                '==============================================================================================================
                                'оформляем балку
                                Dim fullLenghtBeam As Double = Math.Round(acLineShortBeam.Length + userBeam.a + userBeam.b, 3)
                                userBeam.lenght = fullLenghtBeam
                                userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                                '==============================================================================================================
                                'оформляем балку
                                dataBeam = BeamI.createAxis(idBridge)
                                dataBeam.DWGEntity = acLineShortBeam
                                userBeam._elementBridgePoint.StartAxisPoint = acLineShortBeam.StartPoint
                                userBeam._elementBridgePoint.EndAxisPoint = acLineShortBeam.EndPoint
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                dataBeam.KeyParameter = strGSON
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(acLineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeam)
                                listBeamsProlet.Item(numberRowBeam) = dataBeam
                                dictionaryBeams.Item(numberProlet) = listBeamsProlet
                                prevLineShortBeam = acLineShortBeam
                            End If
                        End If
                    Next j
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '2.расставляем оси опор и оси опирания балок
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim prevLeftAxisBeam As DwgLine = Nothing 'предыдущая крайняя левая балка 
            Dim prevUserLeftBeam As BeamI = Nothing
            Dim prevRightAxisBeam As DwgLine = Nothing 'предыдущая крайняя правая балка 
            Dim prevUserRightBeam As BeamI = Nothing
            If IsNothing(dictionaryPillars) = False Then
                If dictionaryPillars.Count > 0 Then
                    For i As Integer = 0 To dictionaryPillars.Count - 1
                        Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key 'номер опоры
                        Dim listPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value 'список осей
                        If IsNothing(listPillar) = True Then Continue For
                        If listPillar.Count > 0 Then
                            Dim axisPillar As DwgLine = listPillar.Item(1).DWGEntity 'ось опоры
                            Dim axisPrevBeamsPillar As DwgLine = listPillar.Item(0).DWGEntity 'ось опирания балок предыдущего пролета
                            Dim axisBeamsPillar As DwgLine = listPillar.Item(2).DWGEntity 'ось опирания балок последующего пролета
                            If dictionaryBeams.ContainsKey(numberPillar) = True Then
                                'получаем список балок текущего пролета
                                Dim listBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberPillar)
                                If IsNothing(listBeams) = True Then Continue For
                                If listBeams.Count = 0 Then Continue For
                                'крайняя левая балка
                                Dim leftAxisBeam As DwgLine = listBeams.First.Value.DWGEntity
                                Dim userLeftBeam As BeamI = Nothing
                                If IsNothing(leftAxisBeam) = False Then
                                    If leftAxisBeam.Length > 0 Then
                                        'проверяе направление балки (по ходу пикетажа)
                                        Dim boolCorrDirection As Boolean = AxisBeamsPillars.correctionAxisDirectionBeam(leftAxisBeam, align)
                                        Dim dataBeam As StructureElement = New StructureElement
                                        Dim boolFindStruct As Boolean = FuncXRecords.getXRecords(leftAxisBeam, dataBeam)
                                        userLeftBeam = dataBeam.getBeamI
                                    Else
                                        MsgBox("Не удалось вычислить длину балки.")
                                        Continue For
                                    End If
                                Else
                                    Continue For
                                End If
                                'крайняя правая балка
                                Dim rightAxisBeam As DwgLine = Nothing
                                Dim userRightBeam As BeamI = Nothing
                                If listPillar.Count > 1 Then
                                    rightAxisBeam = listBeams.Last.Value.DWGEntity
                                    If IsNothing(rightAxisBeam) = False Then
                                        If rightAxisBeam.Length > 0 Then
                                            Dim boolCorrDirection As Boolean = AxisBeamsPillars.correctionAxisDirectionBeam(rightAxisBeam, align)
                                            Dim dataBeam As StructureElement = New StructureElement
                                            Dim boolFindStruct As Boolean = FuncXRecords.getXRecords(rightAxisBeam, dataBeam)
                                            userRightBeam = dataBeam.getBeamI
                                        End If
                                    End If
                                End If
                                'если первая опора или последняя опора (ось опоры - это ось опирания балок)
                                If numberPillar = 1 Then
                                    If IsNothing(userRightBeam) = False Then
                                        'мост состоит из более 1 рядов
                                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                            If IsNothing(axisPillar) = True Then
                                                axisPillar = New DwgLine
                                            End If
                                            'добавляем ось опоры в базу
                                            drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                        End If
                                        'ось строим по концам балок
                                        axisPillar.StartPoint = leftAxisBeam.StartPoint
                                        axisPillar.EndPoint = rightAxisBeam.StartPoint
                                    Else
                                        'мост имеем только один ряд
                                        Dim sta As Double = 0
                                        Dim off As Double = 0
                                        If IsNothing(leftAxisBeam) = False Then
                                            Dim boolPk As Boolean = axisCentrePline3D.PosToStaOffset(leftAxisBeam.StartPoint, sta, off)
                                            If boolPk = True Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                                    If IsNothing(axisPillar) = True Then
                                                        axisPillar = New DwgLine
                                                    End If
                                                    drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                                End If
                                                Dim pt1 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, -1 * userBridge.LeftStructureWidth - deltaAxisPillar) 'выпуск за пределы габарита
                                                Dim pt2 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, userBridge.RightStructureWidth + deltaAxisPillar)
                                                axisPillar.StartPoint = New Vector3D(pt1, leftAxisBeam.StartPoint.Z)
                                                axisPillar.EndPoint = New Vector3D(pt2, leftAxisBeam.StartPoint.Z)
                                            End If
                                        End If
                                    End If
                                    listPillar.Item(1).DWGEntity = axisPillar
                                Else
                                    'промежуточная опора
                                    If IsNothing(userRightBeam) = False Then 'мост состоит из более 1 рядов
                                        'строим ось опирания 2
                                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisBeamsPillar) = False Then
                                            If IsNothing(axisBeamsPillar) = True Then
                                                axisBeamsPillar = New DwgLine
                                            End If
                                            drawingPlacementBeams.ActiveSpace.Entities.Add(axisBeamsPillar)
                                        End If
                                        axisBeamsPillar.StartPoint = leftAxisBeam.StartPoint
                                        axisBeamsPillar.EndPoint = rightAxisBeam.StartPoint
                                        'строим ось опирания 1
                                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPrevBeamsPillar) = False Then
                                            If IsNothing(axisPrevBeamsPillar) = True Then
                                                axisPrevBeamsPillar = New DwgLine
                                            End If
                                            drawingPlacementBeams.ActiveSpace.Entities.Add(axisPrevBeamsPillar)
                                        End If
                                        If IsNothing(prevLeftAxisBeam) = False And IsNothing(prevRightAxisBeam) = False Then
                                            axisPrevBeamsPillar.StartPoint = prevLeftAxisBeam.EndPoint
                                            axisPrevBeamsPillar.EndPoint = prevRightAxisBeam.EndPoint
                                            'строим ось опоры
                                            'находим среднюю точку слева
                                            Dim middlePointLeft As Vector3D = funcBridges.calculateMiddlePointBeams(prevLeftAxisBeam, leftAxisBeam)
                                            'находим среднюю точку справа
                                            Dim middlePointRight As Vector3D = funcBridges.calculateMiddlePointBeams(prevRightAxisBeam, rightAxisBeam)
                                            'строимм ось опоры
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                                If IsNothing(axisPillar) = True Then
                                                    axisPillar = New DwgLine
                                                End If
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                            End If
                                            axisPillar.StartPoint = middlePointLeft
                                            axisPillar.EndPoint = middlePointRight
                                            'удлинняем ось временно на величину габарита моста
                                            Dim boolExtendAxis As Boolean = BridgeGeometry.extendBeam(axisPillar, userBridge.LeftStructureWidth, userBridge.RightStructureWidth)
                                            'находим пересечение оси опоры с левым габаритом
                                            Dim pointLeftIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                                            Dim pointRightIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                                            If pointLeftIntersectCollection.Count > 0 And pointRightIntersectCollection.Count > 0 Then
                                                axisPillar.StartPoint = New Vector3D(pointLeftIntersectCollection(0), middlePointLeft.Z)
                                                axisPillar.EndPoint = New Vector3D(pointRightIntersectCollection(0), middlePointRight.Z)
                                                Dim boolExtendAxisPillar As Boolean = BridgeGeometry.extendBeam(axisPillar, deltaAxisPillar, deltaAxisPillar)
                                            End If
                                        End If
                                    Else
                                        'мост имеем только один ряд
                                        Dim sta As Double = 0
                                        Dim off As Double = 0
                                        Dim middlePointLeft As Vector3D = funcBridges.calculateMiddlePointBeams(prevLeftAxisBeam, leftAxisBeam)
                                        Dim boolPk As Boolean = axisCentrePline3D.PosToStaOffset(middlePointLeft, sta, off)
                                        If boolPk = True Then
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                                If IsNothing(axisPillar) = True Then
                                                    axisPillar = New DwgLine
                                                End If
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                            End If
                                            Dim pt1 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, -1 * userBridge.LeftStructureWidth - deltaAxisPillar)
                                            Dim pt2 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, userBridge.RightStructureWidth + deltaAxisPillar)
                                            axisPillar.StartPoint = New Vector3D(pt1, middlePointLeft.Z)
                                            axisPillar.EndPoint = New Vector3D(pt2, middlePointLeft.Z)
                                        End If
                                    End If
                                    listPillar.Item(1).DWGEntity = axisPillar
                                    listPillar.Item(0).DWGEntity = axisPrevBeamsPillar
                                    listPillar.Item(2).DWGEntity = axisBeamsPillar
                                End If
                                prevLeftAxisBeam = leftAxisBeam
                                prevUserLeftBeam = userLeftBeam
                                prevRightAxisBeam = rightAxisBeam
                                prevUserRightBeam = userRightBeam
                            Else
                                If numberPillar = dictionaryPillars.Count Then
                                    'последняя опора
                                    If IsNothing(prevRightAxisBeam) = False Then
                                        'мост состоит из более 1 рядов
                                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                            If IsNothing(axisPillar) = True Then
                                                axisPillar = New DwgLine
                                            End If
                                            drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                        End If
                                        axisPillar.StartPoint = prevLeftAxisBeam.EndPoint
                                        axisPillar.EndPoint = prevRightAxisBeam.EndPoint
                                    Else
                                        'мост имеем только один ряд
                                        Dim sta As Double = 0
                                        Dim off As Double = 0
                                        Dim boolPk As Boolean = axisCentrePline3D.PosToStaOffset(prevLeftAxisBeam.EndPoint, sta, off)
                                        If boolPk = True Then
                                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPillar) = False Then
                                                If IsNothing(axisPillar) = True Then
                                                    axisPillar = New DwgLine
                                                End If
                                                drawingPlacementBeams.ActiveSpace.Entities.Add(axisPillar)
                                            End If
                                            Dim pt1 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, -1 * userBridge.LeftStructureWidth - deltaAxisPillar)
                                            Dim pt2 As Vector2D = axisCentrePline3D.StaOffsetToPos(sta, userBridge.RightStructureWidth + deltaAxisPillar)
                                            axisPillar.StartPoint = New Vector3D(pt1, prevLeftAxisBeam.EndPoint.Z)
                                            axisPillar.EndPoint = New Vector3D(pt2, prevLeftAxisBeam.EndPoint.Z)
                                        End If
                                    End If
                                    listPillar.Item(1).DWGEntity = axisPillar
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '3. Расставляем промежуточные балки
            If IsNothing(arrayMiddleAxisBeams) = False Then
                For i As Integer = 0 To arrayMiddleAxisBeams.GetUpperBound(1)
                    Dim numberRowBeam As Integer = CInt(arrayMiddleAxisBeams(0, i))
                    Dim offsetPlaneRowAxis As Double = Val(arrayMiddleAxisBeams(1, i))
                    Dim offsetElevRowAxis As Double = Val(arrayMiddleAxisBeams(2, i))
                    Dim axisPline As DwgPolyline = Nothing
                    Dim axisPline3D As IPolyline3D = New Polyline3D()
                    Dim hgTraskObj As String = arrayMiddleAxisBeams(3, i)
                    If IsNothing(hgTraskObj) = False Then
                        Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryObjectBridge, numberRowBeam)
                        If IsNothing(dataPlacementBeams) = False Then
                            If IsNothing(dataPlacementBeams.DWGEntity) = False Then
                                axisPline = dataPlacementBeams.DWGEntity
                                If axisPline.Length > 0 Then
                                    Dim stPoint As Vector2D = axisPline.Item(0).Vertex
                                    Dim enPoint As Vector2D = axisPline.Item(axisPline.Count - 1).Vertex
                                    Dim startPK As Double = 0
                                    Dim off As Double = 0
                                    Dim endPk As Double = 0
                                    Dim boolStartPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(stPoint, startPK, off)
                                    Dim boolEndPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(enPoint, endPk, off)
                                    If boolStartPk = True And boolEndPk = True Then
                                        If startPK > endPk Then
                                            axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, True)
                                        Else
                                            axisPline = FuncAlignment.FuncReversePolyline(drawingPlacementBeams, axisPline, offsetPlaneRowAxis, False)
                                        End If
                                    End If
                                    axisPline.GetPolyline(axisPline3D)
                                Else
                                    axisPline = Nothing
                                End If
                            End If
                        End If
                    End If

                    If IsNothing(axisPline) = True Then
                        axisPline = FuncAlignment.FuncOffsetAlignment(drawingPlacementBeams, align, offsetPlaneRowAxis)
                        axisPline.GetPolyline(axisPline3D)
                        If drawingPlacementBeams.ActiveSpace.Entities.Contains(axisPline) = True Then
                            drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPline)
                        End If
                    End If

                    Dim prevLineShortBeam As DwgLine = Nothing
                    For j As Integer = 0 To dictionaryPillars.Count - 2
                        Dim numberFirstPillar As Integer = dictionaryPillars.ElementAt(j).Key
                        Dim listFirstAxisPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(j).Value
                        Dim numberSecondPillar As Integer = dictionaryPillars.ElementAt(j + 1).Key
                        Dim listSecondAxisPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(j + 1).Value

                        Dim userDataPillar1 As StructureElement = listFirstAxisPillar.Item(1)
                        Dim userPillar1 As Pillar = Nothing
                        If IsNothing(userDataPillar1) = False Then
                            userPillar1 = userDataPillar1.getPillar
                        End If

                        Dim userDataPillar2 As StructureElement = listSecondAxisPillar.Item(1)
                        Dim userPillar2 As Pillar = Nothing
                        If IsNothing(userDataPillar2) = False Then
                            userPillar2 = userDataPillar2.getPillar
                        End If

                        Dim dictProletBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.Item(numberFirstPillar)
                        Dim userBeam As BeamI = Nothing
                        If listFirstAxisPillar.Count > 0 And listSecondAxisPillar.Count > 0 Then
                            Dim dataBeam As StructureElement = Nothing
                            If dictProletBeams.ContainsKey(numberRowBeam) = True Then
                                dataBeam = dictProletBeams.Item(numberRowBeam)
                                userBeam = dataBeam.getBeamI
                            End If
                            If IsNothing(dataBeam) = True Then
                                Continue For
                            End If
                            If IsNothing(userBeam) = True Then
                                Continue For
                            End If
                            'записываем высоту балки над поверхностью
                            userBeam.offsetSurface = offsetElevRowAxis
                            'номер ряда для записи в массив
                            Dim numberWriteRowBeam As Integer = i + 1
                            'балки промежуточных рядов (строятся последними, после крайних рядов)
                            Dim LineShortBeam As DwgLine = New DwgLine()
                            Dim listBridgeBeams As List(Of StructureElement) = New List(Of StructureElement)
                            If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                                listBridgeBeams = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                                If listBridgeBeams.Count > 0 Then
                                    For k As Integer = 0 To listBridgeBeams.Count - 1
                                        Dim tempData As StructureElement = listBridgeBeams.Item(k)
                                        If IsNothing(tempData) = False Then
                                            Dim tempBeam As BeamI = tempData.getBeamI
                                            If tempBeam.numberProlet = numberProlet And tempBeam.numberRow = numberRowBeam Then
                                                LineShortBeam = tempData.DWGEntity
                                                listBridgeBeams.Item(k) = Nothing
                                            End If
                                        End If
                                    Next k
                                End If
                            End If
                            If drawingPlacementBeams.ActiveSpace.Entities.Contains(LineShortBeam) = False Then
                                If IsNothing(LineShortBeam) = False Then
                                    LineShortBeam = New DwgLine()
                                End If
                                drawingPlacementBeams.ActiveSpace.Entities.Add(LineShortBeam)
                            End If
                            If j = 0 And dictionaryPillars.Count > 2 Then
                                'первый полет
                                Dim acLineAxisPillar1 As DwgLine = listFirstAxisPillar.Item(1).DWGEntity
                                Dim acLineAxisPillarBeam2 As DwgLine = listSecondAxisPillar.Item(0).DWGEntity
                                Dim boolCreateAxisBridge As Boolean = CalculationBeams.createMiddleAxisBeam(LineShortBeam, acLineAxisPillar1, 0, axisPline3D, True, 0, acLineAxisPillarBeam2, Nothing, surf, userBeam)
                            ElseIf j = 0 And dictionaryPillars.Count = 2 Then
                                'мост однопролетный
                                Dim acLineAxisPillar1 As DwgLine = listFirstAxisPillar.Item(1).DWGEntity
                                Dim acLineAxisPillar2 As DwgLine = listSecondAxisPillar.Item(1).DWGEntity
                                Dim boolCreateAxisBridge As Boolean = CalculationBeams.createMiddleAxisBeam(LineShortBeam, acLineAxisPillar1, 0, axisPline3D, True, 0, acLineAxisPillar2, Nothing, surf, userBeam)
                            ElseIf j = dictionaryPillars.Count - 2 Then
                                Dim acLineAxisPillar1 As DwgLine = listFirstAxisPillar.Item(2).DWGEntity
                                Dim acLineAxisPillar2 As DwgLine = listSecondAxisPillar.Item(1).DWGEntity
                                Dim boolCreateAxisBridge As Boolean = CalculationBeams.createMiddleAxisBeam(LineShortBeam, acLineAxisPillar1, 0, axisPline3D, True, 0, acLineAxisPillar2, Nothing, surf, userBeam)
                                Dim userListClearence As List(Of Double) = New List(Of Double)
                            Else
                                Dim acLineAxisPillarBeam1 As DwgLine = listFirstAxisPillar.Item(2).DWGEntity
                                Dim acLineAxisPillarBeam2 As DwgLine = listSecondAxisPillar.Item(0).DWGEntity
                                Dim boolCreateAxisBridge As Boolean = CalculationBeams.createMiddleAxisBeam(LineShortBeam, acLineAxisPillarBeam1, 0, axisPline3D, True, 0, acLineAxisPillarBeam2, Nothing, surf, userBeam)
                            End If
                            'делаем коррекцию балки в плане
                            Dim theoryShortLineBeam As Double = LineShortBeam.Length
                            If userBridge.TypeBridge = Bridges.typePlacementBeam.fixed Then
                                theoryShortLineBeam = userBeam.lenght - userBeam.a - userBeam.b
                            Else
                                userBeam.lenght = LineShortBeam.Length + userBeam.a + userBeam.b
                            End If
                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPline3D, theoryShortLineBeam)
                            'корректируем балку по высоте
                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeam, surf, userBeam.offsetSurface)
                            'оформляем балку
                            userBeam.numberProlet = numberFirstPillar
                            userBeam.numberRow = numberRowBeam
                            Dim fullLenghtBeam As Double = Math.Round(LineShortBeam.Length + userBeam.a + userBeam.b, 3)
                            userBeam.lenght = fullLenghtBeam
                            userBeam.axisOffset = Math.Round(offsetPlaneRowAxis, 3)
                            If IsNothing(userPillar1) = False Then
                                userBeam.startLenghtMonolith = userPillar1.SiteMonolit
                            End If
                            If IsNothing(userPillar2) = False Then
                                userBeam.endLenghtMonolith = userPillar2.SiteMonolit
                            End If
                            If j > 0 Then
                                Dim userListClearence As List(Of Double) = New List(Of Double)
                                Dim minZazor As Double = userBridge.zazorPreviousBeam(prevLineShortBeam, LineShortBeam, userListClearence, userBeam)
                                userBeam.clearence = Math.Round(minZazor, 3)
                            End If
                            userBeam._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                            userBeam._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                            dataBeam.DWGEntity = LineShortBeam
                            dataBeam.IdObject = LineShortBeam.ObjectID
                            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                            dataBeam.KeyParameter = strGSON
                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(LineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeam)
                            dictProletBeams.Item(numberRowBeam) = dataBeam
                            dictionaryBeams.Item(numberFirstPillar) = dictProletBeams
                            prevLineShortBeam = LineShortBeam
                        End If
                    Next j
                Next i
            End If

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '4. делаем оформление моста
            '===============================================================================================================================
            'делаем оформление осей опоры и осей опирания
            '4.1 оформляем оси опор и оси опирания балок
            For i As Integer = 0 To dictionaryPillars.Count - 1
                Dim numberPillar As Integer = dictionaryPillars.ElementAt(i).Key
                Dim listPillar As List(Of StructureElement) = dictionaryPillars.ElementAt(i).Value
                If listPillar.Count > 0 Then
                    For j As Integer = 0 To listPillar.Count - 1
                        Dim dataPillar As StructureElement = listPillar.Item(j)
                        Dim axisPillar As DwgLine = listPillar.Item(j).DWGEntity
                        If IsNothing(axisPillar) = False Then
                            If axisPillar.Length > 0 Then
                                If j = 1 Then
                                    Dim userAxisPillar As Pillar = dataPillar.getPillar()
                                    If IsNothing(userAxisPillar) = True Then
                                        userAxisPillar = New Pillar
                                        userAxisPillar.Number = numberPillar
                                        If numberDefinedAxisPillars = numberPillar Then
                                            userAxisPillar.Defining = True
                                        End If
                                    End If
                                    Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                                    Dim startAlignPoint As Vector2D = New Vector2D
                                    If pointIntersectCollection.Count > 0 Then
                                        startAlignPoint = pointIntersectCollection.ElementAt(0)
                                        Dim pk As Double = 0
                                        Dim off As Double = 0
                                        Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(startAlignPoint, pk, off)
                                        'для вычисления начала раскладки моста
                                        If numberPillar = 1 Then
                                            StartPositionBridge = pk
                                        Else
                                            endPositionBridge = pk
                                        End If
                                        'пересечение с левым габаритом
                                        Dim leftPoint As Vector2D = New Vector2D()
                                        Dim pointIntersectCollectionLeft As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                                        If pointIntersectCollectionLeft.Count > 0 Then
                                            leftPoint = pointIntersectCollectionLeft.ElementAt(0)
                                            Dim lenghtLeft As Double = (leftPoint - startAlignPoint).Length
                                        End If
                                        'пересечение с правым габаритом
                                        Dim RightPoint As Vector2D = New Vector2D()
                                        Dim pointIntersectCollectionRight As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                                        If pointIntersectCollectionRight.Count > 0 Then
                                            RightPoint = pointIntersectCollectionRight.ElementAt(0)
                                            Dim lenghtRight As Double = (RightPoint - startAlignPoint).Length
                                        End If
                                        If pointIntersectCollectionLeft.Count > 0 And pointIntersectCollectionRight.Count > 0 Then
                                            axisPillar.StartPoint = New Vector3D(leftPoint, axisPillar.StartPoint.Z)
                                            axisPillar.EndPoint = New Vector3D(RightPoint, axisPillar.EndPoint.Z)
                                            Dim doolExtend As Boolean = BridgeGeometry.extendBeam(axisPillar, 2, 2, 3)
                                        End If
                                        Dim boolStyle As Boolean = styleAxisPillar.setObjectStyle(axisPillar)
                                        userAxisPillar._elementBridgePoint.StartAxisPoint = axisPillar.StartPoint
                                        userAxisPillar._elementBridgePoint.EndAxisPoint = axisPillar.EndPoint
                                        If userAxisPillar.Number > 1 And userAxisPillar.Number < userBridge.ProletCount + 1 Then
                                            userAxisPillar.TypePillar = Pillar.PillarType.MiddlePillar
                                        Else
                                            userAxisPillar.TypePillar = Pillar.PillarType.LastPillar
                                        End If
                                        'Ось уже создана и добавлена в чертеж в блоке расстановки выше.
                                        'drawAxis ищет ось только в исходном dictionaryObjectBridge и при
                                        'первом построении создает вторую DwgLine с той же геометрией.
                                        'Сохраняем семантику на рассчитанной оси вместо повторного рисования.
                                        If dataPillar.Name <> StructureElement.typeObject.axisPillar Then
                                            If userAxisPillar.TypePillar = Pillar.PillarType.LastPillar Then
                                                dataPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.LastPillar)
                                            Else
                                                dataPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.MiddlePillar)
                                            End If
                                        End If
                                        dataPillar.DWGEntity = axisPillar
                                    End If
                                    Dim strGSONBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar)
                                    dataPillar.KeyParameter = strGSONBeam
                                    Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(axisPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
                                    listPillar.Item(j) = dataPillar

                                    'Удаляем точные геометрические копии, оставшиеся от предыдущих
                                    'запусков старой версии PlacementBeams. Рассчитанная axisPillar
                                    'остается единственной осью с актуальной семантикой.
                                    If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.axisPillar) Then
                                        Dim storedPillarAxes As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisPillar)
                                        Const axisEqualityTolerance As Double = 0.000001
                                        For k As Integer = 0 To storedPillarAxes.Count - 1
                                            Dim storedDataPillar As StructureElement = storedPillarAxes.Item(k)
                                            If IsNothing(storedDataPillar) Then Continue For
                                            Dim storedAxisPillar As DwgLine = storedDataPillar.DWGEntity
                                            If IsNothing(storedAxisPillar) Then Continue For
                                            If Object.ReferenceEquals(storedAxisPillar, axisPillar) Then Continue For

                                            Dim sameDirection As Boolean =
                                                (storedAxisPillar.StartPoint - axisPillar.StartPoint).Length <= axisEqualityTolerance AndAlso
                                                (storedAxisPillar.EndPoint - axisPillar.EndPoint).Length <= axisEqualityTolerance
                                            Dim reverseDirection As Boolean =
                                                (storedAxisPillar.StartPoint - axisPillar.EndPoint).Length <= axisEqualityTolerance AndAlso
                                                (storedAxisPillar.EndPoint - axisPillar.StartPoint).Length <= axisEqualityTolerance
                                            If sameDirection OrElse reverseDirection Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(storedAxisPillar) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(storedAxisPillar)
                                                End If
                                                storedPillarAxes.Item(k) = Nothing
                                            End If
                                        Next k
                                        dictionaryObjectBridge.Item(StructureElement.typeObject.axisPillar) = storedPillarAxes
                                    End If
                                Else
                                    'оформляем ось опирания балок
                                    Dim userAxisPillarBeams As AxisBeamsPillars = dataPillar.getAxisBeamsPillar
                                    If IsNothing(userAxisPillarBeams) = True Then
                                        userAxisPillarBeams = New AxisBeamsPillars
                                        userAxisPillarBeams.numberPillar = numberPillar
                                        If j = 0 Then
                                            userAxisPillarBeams.numberProlet = numberPillar - 1
                                        Else
                                            userAxisPillarBeams.numberProlet = numberPillar
                                        End If
                                    End If
                                    userAxisPillarBeams._elementBridgePoint.StartAxisPoint = axisPillar.StartPoint
                                    userAxisPillarBeams._elementBridgePoint.EndAxisPoint = axisPillar.EndPoint
                                    Dim elementAxisBeamsPillar As StructureElement = userAxisPillarBeams.drawAxis(ActivDocument, idBridge, dictionaryObjectBridge, styleAxisBeamsPillar, templateXML)
                                    listPillar.Item(j) = elementAxisBeamsPillar
                                    Dim boolStyle As Boolean = styleAxisBeamsPillar.setObjectStyle(axisPillar)
                                    Dim strGSONBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillarBeams)
                                    dataPillar.KeyParameter = strGSONBeam
                                    Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(axisPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
                                End If
                            End If
                        End If
                    Next j
                End If
                dictionaryPillars.Item(numberPillar) = listPillar
            Next i
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '4.2 делаем оформление балок
            'удаляем все модели
            Dim countDeleteModelBeams As Integer = ModelBeam.deleteAllModelBeams(drawingPlacementBeams, dictionaryObjectBridge)
            Dim arrayElementsTopBeam As UInteger() = {}
            Dim arrayElementsDownBeam As UInteger() = {}
            Dim arrayElementsTLC As UInteger() = {}
            If IsNothing(dictionaryBeams) = False Then
                If dictionaryBeams.Count > 0 Then
                    For i As Integer = 0 To dictionaryBeams.Count - 1
                        Dim listBeamProlet As Dictionary(Of Integer, StructureElement) = dictionaryBeams.ElementAt(i).Value
                        If IsNothing(listBeamProlet) = True Then Continue For
                        If listBeamProlet.Count > 0 Then
                            For j As Integer = 0 To listBeamProlet.Count - 1
                                Dim dataBeam As StructureElement = listBeamProlet.ElementAt(j).Value
                                If IsNothing(dataBeam) = True Then Continue For
                                Dim userBeam As BeamI = dataBeam.getBeamI
                                Dim acLineShortBeam As DwgLine = dataBeam.DWGEntity
                                If IsNothing(acLineShortBeam) = False Then
                                    If acLineShortBeam.Length > 0 Then
                                        'применям стиль оформления осей балок
                                        Dim boolSetStyleBeam As Boolean = styleAxisBeam.setObjectStyle(acLineShortBeam)
                                        'восстанавливаем балку
                                        If IsNothing(userBeam) = False Then
                                            Dim dictCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = CounterBeam.drawContour(ActivDocument, userBeam, idBridge, dictionaryObjectBridge, styleTopBeam, styleBottomBeam, templateXML)
                                            If dictCounter.Count = 2 Then
                                                Dim classCounterBeam As CounterBeam = New CounterBeam()
                                                '==============================================================================================
                                                'записываем точки по верху плиты балки и основания балки в ось
                                                userBeam._elementBridgePoint.ListPointModel = classCounterBeam.setCounterPoint(dictCounter.First.Value, userBeam.heightTopPlate)
                                                userBeam._elementBridgePoint.ListPointSecondModel = classCounterBeam.setCounterPoint(dictCounter.Last.Value)
                                                Dim strGSONBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                                dataBeam.KeyParameter = strGSONBeam
                                                Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(acLineShortBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBeam)
                                                '==============================================================================================
                                                'вставляем балку (ТЛС объект)
                                                Dim listSectionBeam As Dictionary(Of Integer, List(Of Vector2D)) = userBeam.getSection()
                                                If listSectionBeam.Count > 0 Then
                                                    Dim oldDataModelBeam As StructureElement = ModelBeam.getModelBeamI(dictionaryObjectBridge, numberProlet, numberRows)
                                                    If IsNothing(oldDataModelBeam) = False Then
                                                        If IsNothing(oldDataModelBeam.DWGEntity) = False Then
                                                            If ActivDocument.ActiveSpace.Entities.Contains(oldDataModelBeam.DWGEntity) = True Then
                                                                ActivDocument.ActiveSpace.Entities.Remove(oldDataModelBeam.DWGEntity)
                                                            End If
                                                        End If
                                                        End If
                                                    Dim model3dBeam As DwgModel3DElement = ModelBeam.drawModelBeamI(ActivDocument, userBeam, acLineShortBeam, listSectionBeam, idBridge, styleModelBeam, templateXML)
                                                    Dim boolSetStyleModel As Boolean = styleTopBeam.setObjectStyle(model3dBeam)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    Next i
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '4.3 Рисуем габарит сооружения и ось трассы
            '===============================================================================================================================
            If dictionaryPillars.Count > 1 Then
                'рисуем габарит сооружения
                Dim arrayAxisLine1 As List(Of StructureElement) = dictionaryPillars.First.Value
                Dim dataPillar As StructureElement = arrayAxisLine1(1)
                Dim axisLine1 As DwgLine = dataPillar.DWGEntity
                Dim arrayAxisLine2 As List(Of StructureElement) = dictionaryPillars.Last.Value
                Dim dataPillar2 As StructureElement = arrayAxisLine2(1)
                Dim axisLine2 As DwgLine = dataPillar2.DWGEntity
                Dim gb As DwgPolyline = New DwgPolyline()

                If dictionaryObjectBridge.ContainsKey(StructureElement.typeObject.boundaresBridge) = True Then
                    Dim listBridgeBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.boundaresBridge)
                    If IsNothing(listBridgeBeams) = False Then
                        If listBridgeBeams.Count > 0 Then
                            For i As Integer = 0 To listBridgeBeams.Count - 1
                                Dim dataBount As StructureElement = listBridgeBeams.Item(i)
                                If dataBount.IdStructure Like idBridge Then
                                    gb = dataBount.DWGEntity
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                Dim boolDrawBoundaryStructure As Boolean = userBridge.drawAxisAndBoundaryBridge(alignStructure, dictionaryBeams, align, gb)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'делаем оформление
                If boolDrawBoundaryStructure = True Then
                    Dim boolStyleBound As Boolean = styleBoundBridge.setObjectStyle(gb)
                    Dim elementBound As StructureElement = New StructureElement()
                    elementBound.Label = "Мосты и путепроводы"
                    elementBound.ClassObject = StructureElement.classStructure.OtherObject
                    elementBound.Name = StructureElement.typeObject.boundaresBridge
                    elementBound.Description = "Граница сооружения"
                    elementBound.KeyParameter = ""
                    elementBound.IdElement = Guid.NewGuid.ToString
                    elementBound.IdStructure = idBridge
                    elementBound.Note = ""
                    Dim boolRecDataModelbeam = FuncXRecords.setXRecords(gb, StructureElement.tableXRecords.PROJECT_STRUCTURES, elementBound)
                End If

                'рисуем главную ось сооружения
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'делаем оформление
                If IsNothing(alignStructure) = False Then
                    Dim boolStyleAxisBridge As Boolean = styleAxisBridge.setObjectStyle(alignStructure)
                    If userBridge.HorizontalOffset = 0 Then
                        userBridge.startPlacementPosition = StartPositionBridge
                    End If
                    userBridge.startPlacementPosition = StartPositionBridge
                    'userBridge.offsetHPosition = moveBridge
                    Dim strGSON1 As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBridge)
                    Dim elementAxisBridje As StructureElement = New StructureElement()
                    elementAxisBridje.Label = "Мосты и путепроводы"
                    elementAxisBridje.ClassObject = StructureElement.classStructure.OtherObject
                    elementAxisBridje.Name = StructureElement.typeObject.axisBridge
                    elementAxisBridje.Description = "Главная ось сооружения"
                    elementAxisBridje.KeyParameter = strGSON1
                    elementAxisBridje.IdElement = Guid.NewGuid.ToString
                    elementAxisBridje.IdStructure = idBridge
                    elementAxisBridje.Note = ""
                    Dim boolRecDataModelbeam = FuncXRecords.setXRecords(alignStructure, StructureElement.tableXRecords.PROJECT_STRUCTURES, elementAxisBridje)
                End If
            End If

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'удаляем лишние объекты
            If dictionaryObjectBridge.Count > 0 Then
                For i As Integer = 0 To dictionaryObjectBridge.Count - 1
                    Dim keyObjectName As StructureElement.typeObject = dictionaryObjectBridge.ElementAt(i).Key
                    If keyObjectName = StructureElement.typeObject.axisBeam Then
                        Dim listAxisBeams As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisBeam)
                        If IsNothing(listAxisBeams) = False Then
                            If listAxisBeams.Count > 0 Then
                                For j As Integer = 0 To listAxisBeams.Count - 1
                                    Dim dataBeams As StructureElement = listAxisBeams.Item(j)
                                    If IsNothing(dataBeams) = False Then
                                        Dim userBeam As BeamI = dataBeams.getBeamI
                                        Dim boolRemoveObject As Boolean = False
                                        If userBeam.numberProlet > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        ElseIf userBeam.numberRow > userBridge.RightRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userBeam.numberRow < -1 * userBridge.LeftRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userBeam.numberRow = 0 And userBridge.centerAxis = False Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataBeams.DWGEntity) = False Then
                                                drawingPlacementBeams.ActiveSpace.Entities.Remove(dataBeams.DWGEntity)
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    ElseIf keyObjectName = StructureElement.typeObject.counterTopBeam Then
                        Dim listObject As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.counterTopBeam)
                        If IsNothing(listObject) = False Then
                            If listObject.Count > 0 Then
                                For j As Integer = 0 To listObject.Count - 1
                                    Dim dataObject As StructureElement = listObject.Item(j)
                                    If IsNothing(dataObject) = False Then
                                        Dim userObject As CounterBeam = dataObject.getCounterBeam
                                        Dim boolRemoveObject As Boolean = False
                                        If userObject.numberProlet > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow > userBridge.RightRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow < -1 * userBridge.LeftRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow = 0 And userBridge.centerAxis = False Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataObject.DWGEntity) = False Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(dataObject.DWGEntity) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(dataObject.DWGEntity)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    ElseIf keyObjectName = StructureElement.typeObject.counterBottomBeam Then
                        Dim listObject As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.counterBottomBeam)
                        If IsNothing(listObject) = False Then
                            If listObject.Count > 0 Then
                                For j As Integer = 0 To listObject.Count - 1
                                    Dim dataObject As StructureElement = listObject.Item(j)
                                    If IsNothing(dataObject) = False Then
                                        Dim userObject As CounterBeam = dataObject.getCounterBeam
                                        Dim boolRemoveObject As Boolean = False
                                        If userObject.numberProlet > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow > userBridge.RightRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow < -1 * userBridge.LeftRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow = 0 And userBridge.centerAxis = False Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataObject.DWGEntity) = False Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(dataObject.DWGEntity) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(dataObject.DWGEntity)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    ElseIf keyObjectName = StructureElement.typeObject.modelBeam Then
                        Dim listObject As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.modelBeam)
                        If IsNothing(listObject) = False Then
                            If listObject.Count > 0 Then
                                For j As Integer = 0 To listObject.Count - 1
                                    Dim dataObject As StructureElement = listObject.Item(j)
                                    If IsNothing(dataObject) = False Then
                                        Dim userObject As ModelBeam = dataObject.getModelBeam
                                        Dim boolRemoveObject As Boolean = False
                                        If userObject.numberProlet > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow > userBridge.RightRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow < -1 * userBridge.LeftRowsCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberRow = 0 And userBridge.centerAxis = False Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataObject.DWGEntity) = False Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(dataObject.DWGEntity) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(dataObject.DWGEntity)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    ElseIf keyObjectName = StructureElement.typeObject.axisPillar Then
                        Dim listObject As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisPillar)
                        If IsNothing(listObject) = False Then
                            If listObject.Count > 0 Then
                                For j As Integer = 0 To listObject.Count - 1
                                    Dim dataObject As StructureElement = listObject.Item(j)
                                    If IsNothing(dataObject) = False Then
                                        Dim userObject As Pillar = dataObject.getPillar
                                        Dim boolRemoveObject As Boolean = False
                                        If userObject.Number > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataObject.DWGEntity) = False Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(dataObject.DWGEntity) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(dataObject.DWGEntity)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    ElseIf keyObjectName = StructureElement.typeObject.axisPillarBeams Then
                        Dim listObject As List(Of StructureElement) = dictionaryObjectBridge.Item(StructureElement.typeObject.axisPillarBeams)
                        If IsNothing(listObject) = False Then
                            If listObject.Count > 0 Then
                                For j As Integer = 0 To listObject.Count - 1
                                    Dim dataObject As StructureElement = listObject.Item(j)
                                    If IsNothing(dataObject) = False Then
                                        Dim userObject As AxisBeamsPillars = dataObject.getAxisBeamsPillar
                                        Dim boolRemoveObject As Boolean = False
                                        If userObject.numberPillar > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        ElseIf userObject.numberProlet > userBridge.ProletCount Then
                                            boolRemoveObject = True
                                        End If
                                        If boolRemoveObject = True Then
                                            If IsNothing(dataObject.DWGEntity) = False Then
                                                If drawingPlacementBeams.ActiveSpace.Entities.Contains(dataObject.DWGEntity) Then
                                                    drawingPlacementBeams.ActiveSpace.Entities.Remove(dataObject.DWGEntity)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    End If
                Next i
            End If
        Finally
            If IsNothing(axisCentrePline) = False Then
                drawingPlacementBeams.ActiveSpace.Entities.Remove(axisCentrePline)
            End If
            If IsNothing(axisPlineDirect) = False Then
                drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPlineDirect)
            End If
            If IsNothing(axisPlineReverse) = False Then
                drawingPlacementBeams.ActiveSpace.Entities.Remove(axisPlineReverse)
            End If
            drawingPlacementBeams.EndUpdate()
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'раскладка крайней опоры
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'предварительный расчет элементов крайней опоры
    Public Sub PlacementLastPillar(ByRef dataPillar As StructureElement, userNozzle As NozzlePillar, ByRef arraySubFerment As SubFermenters(), ByRef userCabinetWall As CabinetWallPillar, ByRef userLeftHand As HandPillar, ByRef userRightHand As HandPillar, ByRef userLeftPostcard As PostcardPillar, ByRef userRightPostcard As PostcardPillar, ByRef arrayRacks As RackPillar(), ByRef userGrillage As GrillagePillar, ByRef userPreparation As PreparationPillar, ByRef arrayPiles As PilePillar(), ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByRef projectSurface As Surface, ByRef egSurface As Surface, ByRef projectAlignment As Alignment)
        If IsNothing(dataPillar) = True Then Exit Sub
        'получаем опору
        Dim userPillar As Pillar = dataPillar.getPillar()
        If IsNothing(userPillar) = True Then Exit Sub
        'получаем ось опоры
        Dim axisLinePillar As DwgLine = dataPillar.DWGEntity
        If IsNothing(axisLinePillar) = True Then Exit Sub
        If axisLinePillar.Length = 0 Then Exit Sub
        'получаем номер опоры
        Dim numberPillar As Integer = userPillar.Number
        'получаем id сооружения
        Dim idBridge As String = dataPillar.IdStructure
        If IsNothing(idBridge) = True Then Exit Sub
        If idBridge.Trim.Length = 0 Then Exit Sub
        'ось трассы
        Dim acPolyAlign As Polyline3D = New Polyline3D
        projectAlignment.Plan.CompoundLine.ToPolyLine(acPolyAlign)
        'находим балки для выбранной опоры (0-предыдущий пролет, 1-следующий пролет)
        Dim listBeamsPillar As List(Of Dictionary(Of Integer, StructureElement)) = Pillar.getBeamsPillarByNumber(numberPillar, dictionaryBridgeElements)
        If listBeamsPillar.Count = 0 Then
            MsgBox("Не удалось найти балки для выбранного пролета мостового сооружения!!!")
            Exit Sub
        End If
        'находим балку с минимальной высотой
        Dim minElevationBeam As Double = CalculationBeams.getBeamToMinElevation(listBeamsPillar)
        If minElevationBeam = 999999 Then
            MsgBox("Не удалось найти отметку самой нижней балки!!!")
            Exit Sub
        End If
        'находим положение начала шкафной стенки
        Dim clearanceBeam As Double = 0
        If numberPillar = 1 Then
            clearanceBeam = CabinetWallPillar.getPositionCabinetWall(axisLinePillar, listBeamsPillar, True)
        Else
            clearanceBeam = CabinetWallPillar.getPositionCabinetWall(axisLinePillar, listBeamsPillar, False)
        End If
        clearanceBeam += userPillar.Clearence
        'находим отметки насадки
        userNozzle.TopElevation = Math.Round(minElevationBeam - userNozzle.MinElevationBeams, 3)
        userNozzle.BottomElevation = Math.Round(userNozzle.TopElevation - userNozzle.SecondHeight, 3)
        'из массива найденных балок находим крайние (справа и с лева
        Dim listExtremeBeams As List(Of StructureElement) = CalculationBeams.getExtrmBeamsToPillar(listBeamsPillar)
        Dim leftDataBeam As StructureElement = Nothing
        Dim rightDataBeam As StructureElement = Nothing
        If IsNothing(listExtremeBeams.Item(0)) = False Then
            leftDataBeam = listExtremeBeams.Item(0)
        ElseIf IsNothing(listExtremeBeams.Item(2)) = False Then
            leftDataBeam = listExtremeBeams.Item(2)
        End If
        If IsNothing(listExtremeBeams.Item(1)) = False Then
            rightDataBeam = listExtremeBeams.Item(1)
        ElseIf IsNothing(listExtremeBeams.Item(3)) = False Then
            rightDataBeam = listExtremeBeams.Item(3)
        End If

        Dim leftAxisBeam As DwgLine = leftDataBeam.DWGEntity
        Dim rightAxisBeam As DwgLine = rightDataBeam.DWGEntity
        'находим направления короткой стороны насадки
        userNozzle.LeftDirection = Math.Round(leftAxisBeam.Rotation, 6)
        userNozzle.RightDirection = Math.Round(rightAxisBeam.Rotation, 6)
        'вычислыем линию шкафной стенки
        Dim lineCabinetWall As DwgLine = CabinetWallPillar.calculateLineCabinetWall(axisLinePillar, clearanceBeam, leftDataBeam, rightDataBeam, numberPillar, userNozzle)
        If IsNothing(lineCabinetWall) = True Then
            MsgBox("Не удалось вычислить положение линии начала Шкафной стенки!!1")
            Exit Sub
        End If
        If lineCabinetWall.Length = 0 Then
            MsgBox("Не удалось вычислить положение линии начала Шкафной стенки!!")
            Exit Sub
        End If
        'вычисляем положение Насадки
        Dim boolCalculateNozzle As Boolean = userNozzle.calculateNozzle(lineCabinetWall, acPolyAlign)
        If boolCalculateNozzle = False Then
            MsgBox("Не удалось вычислить положение Насадки!!")
            Exit Sub
        End If
        'вычисляем позиции подферменников
        Dim dictSubFermers As Dictionary(Of Integer, Dictionary(Of Integer, SubFermenters)) = SubFermenters.calculateSubFermenters(userNozzle, Nothing, listBeamsPillar, arraySubFerment, True)
        ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        ''вычисляем положение шкафной стенки
        Dim boolCalcCabinetWall As Boolean = userCabinetWall.calculateCabinetWall(userNozzle, userLeftHand, userRightHand, projectSurface)
        If boolCalcCabinetWall = False Then
            MsgBox("Не удалось вычислить положение Шкафной стенки!!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение откосного крыла левого
        Dim boolCalcLeftHand As Boolean = userLeftHand.calculateHand(userNozzle, userCabinetWall, projectSurface)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение откосного крыла правого
        Dim boolCalcRightHand As Boolean = userRightHand.calculateHand(userNozzle, userCabinetWall, projectSurface)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение открылка левого
        Dim boolCalcLeftPostcard As Boolean = userLeftPostcard.calculatePostcard(userNozzle)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение открылка правого
        Dim boolCalcRightPostcard As Boolean = userRightPostcard.calculatePostcard(userNozzle)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем расчет стоек
        Dim dictRack As Dictionary(Of Integer, RackPillar) = Nothing
        If userPillar.PresenceRacks = True Then
            If IsArray(arrayRacks) = True Then
                dictRack = RackPillar.calculateRacks(userNozzle, Nothing, userGrillage, arrayRacks, numberPillar, userPillar.ElevationLand, egSurface)
                'вычисляем положение ростверка
                If userPillar.PresencGrillage = True Then
                    Dim boolCalcGrillage As Boolean = userGrillage.calculateGrillage(axisLinePillar, arrayRacks)
                End If
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение подготовки
        If userPillar.PresencPreparation = True Then
            Dim boolCalcPreparation As Boolean = userPreparation.calculatePreparation(userGrillage, userNozzle)
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем позицию свай
        Dim dictPile As Dictionary(Of Integer, Dictionary(Of Integer, PilePillar)) = PilePillar.calculatePile(userNozzle, userGrillage, arrayPiles, dictRack, userPillar.PileInRack)
    End Sub
    'рисование крайней опоры
    Public Sub DrawingLastPillar(ByVal dataPillar As StructureElement, userNozzle As NozzlePillar, ByRef arraySubFerment As SubFermenters(), ByRef userCabinetWall As CabinetWallPillar, ByRef userLeftHand As HandPillar, ByRef userRightHand As HandPillar, ByRef userLeftPostcard As PostcardPillar, ByRef userRightPostcard As PostcardPillar, ByRef arrayRacks As RackPillar(), ByRef userGrillage As GrillagePillar, ByRef userPreparation As PreparationPillar, ByRef arrayPiles As PilePillar(), ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByRef projectSurface As Surface, ByRef egSurface As Surface, ByRef docPileTLC As ConstructionDocument, ByVal templateXML As String)
        If IsNothing(dataPillar) = True Then Exit Sub
        'получаем опору
        Dim userPillar As Pillar = dataPillar.getPillar()
        If IsNothing(userPillar) = True Then Exit Sub
        'получаем id сооружения
        Dim idBridge As String = dataPillar.IdStructure
        If IsNothing(idBridge) = True Then Exit Sub
        If idBridge.Trim.Length = 0 Then Exit Sub
        '===============================================================================================================================
        'рисуем насадку
        If IsNothing(userNozzle) = False Then
            Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userNozzle._elementBridgePoint.ListPointModel
            If ListPointModel.Count > 3 Then
                'ось насадки
                Dim dataNozzle As StructureElement = userNozzle.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = NozzleContour.drawContoursNozzle(ActivDocument, userNozzle, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim topElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.contourNozzleTop)
                Dim BottomElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.contourNozzleBottom)
                Dim middlePt1 As Vector3D = userNozzle.getPointByCode("middlePt1", True)
                Dim middlePt2 As Vector3D = userNozzle.getPointByCode("middlePt2", True)
                Dim tempLineCabinetWall As DwgLine = New DwgLine
                tempLineCabinetWall.StartPoint = middlePt1
                tempLineCabinetWall.EndPoint = middlePt2
                Dim model As DwgModel3DElement = NozzleModel.drawModelNozzle(ActivDocument, topElement, BottomElement, tempLineCabinetWall, dataNozzle, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем подферменники
        If IsNothing(arraySubFerment) = False Then
            If arraySubFerment.Length > 0 Then
                For i As Integer = 0 To arraySubFerment.Length - 1
                    Dim userSubFerm As SubFermenters = arraySubFerment(i)
                    If userSubFerm._elementBridgePoint.ListPointModel.Count > 3 Then
                        'ось подферменника
                        Dim dataSubFerm As StructureElement = userSubFerm.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = SubFermenterContour.drawContour(ActivDocument, userSubFerm, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim topElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterSubFermentersTop)
                        Dim BottomElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterSubFermentersBottom)
                        'рисуем модель
                        Dim model As Boolean = SubFermenterModel.drawModel(ActivDocument, userSubFerm, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
        '===============================================================================================================================
        'рисуем шкафную стенку
        If IsNothing(userCabinetWall) = False Then
            If userCabinetWall._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось шкафной стенки
                Dim dataCabinetWall As StructureElement = userCabinetWall.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = CabinetWallContour.drawContours(ActivDocument, userCabinetWall, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = CabinetWallModel.drawModel(ActivDocument, userCabinetWall, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем обратный открылок левый
        If IsNothing(userLeftHand) = False Then
            If userLeftHand._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataHand As StructureElement = userLeftHand.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = HandContour.drawContours(ActivDocument, userLeftHand, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = HandModel.drawModel(ActivDocument, userLeftHand, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем обратный открылок правый
        If IsNothing(userRightHand) = False Then
            If userRightHand._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataHand As StructureElement = userRightHand.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = HandContour.drawContours(ActivDocument, userRightHand, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = HandModel.drawModel(ActivDocument, userRightHand, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем левое откосное крыло
        If IsNothing(userLeftPostcard) = False Then
            If userLeftPostcard._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataPostcard As StructureElement = userLeftPostcard.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PostcardContour.drawContour(ActivDocument, userLeftPostcard, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = PostcardModel.drawModel(ActivDocument, userLeftPostcard, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем правое откосное крыло
        If IsNothing(userRightPostcard) = False Then
            If userRightPostcard._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataPostcard As StructureElement = userRightPostcard.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PostcardContour.drawContour(ActivDocument, userRightPostcard, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = PostcardModel.drawModel(ActivDocument, userRightPostcard, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем стойки
        If IsNothing(arrayRacks) = False Then
            If arrayRacks.Length > 0 Then
                For i As Integer = 0 To arrayRacks.Length - 1
                    Dim userRack As RackPillar = arrayRacks(i)
                    If userRack._elementBridgePoint.ListPointModel.Count > 3 Then
                        'ось стойки
                        Dim dataSubFerm As StructureElement = userRack.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgEntity) = RackContour.drawContour(ActivDocument, userRack, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim model As Boolean = RackModel.drawModel(ActivDocument, userRack, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
        '===============================================================================================================================
        'рисуем ростверк
        If IsNothing(userGrillage) = False Then
            If userGrillage._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataGrillage As StructureElement = userGrillage.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = GrillageContour.drawContour(ActivDocument, userGrillage, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = GrillageModel.drawModel(ActivDocument, userGrillage, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем подготовку
        If IsNothing(userPreparation) = False Then
            If userPreparation._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataPreparation As StructureElement = userPreparation.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PreparationContour.drawContour(ActivDocument, userPreparation, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = PreparationModel.drawModel(ActivDocument, userPreparation, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем сваи
        If IsNothing(arrayPiles) = False Then
            If arrayPiles.Length > 0 Then
                For i As Integer = 0 To arrayPiles.Length - 1
                    Dim userPile As PilePillar = arrayPiles(i)
                    If userPile.Height > 0 Then
                        'ось стойки
                        Dim dataPile As StructureElement = userPile.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim drawCounter As Boolean = PileContour.drawContour(ActivDocument, userPile, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim model As Boolean = PileModel.drawModel(ActivDocument, userPile, docPileTLC, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
        Dim entAxisPillar As DwgEntity = dataPillar.DWGEntity
        If IsNothing(entAxisPillar) = False Then
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entAxisPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'раскладка промежуточной опоры
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'предварительный расчет элементов промежуточной опоры
    Public Sub PlacementMiddlePillar(ByRef dataPillar As StructureElement, userRigel As RigelPillar, ByRef arraySubFerment As SubFermenters(), ByRef arrayRacks As RackPillar(), ByRef userGrillage As GrillagePillar, ByRef userPreparation As PreparationPillar, ByRef arrayPiles As PilePillar(), ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByRef projectSurface As Surface, ByRef egSurface As Surface, ByRef projectAlignment As Alignment)
        If IsNothing(dataPillar) = True Then Exit Sub
        'получаем опору
        Dim userPillar As Pillar = dataPillar.getPillar()
        If IsNothing(userPillar) = True Then Exit Sub
        'получаем ось опоры
        Dim axisLinePillar As DwgLine = dataPillar.DWGEntity
        If IsNothing(axisLinePillar) = True Then Exit Sub
        If axisLinePillar.Length = 0 Then Exit Sub
        'получаем номер опоры
        Dim numberPillar As Integer = userPillar.Number
        'получаем id сооружения
        Dim idBridge As String = dataPillar.IdStructure
        If IsNothing(idBridge) = True Then Exit Sub
        If idBridge.Trim.Length = 0 Then Exit Sub
        'находим балки для выбранной опоры (0-предыдущий пролет, 1-следующий пролет)
        Dim listBeamsPillar As List(Of Dictionary(Of Integer, StructureElement)) = Pillar.getBeamsPillarByNumber(numberPillar, dictionaryBridgeElements)
        If listBeamsPillar.Count = 0 Then
            MsgBox("Не удалось найти балки для выбранного пролета мостового сооружения!!!")
            Exit Sub
        End If
        ''находим балку с минимальной высотой
        'Dim minElevationBeam As Double = CalculationBeams.getBeamToMinElevation(listBeamsPillar)
        'If minElevationBeam = 999999 Then
        '    MsgBox("Не удалось найти отметку самой нижней балки!!!")
        '    Exit Sub
        'End If
        ''находим отметки насадки
        'userRigel.TopElevation = Math.Round(minElevationBeam - userRigel.MinElevationBeams, 3)
        'userRigel.BottomElevation = Math.Round(userRigel.TopElevation - userRigel.Height, 3)
        ''из массива найденных балок находим крайние (справа и с лева
        'Dim listExtremeBeams As List(Of StructureElement) = CalculationBeams.getExtrmBeamsToPillar(listBeamsPillar)
        ''левые балки смежных пролетов
        'Dim leftDataBeam1 As StructureElement = Nothing
        'Dim leftDataBeam2 As StructureElement = Nothing
        ''правые балки смежных пролетов
        'Dim rightDataBeam1 As StructureElement = Nothing
        'Dim rightDataBeam2 As StructureElement = Nothing
        'If IsNothing(listExtremeBeams.Item(0)) = False Then
        '    leftDataBeam1 = listExtremeBeams.Item(0)
        'End If
        'If IsNothing(listExtremeBeams.Item(2)) = False Then
        '    leftDataBeam2 = listExtremeBeams.Item(2)
        'End If
        'If IsNothing(listExtremeBeams.Item(1)) = False Then
        '    rightDataBeam1 = listExtremeBeams.Item(1)
        'End If
        'If IsNothing(listExtremeBeams.Item(3)) = False Then
        '    rightDataBeam2 = listExtremeBeams.Item(3)
        'End If
        'Dim leftRotationEgeRigel As Double = 0
        'Dim leftAxisBeam1 As DwgLine = leftDataBeam1.DWGEntity
        'Dim leftAxisBeam2 As DwgLine = leftDataBeam2.DWGEntity
        'If leftAxisBeam1.Length > 0 And leftAxisBeam2.Length > 0 Then
        '    leftRotationEgeRigel = (leftAxisBeam1.Rotation + leftAxisBeam2.Rotation) / 2
        'ElseIf leftAxisBeam1.Length > 0 Then
        '    leftRotationEgeRigel = leftAxisBeam1.Rotation
        'ElseIf leftAxisBeam2.Length > 0 Then
        '    leftRotationEgeRigel = leftAxisBeam2.Rotation
        'Else
        '    MsgBox("Не удалось рассчитать направление левой грани ригеля.")
        '    Exit Sub
        'End If
        'Dim rightRotationEgeRigel As Double = 0
        'Dim rightAxisBeam1 As DwgLine = rightDataBeam1.DWGEntity
        'Dim rightAxisBeam2 As DwgLine = rightDataBeam2.DWGEntity
        'If rightAxisBeam1.Length > 0 And rightAxisBeam2.Length > 0 Then
        '    rightRotationEgeRigel = (rightAxisBeam1.Rotation + rightAxisBeam2.Rotation) / 2
        'ElseIf rightAxisBeam1.Length > 0 Then
        '    rightRotationEgeRigel = rightAxisBeam1.Rotation
        'ElseIf rightAxisBeam2.Length > 0 Then
        '    rightRotationEgeRigel = rightAxisBeam2.Rotation
        'Else
        '    MsgBox("Не удалось рассчитать направление правой грани ригеля.")
        '    Exit Sub
        'End If
        ''находим направления короткой стороны насадки
        'userRigel.LeftDirection = Math.Round(leftRotationEgeRigel, 6)
        'userRigel.RightDirection = Math.Round(rightRotationEgeRigel, 6)
        'вычисляем положение Насадки
        Dim boolCalculateRigel As Boolean = userRigel.calculateRigel(axisLinePillar, listBeamsPillar)
        If boolCalculateRigel = False Then
            MsgBox("Не удалось вычислить положение Ригеля!!")
            Exit Sub
        End If
        'вычисляем позиции подферменников
        Dim dictSubFermers As Dictionary(Of Integer, Dictionary(Of Integer, SubFermenters)) = SubFermenters.calculateSubFermenters(Nothing, userRigel, listBeamsPillar, arraySubFerment, True)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем расчет стоек
        Dim dictRack As Dictionary(Of Integer, RackPillar) = Nothing
        If userPillar.PresenceRacks = True Then
            If IsArray(arrayRacks) = True Then
                dictRack = RackPillar.calculateRacks(Nothing, userRigel, userGrillage, arrayRacks, numberPillar, userPillar.ElevationLand, egSurface)
                'вычисляем положение ростверка
                If userPillar.PresencGrillage = True Then
                    Dim boolCalcGrillage As Boolean = userGrillage.calculateGrillage(axisLinePillar, arrayRacks)
                End If
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем положение подготовки
        If userPillar.PresencPreparation = True Then
            Dim boolCalcPreparation As Boolean = userPreparation.calculatePreparation(userGrillage, Nothing)
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вычисляем позицию свай
        Dim dictPile As Dictionary(Of Integer, Dictionary(Of Integer, PilePillar)) = PilePillar.calculatePile(Nothing, userGrillage, arrayPiles, dictRack, userPillar.PileInRack)
    End Sub
    'рисование промежуточной опоры
    Public Sub DrawingMiddlePillar(ByVal dataPillar As StructureElement, userRigel As RigelPillar, ByRef arraySubFerment As SubFermenters(), ByRef arrayRacks As RackPillar(), ByRef userGrillage As GrillagePillar, ByRef userPreparation As PreparationPillar, ByRef arrayPiles As PilePillar(), ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByRef projectSurface As Surface, ByRef egSurface As Surface, ByRef docPileTLC As ConstructionDocument, ByVal templateXML As String)
        If IsNothing(dataPillar) = True Then Exit Sub
        'получаем опору
        Dim userPillar As Pillar = dataPillar.getPillar()
        If IsNothing(userPillar) = True Then Exit Sub
        'получаем id сооружения
        Dim idBridge As String = dataPillar.IdStructure
        If IsNothing(idBridge) = True Then Exit Sub
        If idBridge.Trim.Length = 0 Then Exit Sub
        '===============================================================================================================================
        'рисуем насадку
        If IsNothing(userRigel) = False Then
            If userRigel._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось насадки
                Dim dataRigel As StructureElement = userRigel.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = RigelContour.drawContour(ActivDocument, userRigel, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                If dictionaryCounter.Count > 1 Then
                    If dictionaryCounter.ContainsKey(typeObject.counterRigelTop) = True Then
                        If dictionaryCounter.ContainsKey(typeObject.counterRigelBottom) = True Then
                            Dim topElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterRigelTop)
                            Dim BottomElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterRigelBottom)
                            Dim middlePt1 As Vector3D = userRigel.getPointByCode("middlePt1", True)
                            Dim middlePt2 As Vector3D = userRigel.getPointByCode("middlePt2", True)
                            Dim tempLineRigel As DwgLine = New DwgLine
                            tempLineRigel.StartPoint = middlePt1
                            tempLineRigel.EndPoint = middlePt2
                            Dim drawmodel As Boolean = RigelModel.drawModel(ActivDocument, userRigel, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                        End If
                    End If
                End If

            End If
        End If
        '===============================================================================================================================
        'рисуем подферменники
        If IsNothing(arraySubFerment) = False Then
            If arraySubFerment.Length > 0 Then
                For i As Integer = 0 To arraySubFerment.Length - 1
                    Dim userSubFerm As SubFermenters = arraySubFerment(i)
                    If userSubFerm._elementBridgePoint.ListPointModel.Count > 3 Then
                        'ось подферменника
                        Dim dataSubFerm As StructureElement = userSubFerm.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = SubFermenterContour.drawContour(ActivDocument, userSubFerm, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim topElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterSubFermentersTop)
                        Dim BottomElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.counterSubFermentersBottom)
                        'рисуем модель
                        Dim model As Boolean = SubFermenterModel.drawModel(ActivDocument, userSubFerm, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
        '===============================================================================================================================
        'рисуем стойки
        If IsNothing(arrayRacks) = False Then
            If arrayRacks.Length > 0 Then
                For i As Integer = 0 To arrayRacks.Length - 1
                    Dim userRack As RackPillar = arrayRacks(i)
                    If userRack._elementBridgePoint.ListPointModel.Count > 3 Then
                        'ось стойки
                        Dim dataSubFerm As StructureElement = userRack.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgEntity) = RackContour.drawContour(ActivDocument, userRack, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim model As Boolean = RackModel.drawModel(ActivDocument, userRack, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
        '===============================================================================================================================
        'рисуем ростверк
        If IsNothing(userGrillage) = False Then
            If userGrillage._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataGrillage As StructureElement = userGrillage.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = GrillageContour.drawContour(ActivDocument, userGrillage, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = GrillageModel.drawModel(ActivDocument, userGrillage, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем подготовку
        If IsNothing(userPreparation) = False Then
            If userPreparation._elementBridgePoint.ListPointModel.Count > 3 Then
                'ось левого обратного открылка
                Dim dataPreparation As StructureElement = userPreparation.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                'рисуем контура
                Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PreparationContour.drawContour(ActivDocument, userPreparation, idBridge, dictionaryBridgeElements, templateXML)
                'рисуем модель
                Dim model As Boolean = PreparationModel.drawModel(ActivDocument, userPreparation, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
            End If
        End If
        '===============================================================================================================================
        'рисуем сваи
        If IsNothing(arrayPiles) = False Then
            If arrayPiles.Length > 0 Then
                For i As Integer = 0 To arrayPiles.Length - 1
                    Dim userPile As PilePillar = arrayPiles(i)
                    If userPile.Height > 0 Then
                        'ось стойки
                        Dim dataPile As StructureElement = userPile.drawAxis(ActivDocument, idBridge, templateXML, dictionaryBridgeElements)
                        'рисуем контура
                        Dim drawCounter As Boolean = PileContour.drawContour(ActivDocument, userPile, idBridge, dictionaryBridgeElements, templateXML)
                        'рисуем модель
                        Dim model As Boolean = PileModel.drawModel(ActivDocument, userPile, docPileTLC, idBridge, dictionaryBridgeElements, templateXML)
                    End If
                Next i
            End If
        End If
    End Sub

    '====================================================================================================================
    'удалить все элементы опоры
    Public Sub RemovePillars(ByVal listAxisPillar As List(Of StructureElement), ByRef dictionaryObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)))
        If listAxisPillar.Count > 0 Then
            For i As Integer = 0 To listAxisPillar.Count - 1
                Dim dataPillar As StructureElement = listAxisPillar.Item(i)
                If IsNothing(dataPillar.DWGEntity) = False Then
                    Dim axisPillar As DwgLine = dataPillar.DWGEntity
                    Dim numberPillar As Integer = FuncGSON.getValue(dataPillar.KeyParameter, "Number")
                    If numberPillar > 0 Then
                        For j As Integer = 0 To dictionaryObjectBridge.Count - 1
                            Dim typeObject As StructureElement.typeObject = dictionaryObjectBridge.ElementAt(j).Key
                            Dim valueObject As List(Of StructureElement) = dictionaryObjectBridge.ElementAt(j).Value
                            If typeObject > 119 And typeObject < 182 Then
                                For k As Integer = 0 To valueObject.Count - 1
                                    Dim dataElement As StructureElement = valueObject.Item(k)
                                    Dim keyParamElement As String = dataElement.KeyParameter
                                    If FuncGSON.IsValidJson(keyParamElement) = True Then
                                        Dim numberPillarElement As Integer = FuncGSON.getValue(keyParamElement, "NumberPillar")
                                        If numberPillar = numberPillarElement Then
                                            If IsNothing(dataElement.DWGEntity) = False Then
                                                Dim entity As DwgEntity = dataElement.DWGEntity
                                                Dim activProjectDocument As Drawing = entity.Drawing
                                                ActivDocument.ActiveSpace.Entities.Remove(entity)
                                            End If
                                        End If
                                    End If
                                Next k
                            End If
                        Next j
                    End If
                End If
            Next i
        End If
    End Sub

End Class
