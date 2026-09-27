Imports System.ComponentModel
Imports CivilEnginStructures.BeamI
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Operation.Buffer
Imports Newtonsoft.Json.Linq
Imports Topomatic.Alg
Imports Topomatic.Alg.Road.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Public Class Bridges
    'Inherits StructureElement
    ' Перечисление для типов размещения балок
    Public Enum typePlacementBeam
        fixed                     ' Фиксированные размеры балок
        float                     ' Балки индивидуального проектирования
        maxClearence              ' Расчетный максимальный зазор
    End Enum
    ' Поля класса
    Private _name As String                                         ' Имя сооружения
    Private _typeBridge As typePlacementBeam                        ' Тип путепровода (0-фиксированные, 1-индивидуальные, 2-макс. зазор)
    Private _countProlet As Integer                                 ' Число пролетов
    Private _centreAxisBeam As Boolean                              ' Наличие осевого ряда балок
    Private _countLeftRows As Integer                               ' Число рядов балок влево от оси
    Private _countRightRows As Integer                              ' Число рядов балок вправо от оси
    Private _dimLeftStructure As Double                             ' Ширина сооружения влево от оси
    Private _dimRightStructure As Double                            ' Ширина сооружения вправо от оси
    Private _startPlacementPosition As Double                       ' Пикет начальной раскладки мостового сооружения (не зависит от смещения)
    Private _offsetHPosition As Double                              ' Горизонтальное смещение вдоль оси
    Private _offsetHTPosition As Double                             ' Горизонтальное смещение поперек оси
    Private _offsetVPosition As Double                              ' Вертикальное смещение
    Private _nameSurface As String                                  ' Имя проектной поверхности
    Private _nameEgSurface As String                                ' Имя поверхности земли
    Private _nameAlignment As String                                ' Имя трассы
    Public _elementBridgePoint As PointsCollections
    ' Конструктор по умолчанию
    Public Sub New()
        _name = ""
        _typeBridge = typePlacementBeam.fixed
        _countProlet = 0
        _centreAxisBeam = True
        _countRightRows = 0
        _countLeftRows = 0
        _dimLeftStructure = 0
        _dimRightStructure = 0
        _startPlacementPosition = 0
        _offsetHPosition = 0
        _offsetVPosition = 0
        _offsetHTPosition = 0
        _nameSurface = ""
        _nameAlignment = ""
        _nameEgSurface = ""
        '_axisEntity = New DwgPolyline
    End Sub

    <Browsable(True)>
    <Description("Наименование сооружения")>
    <Category("Свойства")>
    <DisplayName("Наименование сооружения")>
    Public Property NameBridge() As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Способ раскладки балок (фиксированное, с произвольной длиной балки, с использованием минимального и максимального зазоров)")>
    <Category("Свойства")>
    <DisplayName("Способ раскладки балок")>
    Public Property TypeBridge() As typePlacementBeam
        Get
            Return _typeBridge
        End Get
        Set(value As typePlacementBeam)
            _typeBridge = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Число пролетов сооружения")>
    <Category("Свойства")>
    <DisplayName("Число пролетов")>
    Public Property ProletCount() As Integer
        Get
            Return _countProlet
        End Get
        Set(value As Integer)
            _countProlet = Math.Abs(value)
        End Set
    End Property

    <Browsable(True)>
    <Description("Наличие осевого ряда")>
    <Category("Свойства")>
    <DisplayName("Наличие осевого ряда")>
    Public Property centerAxis() As Boolean
        Get
            Return _centreAxisBeam
        End Get
        Set(value As Boolean)
            _centreAxisBeam = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Чмсло рядов балок слева")>
    <Category("Свойства")>
    <DisplayName("Чмсло рядов слева")>
    Public Property LeftRowsCount() As Integer
        Get
            Return _countLeftRows
        End Get
        Set(value As Integer)
            _countLeftRows = Math.Abs(value)
        End Set
    End Property

    <Browsable(True)>
    <Description("Чмсло рядов балок справа")>
    <Category("Свойства")>
    <DisplayName("Чмсло рядов справа")>
    Public Property RightRowsCount() As Integer
        Get
            Return _countRightRows
        End Get
        Set(value As Integer)
            _countRightRows = Math.Abs(value)
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина сооружения слева от оси, м")>
    <Category("Свойства")>
    <DisplayName("Ширина сооружения слева")>
    Public Property LeftStructureWidth() As Double
        Get
            Return _dimLeftStructure
        End Get
        Set(value As Double)
            _dimLeftStructure = Math.Abs(value)
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина сооружения справа от оси, м")>
    <Category("Свойства")>
    <DisplayName("Ширина сооружения справа")>
    Public Property RightStructureWidth() As Double
        Get
            Return _dimRightStructure
        End Get
        Set(value As Double)
            _dimRightStructure = Math.Abs(value)
        End Set
    End Property

    <Browsable(True)>
    <Description("Начальный пикет раскладки балок, ПК+")>
    <Category("Свойства")>
    <DisplayName("Начальный пикет")>
    Public Property startPlacementPosition() As Double
        Get
            Return _startPlacementPosition
        End Get
        Set(value As Double)
            _startPlacementPosition = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Смещение сооружения по ходу пикетажа, относительно начального пикета раскладки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение вдоль оси")>
    Public Property HorizontalOffset() As Double
        Get
            Return _offsetHPosition
        End Get
        Set(value As Double)
            _offsetHPosition = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Смещение сооружения влево\право, относительно начального пикета раскладки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение поперек оси")>
    Public Property TransverseOffset() As Double
        Get
            Return _offsetHTPosition
        End Get
        Set(value As Double)
            _offsetHTPosition = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Смещение по вертикали относительно проектной поверхности, м")>
    <Category("Свойства")>
    <DisplayName("Вертикальное смещение")>
    Public Property VerticalOffset() As Double
        Get
            Return _offsetVPosition
        End Get
        Set(value As Double)
            _offsetVPosition = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Имя проектной поверхности")>
    <Category("Свойства")>
    <DisplayName("Проектная поверхность")>
    Public Property projectSurfaceName() As String
        Get
            Return _nameSurface
        End Get
        Set(value As String)
            _nameSurface = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Имя поверхности существующей земли")>
    <Category("Свойства")>
    <DisplayName("Фактическая поверхность")>
    Public Property EarthSurfaceName() As String
        Get
            Return _nameEgSurface
        End Get
        Set(value As String)
            _nameEgSurface = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Имя проектной трассы")>
    <Category("Свойства")>
    <DisplayName("Ось трассы")>
    Public Property AlignmentName() As String
        Get
            Return _nameAlignment
        End Get
        Set(value As String)
            _nameAlignment = value
        End Set
    End Property

    Public Shared Function createBridge() As StructureElement
        Dim elementBridge As StructureElement = New StructureElement()
        elementBridge.Label = "Мосты и путепроводы"
        elementBridge.ClassBridgeObject = StructureElement.classBridge.OtherElements
        elementBridge.ClassObject = StructureElement.classStructure.Bridges
        elementBridge.Name = StructureElement.typeObject.axisBridge
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.axisBridge)
        elementBridge.Description = deskObject
        Dim userBridge As Bridges = New Bridges
        Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBridge)
        elementBridge.KeyParameter = keyParam
        elementBridge.IdElement = Guid.NewGuid.ToString
        elementBridge.IdStructure = Guid.NewGuid.ToString
        elementBridge.Note = ""
        elementBridge.DWGEntity = Nothing
        Return elementBridge
    End Function
    '======================================================================================================================================
    'получить все элементы мостового сооружения 
    Public Function getBridgeObjects(ByVal axisBridge As DwgEntity) As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))
        Dim result As New Dictionary(Of StructureElement.typeObject, List(Of StructureElement))
        If IsNothing(axisBridge) = False Then
            Dim structData As StructureElement = New StructureElement
            Dim boolReadData As Boolean = FuncXRecords.getXRecords(axisBridge, structData)
            If boolReadData = True Then
                If IsNothing(structData) = False Then
                    Dim idBridge As String = structData.IdStructure
                    If IsNothing(idBridge) = True Then Return result
                    If idBridge.Trim.Length = 0 Then Return result
                    Dim activeDocument As Topomatic.Dwg.Drawing = axisBridge.Drawing
                    If IsNothing(activeDocument) = False Then
                        For Each entity As DwgEntity In activeDocument.ActiveSpace.Entities
                            Dim xdataRecord As StructureElement = New StructureElement()
                            Dim boolFindElementInfo As Boolean = FuncXRecords.getXRecords(entity, xdataRecord)
                            If boolFindElementInfo = True Then
                                If IsNothing(xdataRecord) = True Then Continue For
                                Dim idStructureElement As String = xdataRecord.IdStructure
                                If IsNothing(idStructureElement) = True Then Continue For
                                If idStructureElement.Trim.Length = 0 Then Continue For
                                If idBridge Like idStructureElement Then
                                    AddToDictionary(result, xdataRecord)
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        Return result
    End Function
    Private Sub AddToDictionary(ByRef dictionary As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)),
                           ByVal elementInfo As StructureElement)
        If Not dictionary.ContainsKey(elementInfo.Name) Then
            dictionary.Add(elementInfo.Name, New List(Of StructureElement))
        End If
        dictionary(elementInfo.Name).Add(elementInfo)
    End Sub
    '====================================================================================================================================
    'функция ищет оси опор и оси опирания балок
    Public Function getPillars(ByVal dictBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional idBridge As String = "") As Dictionary(Of Integer, List(Of StructureElement))
        Dim result As Dictionary(Of Integer, List(Of StructureElement)) = New Dictionary(Of Integer, List(Of StructureElement))
        Dim countBridgeProlet As Integer = _countProlet
        If countBridgeProlet > 0 Then
            For i As Integer = 1 To countBridgeProlet + 1
                Dim dataPrevAxisBeamsPillar As StructureElement = Nothing
                Dim dataAxisPillar As StructureElement = Nothing
                Dim dataNextAxisBeamsPillar As StructureElement = Nothing
                If i = 1 Or i = countBridgeProlet + 1 Then
                    dataAxisPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.LastPillar)
                    Dim axisPillar As Pillar = New Pillar()
                    axisPillar.Number = i
                    Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(axisPillar)
                    dataAxisPillar.KeyParameter = keyParam
                Else
                    dataPrevAxisBeamsPillar = AxisBeamsPillars.createAxis(idBridge)
                    Dim prevAxisPillar As AxisBeamsPillars = New AxisBeamsPillars()
                    prevAxisPillar.numberPillar = i
                    prevAxisPillar.numberProlet = i - 1
                    Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(prevAxisPillar)
                    dataPrevAxisBeamsPillar.KeyParameter = keyParam

                    dataAxisPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.LastPillar)
                    Dim axisPillar As Pillar = New Pillar()
                    axisPillar.Number = i
                    keyParam = Newtonsoft.Json.JsonConvert.SerializeObject(axisPillar)
                    dataAxisPillar.KeyParameter = keyParam

                    dataNextAxisBeamsPillar = AxisBeamsPillars.createAxis(idBridge)
                    Dim nextAxisPillar As AxisBeamsPillars = New AxisBeamsPillars()
                    nextAxisPillar.numberPillar = i
                    nextAxisPillar.numberProlet = i
                    keyParam = Newtonsoft.Json.JsonConvert.SerializeObject(nextAxisPillar)
                    dataNextAxisBeamsPillar.KeyParameter = keyParam
                End If
                Dim listAxisPillar As List(Of StructureElement) = New List(Of StructureElement) From {dataPrevAxisBeamsPillar, dataAxisPillar, dataNextAxisBeamsPillar}
                result.Add(i, listAxisPillar)
            Next i
        End If
        'ищем существующие оси опор
        If IsNothing(dictBridgeElements) = False Then
            If dictBridgeElements.Count > 0 Then
                If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
                    Dim listPillar As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisPillar)
                    If listPillar.Count > 0 Then
                        For i As Integer = 0 To listPillar.Count - 1
                            Dim userPillar As StructureElement = listPillar.Item(i)
                            Dim keyParameter As String = userPillar.KeyParameter
                            If FuncGSON.IsValidJson(keyParameter) = True Then
                                Dim numberPillar As Integer = FuncGSON.getValue(keyParameter, "Number")
                                If numberPillar > 0 And numberPillar <= result.Count Then
                                    Dim listAxisPillar As List(Of StructureElement) = result.Item(numberPillar)
                                    listAxisPillar.Item(1) = userPillar
                                    result.Item(numberPillar) = listAxisPillar
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        End If
        'ищем существующие оси опирания балок
        If IsNothing(dictBridgeElements) = False Then
            If dictBridgeElements.Count > 0 Then
                If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisPillarBeams) = True Then
                    Dim listPillar As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisPillarBeams)
                    If listPillar.Count > 0 Then
                        For i As Integer = 0 To listPillar.Count - 1
                            Dim userBeambPillar As StructureElement = listPillar.Item(i)
                            Dim keyParameter As String = userBeambPillar.KeyParameter
                            If FuncGSON.IsValidJson(keyParameter) = True Then
                                Dim numberPillar As Integer = Val(FuncGSON.getValue(keyParameter, "numberPillar"))
                                If numberPillar > 1 And numberPillar < result.Count Then
                                    Dim listAxisPillar As List(Of StructureElement) = result.Item(numberPillar)
                                    Dim numberBeamsPillar As Integer = FuncGSON.getValue(keyParameter, "numberProlet")
                                    If numberBeamsPillar < numberPillar Then
                                        listAxisPillar.Item(0) = userBeambPillar
                                        result.Item(numberPillar) = listAxisPillar
                                    Else
                                        listAxisPillar.Item(2) = userBeambPillar
                                        result.Item(numberPillar) = listAxisPillar
                                    End If
                                End If
                            End If

                        Next i
                    End If
                End If
            End If
        End If
        Return result
    End Function
    '====================================================================================================================================
    'функция получает оси опор
    Public Function getAxisPillars(ByVal dictBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional middlePillar As Boolean = False) As Dictionary(Of Integer, StructureElement)
        Dim result As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        Dim countBridgeProlet As Integer = ProletCount
        If countBridgeProlet > 1 Then
            'ищем существующие оси опор
            If IsNothing(dictBridgeElements) = False Then
                If dictBridgeElements.Count > 0 Then
                    If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
                        Dim listPillar As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisPillar)
                        If listPillar.Count > 0 Then
                            For i As Integer = 0 To listPillar.Count - 1
                                Dim dataPillar As StructureElement = listPillar.Item(i)
                                Dim keyParameter As String = dataPillar.KeyParameter
                                If FuncGSON.IsValidJson(keyParameter) = True Then
                                    Dim jsonObj As JObject = JObject.Parse(keyParameter)
                                    Dim numberPillar As Integer = Val(FuncGSON.getValue(keyParameter, "number"))
                                    If numberPillar > 0 Then
                                        If numberPillar = 1 And middlePillar = True Then
                                            Continue For
                                        ElseIf numberPillar = countBridgeProlet + 1 And middlePillar = True Then
                                            Continue For
                                        End If
                                        If result.ContainsKey(numberPillar) = False Then
                                            result.Add(numberPillar, dataPillar)
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
        If result.Count > 1 Then
            result = result.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return result
    End Function
    '====================================================================================================================================
    'функция ищет все балки в сооружении
    Public Function getBeams(ByVal dictBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        Dim result As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        Dim countBridgeProlet As Integer = _countProlet
        Dim countBridgeLeft As Integer = _countLeftRows
        Dim countBridgeRight As Integer = _countRightRows
        Dim boolCenterAxis As Boolean = _centreAxisBeam
        If countBridgeProlet > 0 Then
            For i As Integer = 1 To countBridgeProlet
                Dim dictBeamsRow As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
                If countBridgeLeft > 0 Then
                    For j As Integer = countBridgeLeft To 1 Step -1
                        Dim numberRow As Integer = -1 * j
                        dictBeamsRow.Add(numberRow, New StructureElement)
                    Next j
                End If
                If boolCenterAxis = True Then
                    dictBeamsRow.Add(0, New StructureElement)
                End If
                If countBridgeRight > 0 Then
                    For j As Integer = 1 To countBridgeRight
                        dictBeamsRow.Add(j, New StructureElement)
                    Next j
                End If
                result.Add(i, dictBeamsRow)
            Next i
        End If
        If result.Count > 0 Then
            'ищем существующие оси опор
            If IsNothing(dictBridgeElements) = False Then
                If dictBridgeElements.Count > 0 Then
                    If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
                        Dim listBeams As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisBeam)
                        If listBeams.Count > 0 Then
                            For i As Integer = 0 To listBeams.Count - 1
                                Dim userBeam As StructureElement = listBeams.Item(i)
                                Dim keyParameter As String = userBeam.KeyParameter
                                If FuncGSON.IsValidJson(keyParameter) = True Then
                                    Dim jsonObj As JObject = JObject.Parse(keyParameter)
                                    If jsonObj("numberProlet") IsNot Nothing Then
                                        Dim numberProlet As Integer = jsonObj("numberProlet").Value(Of Integer)()
                                        If jsonObj("numberRow") IsNot Nothing Then
                                            Dim numberRow As Integer = jsonObj("numberRow").Value(Of Integer)()
                                            If numberProlet > 0 And numberProlet <= result.Count Then
                                                If result.ContainsKey(numberProlet) = True Then
                                                    Dim listBeamsProlet As Dictionary(Of Integer, StructureElement) = result.Item(numberProlet)
                                                    If listBeamsProlet.ContainsKey(numberRow) = True Then
                                                        listBeamsProlet.Item(numberRow) = userBeam
                                                        result.Item(numberProlet) = listBeamsProlet
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function
    '====================================================================================================================================
    'функция ищет все траектории раскладки балок в сооружении
    Public Function getTrajectoryPlacementBeams(ByVal dictBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Dictionary(Of Integer, StructureElement)
        Dim result As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        'ищем существующие оси опор
        If IsNothing(dictBridgeElements) = False Then
            If dictBridgeElements.Count > 0 Then
                If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisTrajectoryPlacementBeams) = True Then
                    Dim listBeams As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisTrajectoryPlacementBeams)
                    If listBeams.Count > 0 Then
                        For i As Integer = 0 To listBeams.Count - 1
                            Dim dataElement As StructureElement = listBeams.Item(i)
                            Dim keyParameter As String = dataElement.KeyParameter
                            If FuncGSON.IsValidJson(keyParameter) = True Then
                                Dim jsonObj As JObject = JObject.Parse(keyParameter)
                                If jsonObj("numberRows") IsNot Nothing Then
                                    Dim numberRow As Integer = jsonObj("numberRows").Value(Of Integer)()
                                    If result.ContainsKey(numberRow) Then
                                        result.Add(numberRow, dataElement)
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        End If
        Return result
    End Function
    '====================================================================================================================================
    'функция ищет все участки омоноличивания балок
    Public Function getSitesMonolitBeams(ByVal dictBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        Dim result As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        If IsNothing(dictBridgeElements) = True Then Return result
        If dictBridgeElements.Count = 0 Then Return result
        Dim countBridgeProlet As Integer = _countProlet
        Dim countBridgeLeft As Integer = _countLeftRows
        Dim countBridgeRight As Integer = _countRightRows
        Dim boolCenterAxis As Boolean = _centreAxisBeam
        If countBridgeProlet > 0 Then
            For i As Integer = 1 To countBridgeProlet
                Dim dictSitesonolitBeams As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
                If countBridgeLeft > 1 Then
                    For j As Integer = countBridgeLeft To 2 Step -1
                        Dim numberRow As Integer = -1 * j
                        dictSitesonolitBeams.Add(numberRow, New StructureElement)
                    Next j
                End If
                If boolCenterAxis = True Then
                    dictSitesonolitBeams.Add(0, New StructureElement)
                End If
                If countBridgeRight > 1 Then
                    For j As Integer = 2 To countBridgeRight
                        dictSitesonolitBeams.Add(j, New StructureElement)
                    Next j
                End If
                result.Add(i, dictSitesonolitBeams)
            Next i
        End If
        If result.Count > 0 Then
            If IsNothing(dictBridgeElements) = False Then
                If dictBridgeElements.Count > 0 Then
                    If dictBridgeElements.ContainsKey(StructureElement.typeObject.axisSiteMonolitBeams) = True Then
                        Dim listSitesMonolit As List(Of StructureElement) = dictBridgeElements.Item(StructureElement.typeObject.axisSiteMonolitBeams)
                        If listSitesMonolit.Count > 0 Then
                            For i As Integer = 0 To listSitesMonolit.Count - 1
                                Dim userSiteMonolit As StructureElement = listSitesMonolit.Item(i)
                                Dim keyParameter As String = userSiteMonolit.KeyParameter
                                If FuncGSON.IsValidJson(keyParameter) = True Then
                                    Dim jsonObj As JObject = JObject.Parse(keyParameter)
                                    If jsonObj("numberProlet") IsNot Nothing Then
                                        Dim numberProlet As Integer = jsonObj("numberProlet").Value(Of Integer)()
                                        If jsonObj("numberRow") IsNot Nothing Then
                                            Dim numberRow As Integer = jsonObj("numberRow").Value(Of Integer)()
                                            If numberProlet > 0 And numberProlet <= result.Count Then
                                                If result.ContainsKey(numberProlet) = True Then
                                                    Dim listSitesNonolitBeams As Dictionary(Of Integer, StructureElement) = result.Item(numberProlet)
                                                    If listSitesNonolitBeams.ContainsKey(numberRow) = True Then
                                                        listSitesNonolitBeams.Item(numberRow) = userSiteMonolit
                                                        result.Item(numberProlet) = listSitesNonolitBeams
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция возвращает минимальный зазор (не рабочая)
    Public Function zazorPreviousBeam2(ByVal prevLineShortBeam As DwgLine, ByVal lineShortBeam As DwgLine, ByRef listZazor As List(Of Double)) As Double
        zazorPreviousBeam2 = -1
        If IsNothing(lineShortBeam) = True Then Exit Function
        If IsNothing(prevLineShortBeam) = True Then Exit Function
        Dim dataPrevBeam As StructureElement = Nothing
        Dim userPrevBeam As BeamI = Nothing
        Dim boolReadData As Boolean = FuncXRecords.getXRecords(prevLineShortBeam, dataPrevBeam)
        If boolReadData = True Then
            userPrevBeam = dataPrevBeam.getBeamI()
        End If
        If IsNothing(userPrevBeam) = True Then Return -1

        Dim dataBeam As StructureElement = Nothing
        Dim userBeam As BeamI = Nothing
        boolReadData = FuncXRecords.getXRecords(lineShortBeam, dataBeam)
        If boolReadData = True Then
            userBeam = dataBeam.getBeamI()
        End If
        If IsNothing(userBeam) = True Then Return -1

        'получаем сечения предыдущей балки
        Dim startSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionPrevBeam As Boolean = CalculationBeams.restoreElementsBeam(prevLineShortBeam, userBeam, startSectionPrevBeam, endSectionPrevBeam, CalculationBeams.rectoreBeam.fullBeam)
        'получаем сечения балки
        Dim startSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionBeam As Boolean = CalculationBeams.restoreElementsBeam(lineShortBeam, userBeam, startSectionBeam, endSectionBeam, CalculationBeams.rectoreBeam.fullBeam)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'приводим балки к одному уровню по низу
        'верх балки предыдущей
        Dim topLeftPrevBeam1 As Vector3D = startSectionPrevBeam.Item(0) 'лево
        Dim topRightPrevBeam1 As Vector3D = startSectionPrevBeam.Item(1) 'право
        'низ балки предыдущей
        Dim downLeftPrevBeam1 As Vector3D = startSectionPrevBeam.Item(2) 'лево
        Dim downRightPrevBeam1 As Vector3D = startSectionPrevBeam.Item(3) 'право
        'верх балки предыдущей
        Dim topLeftPrevBeam2 As Vector3D = endSectionPrevBeam.Item(0) 'лево
        Dim topRightPrevBeam2 As Vector3D = endSectionPrevBeam.Item(1) 'право
        'низ балки предыдущей
        Dim downLeftPrevBeam2 As Vector3D = endSectionPrevBeam.Item(2) 'лево
        Dim downRightPrevBeam2 As Vector3D = endSectionPrevBeam.Item(3) 'право
        'высота по низу предыдущей балки

        'верх определяемая балка
        Dim topLeftBeam1 As Vector3D = startSectionBeam.Item(0) 'лево
        Dim topRightBeam1 As Vector3D = startSectionBeam.Item(1) 'право
        'низ балки определяемой
        Dim downLeftBeam1 As Vector3D = startSectionBeam.Item(2) 'лево
        Dim downRightBeam1 As Vector3D = startSectionBeam.Item(3) 'право
        'верх определяемая балка
        Dim topLeftBeam2 As Vector3D = endSectionBeam.Item(0) 'лево
        Dim topRightBeam2 As Vector3D = endSectionBeam.Item(1) 'право
        'низ балки определяемой
        Dim downLeftBeam2 As Vector3D = endSectionBeam.Item(2) 'лево
        Dim downRightBeam2 As Vector3D = endSectionBeam.Item(3) 'право

        '=======================================================================================================================
        'проверяем пересечения верх лево
        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftBeam1, topRightBeam1,
                                                                                topLeftPrevBeam2, New Vector3D(topLeftPrevBeam1.Pos, topLeftPrevBeam2.Z)))

        '=======================================================================================================================
        'проверяем пересечения верх право
        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftBeam1, topRightBeam1,
                                                                                topRightPrevBeam2, New Vector3D(topRightPrevBeam1.Pos, topRightPrevBeam2.Z)))

        '=======================================================================================================================
        'проверяем пересечения низ право
        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftBeam1, topRightBeam1,
                                                                                downRightPrevBeam2, New Vector3D(downRightPrevBeam1.Pos, downRightPrevBeam2.Z)))

        '=======================================================================================================================
        'проверяем пересечения низ лево
        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftBeam1, topRightBeam1,
                                                                                downLeftPrevBeam2, New Vector3D(downLeftPrevBeam1.Pos, downLeftPrevBeam2.Z)))
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем проверку с другой стороны
        'проверяем пересечения верх лево
        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftPrevBeam2, topRightPrevBeam2,
                                                                                topLeftBeam1, New Vector3D(topLeftBeam2.Pos, topLeftBeam1.Z)))

        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftPrevBeam2, topRightPrevBeam2,
                                                                                topRightBeam1, New Vector3D(topRightBeam2.Pos, topRightBeam1.Z)))

        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftPrevBeam2, topRightPrevBeam2,
                                                                                downRightBeam1, New Vector3D(downRightBeam2.Pos, downRightBeam1.Z)))

        listZazor.Add(MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(topLeftPrevBeam2, topRightPrevBeam2,
                                                                                downLeftBeam1, New Vector3D(downLeftBeam2.Pos, downLeftBeam1.Z)))
        Return listZazor.Min
    End Function
    '=========================================================================================================
    'функция возвращает минимальный зазор
    Public Function zazorPreviousBeam(ByVal prevLineShortBeam As DwgLine, ByVal lineShortBeam As DwgLine, ByRef listZazor As List(Of Double), Optional userDataBeam As BeamI = Nothing) As Double
        zazorPreviousBeam = -1
        If IsNothing(lineShortBeam) = True Then Exit Function
        If IsNothing(prevLineShortBeam) = True Then Exit Function
        Dim dataPrevBeam As StructureElement = Nothing
        Dim userPrevBeam As BeamI = Nothing
        Dim boolReadData As Boolean = FuncXRecords.getXRecords(prevLineShortBeam, dataPrevBeam)
        If boolReadData = True Then
            userPrevBeam = dataPrevBeam.getBeamI()
        End If
        If IsNothing(userPrevBeam) = True Then Return -1

        Dim userBeam As BeamI = userDataBeam
        If IsNothing(userBeam) = True Then
            Dim dataBeam As StructureElement = Nothing
            boolReadData = FuncXRecords.getXRecords(lineShortBeam, dataBeam)
            If boolReadData = True Then
                userBeam = dataBeam.getBeamI()
            End If
        End If
        If IsNothing(userBeam) = True Then Return -1

        'получаем сечения предыдущей балки
        Dim startSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionPrevBeam As Boolean = CalculationBeams.restoreElementsBeam(prevLineShortBeam, userBeam, startSectionPrevBeam, endSectionPrevBeam, CalculationBeams.rectoreBeam.fullBeam)
        'получаем сечения балки
        Dim startSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionBeam As Boolean = CalculationBeams.restoreElementsBeam(lineShortBeam, userBeam, startSectionBeam, endSectionBeam, CalculationBeams.rectoreBeam.fullBeam)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'приводим балки к одному уровню по низу
        'верх балки предыдущей
        If boolSectionBeam = True Then
            Dim topLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(0) 'лево
            Dim topRightPrevBeam As Vector3D = endSectionPrevBeam.Item(1) 'право
            'низ балки предыдущей
            Dim downLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(2) 'лево
            Dim downRightPrevBeam As Vector3D = endSectionPrevBeam.Item(3) 'право
            'высота по низу предыдущей балки
            Dim elevPrevBeam As Double = (downLeftPrevBeam.Z + downRightPrevBeam.Z) / 2
            'верх определяемая балка
            Dim topLeftBeam As Vector3D = startSectionBeam.Item(0) 'лево
            Dim topRightBeam As Vector3D = startSectionBeam.Item(1) 'право
            'низ балки определяемой
            Dim downLeftBeam As Vector3D = startSectionBeam.Item(2) 'лево
            Dim downRightBeam As Vector3D = startSectionBeam.Item(3) 'право
            'высота по низу предыдущей балки
            Dim elevBeam As Double = (downLeftBeam.Z + downRightBeam.Z) / 2
            'дельта по высоте
            Dim deltaH As Double = elevPrevBeam - elevBeam
            If Math.Abs(deltaH) > 0.01 Then
                Dim oldLenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length
                If deltaH < 0 Then
                    'предудущая балка ниже чем расчетная (приводим ее уровню расчетной бплки
                    downLeftPrevBeam = MathFunction.FuncCalcPointInLine(downLeftPrevBeam, topLeftPrevBeam, Math.Abs(deltaH), 3)
                    downRightPrevBeam = MathFunction.FuncCalcPointInLine(downRightPrevBeam, topRightPrevBeam, Math.Abs(deltaH), 3)
                    Dim lenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length - oldLenghtDownWidthBeam
                    If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                        Dim dLen As Double = lenghtDownWidthBeam / 2
                        Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftPrevBeam, downRightPrevBeam, -1 * dLen, -1 * dLen, 3)
                    End If
                Else 'расчетная балка ниже чем предыдущая (подтягиваем расчетную балку к предыдущей)
                    downLeftBeam = MathFunction.FuncCalcPointInLine(downLeftBeam, topLeftBeam, deltaH)
                    downRightBeam = MathFunction.FuncCalcPointInLine(downRightBeam, topRightBeam, deltaH)
                    Dim lenghtDownWidthBeam As Double = (downLeftBeam - downRightBeam).Length - oldLenghtDownWidthBeam
                    If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                        Dim dLen As Double = lenghtDownWidthBeam / 2
                        Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftBeam, downRightBeam, -1 * dLen, -1 * dLen, 3)
                    End If
                End If
            End If

            '====================================================================================================================
            'анализируем зазор по верху
            '====================================================================================================================
            'проверяем пересечение верха балки с горизонтальной линией предыдущей балки
            'верх лево
            Dim topLeftZazor As Double = 99999999
            Dim intersectPoint As Vector2D = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(0).Pos, topLeftBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topLeftBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - topLeftPrevBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - topLeftPrevBeam.Pos).Length
                        If tempZazor < topLeftZazor Then
                            topLeftZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предыдущей балки и смотрим пересечение
            If topLeftZazor = 99999999 Then
                Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
                Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'верх право
            Dim topRightZazor As Double = 99999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(1).Pos, topRightBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topRightBeam.Pos).Length
                If tempZazor < topRightZazor Then
                    topRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                    If tempZazor < topRightZazor Then
                        topRightZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - topRightPrevBeam.Pos).Length
                    If tempZazor < topRightZazor Then
                        topRightZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - topRightPrevBeam.Pos).Length
                        If tempZazor < topRightZazor Then
                            topRightZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If topRightZazor = 99999999 Then
                Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
                Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topRightZazor = tempZazor
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'низ лево
            Dim downLeftZazor As Double = 99999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(2).Pos, downLeftBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downLeftBeam.Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - downLeftPrevBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - downLeftPrevBeam.Pos).Length
                        If tempZazor < downLeftZazor Then
                            downLeftZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If downLeftZazor = 99999999 Then
                Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
                Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                End If
            End If

            'низ право
            Dim downRightZazor As Double = 99999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(3).Pos, downRightBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downRightBeam.Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - downRightPrevBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - downRightPrevBeam.Pos).Length
                        If tempZazor < downRightZazor Then
                            downRightZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If downRightZazor = 99999999 Then
                Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
                Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                End If
            End If
            listZazor.Add(topLeftZazor)
            listZazor.Add(topRightZazor)
            listZazor.Add(downLeftZazor)
            listZazor.Add(downRightZazor)
            Return listZazor.Min
        End If
    End Function
    'функция вычисляет середину между смежными балками
    Public Shared Function calculateMiddlePointBeams(ByVal prevLineShortBeam As DwgLine, ByVal lineShortBeam As DwgLine) As Vector3D
        calculateMiddlePointBeams = Nothing
        If IsNothing(lineShortBeam) = True Then Exit Function
        If IsNothing(prevLineShortBeam) = True Then Exit Function
        Dim userPrevBeam As BeamI = Nothing
        Dim dataBeamPrev As StructureElement = Nothing
        Dim boolFindRec As Boolean = FuncXRecords.getXRecords(prevLineShortBeam, dataBeamPrev)
        If boolFindRec = True Then
            userPrevBeam = dataBeamPrev.getBeamI()
        End If
        Dim userBeam As BeamI = Nothing
        Dim dataBeam As StructureElement = Nothing
        Dim boolFindBeamRec As Boolean = FuncXRecords.getXRecords(lineShortBeam, dataBeam)
        If boolFindBeamRec = True Then
            userBeam = dataBeam.getBeamI()
        End If
        If IsNothing(userBeam) = True Then Return Nothing
        If IsNothing(userPrevBeam) = True Then Return Nothing
        Dim startPointPrevBeam As Vector3D = New Vector3D()
        Dim endPointPrevBeam As Vector3D = New Vector3D()
        Dim boolFindPoint As Boolean = MathFunction.FuncVirtualExtendLine(prevLineShortBeam, 0, userPrevBeam.b, startPointPrevBeam, endPointPrevBeam)

        Dim startPointBeam As Vector3D = New Vector3D()
        Dim endPointBeam As Vector3D = New Vector3D()
        boolFindPoint = MathFunction.FuncVirtualExtendLine(lineShortBeam, userBeam.a, 0, startPointBeam, endPointBeam)

        Dim middlePoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(endPointPrevBeam, startPointBeam)
        Return middlePoint
    End Function
    '=========================================================================================================
    'функция оформляет и создает границу мостового сооружения и ось трассы автодороги
    Public Function drawAxisAndBoundaryBridge(ByRef axisBridge As DwgPolyline, ByVal dictionaryBridgeBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal align As Alignment, Optional ByRef boundStructure As DwgPolyline = Nothing) As Boolean
        drawAxisAndBoundaryBridge = False
        If _dimLeftStructure = 0 Then Return False
        If _dimRightStructure = 0 Then Return False
        If IsNothing(axisBridge) = True Then Return False
        If IsNothing(dictionaryBridgeBeams) = True Then Return False
        If dictionaryBridgeBeams.Count = 0 Then Return False
        Dim ActivDocument As Topomatic.Dwg.Drawing = axisBridge.Drawing
        If IsNothing(ActivDocument) = True Then Return False
        'граница габарита моста слева
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.getPolylineByAlignment(ActivDocument, align, -1 * _dimLeftStructure + _offsetHTPosition)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)
        'граница габарита моста справа
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.getPolylineByAlignment(ActivDocument, align, -1 * _dimRightStructure + _offsetHTPosition, True)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)

        Dim axisCentrePline As DwgPolyline = FuncAlignment.getPolylineByAlignment(ActivDocument, align, _offsetHTPosition)
        Dim axisCentrePline3D As IPolyline3D = New Polyline3D()
        axisCentrePline.GetPolyline(axisCentrePline3D)
        Try
            Dim firstProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.First.Value
            Dim startFirtPoint As Vector2D = Nothing
            Dim secondFirtPoint As Vector2D = Nothing
            Dim pkStartPoint As Double = 9999999999
            Dim pkEndPoint As Double = -9999999999
            Dim off As Double = 0
            Dim topFirstBeam As StructureElement = firstProlet.First.Value
            If IsNothing(topFirstBeam) = True Then Return False
            Dim userBeam1 As BeamI = topFirstBeam.getBeamI
            'восстанавливаем 1 балку
            Dim startSectionPoint1 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim endSectionPoint1 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim restoreBeam1 As Boolean = CalculationBeams.restoreElementsBeam(topFirstBeam.DWGEntity, userBeam1, startSectionPoint1, endSectionPoint1, CalculationBeams.rectoreBeam.fullBeam)
            If firstProlet.Count > 1 Then
                Dim topLastBeam As StructureElement = firstProlet.Last.Value
                If IsNothing(topLastBeam) = True Then Return False
                Dim userBeam2 As BeamI = topLastBeam.getBeamI
                'восстанавливаем 2 балку
                Dim startSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam2 As Boolean = CalculationBeams.restoreElementsBeam(topLastBeam.DWGEntity, userBeam2, startSectionPoint2, endSectionPoint2, CalculationBeams.rectoreBeam.fullBeam)
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(0).Pos, startSectionPoint2.Item(0).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(0).Pos
                        secondFirtPoint = startSectionPoint2.Item(0).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(1).Pos, startSectionPoint2.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(1).Pos
                        secondFirtPoint = startSectionPoint2.Item(1).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(2).Pos, startSectionPoint2.Item(2).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(2).Pos
                        secondFirtPoint = startSectionPoint2.Item(2).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(3).Pos, startSectionPoint2.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(3).Pos
                        secondFirtPoint = startSectionPoint2.Item(3).Pos
                        pkStartPoint = pk
                    End If
                End If
            ElseIf firstProlet.Count = 1 Then
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(0).Pos, startSectionPoint1.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(0).Pos
                        secondFirtPoint = startSectionPoint1.Item(1).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(2).Pos, startSectionPoint1.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(2).Pos
                        secondFirtPoint = startSectionPoint1.Item(3).Pos
                        pkStartPoint = pk
                    End If
                End If
            End If
            Dim LastProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.Last.Value
            Dim startLastPoint As Vector2D = Nothing
            Dim secondLastPoint As Vector2D = Nothing
            Dim downLastBeam As StructureElement = LastProlet.First.Value
            If IsNothing(downLastBeam) = True Then Return False
            userBeam1 = downLastBeam.getBeamI
            'восстанавливаем 1 балку
            startSectionPoint1 = New List(Of Topomatic.Cad.Foundation.Vector3D)
            endSectionPoint1 = New List(Of Topomatic.Cad.Foundation.Vector3D)
            restoreBeam1 = CalculationBeams.restoreElementsBeam(downLastBeam.DWGEntity, userBeam1, startSectionPoint1, endSectionPoint1, CalculationBeams.rectoreBeam.fullBeam)
            If LastProlet.Count > 1 Then
                downLastBeam = LastProlet.Last.Value
                If IsNothing(downLastBeam) = True Then Return False
                Dim userBeam2 As BeamI = downLastBeam.getBeamI
                'восстанавливаем 2 балку
                Dim startSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam2 As Boolean = CalculationBeams.restoreElementsBeam(downLastBeam.DWGEntity, userBeam2, startSectionPoint2, endSectionPoint2, CalculationBeams.rectoreBeam.fullBeam)
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(0).Pos, endSectionPoint2.Item(0).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(0).Pos
                        secondLastPoint = endSectionPoint2.Item(0).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(1).Pos, endSectionPoint2.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(1).Pos
                        secondLastPoint = endSectionPoint2.Item(1).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(2).Pos, endSectionPoint2.Item(2).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(2).Pos
                        secondLastPoint = endSectionPoint2.Item(2).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(3).Pos, endSectionPoint2.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(3).Pos
                        secondLastPoint = endSectionPoint2.Item(3).Pos
                        pkEndPoint = pk
                    End If
                End If
            ElseIf firstProlet.Count = 1 Then
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(0).Pos, endSectionPoint1.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(0).Pos
                        secondLastPoint = endSectionPoint1.Item(1).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(2).Pos, endSectionPoint1.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(2).Pos
                        secondLastPoint = endSectionPoint1.Item(3).Pos
                        pkEndPoint = pk
                    End If
                End If
            End If

            Dim pt1 As Vector3D = New Vector3D(startFirtPoint, 0)
            Dim pt2 As Vector3D = New Vector3D(secondFirtPoint, 0)
            Dim boolExtendLine1 As Boolean = MathFunction.FuncExtendPos(pt1, pt2, 100, 100)

            Dim pt3 As Vector3D = New Vector3D(startLastPoint, 0)
            Dim pt4 As Vector3D = New Vector3D(secondLastPoint, 0)
            Dim boolExtendLine2 As Boolean = MathFunction.FuncExtendPos(pt3, pt4, 100, 100)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем пересечение с левыс и правым габаритом
            Dim pkStartDirect As Double = -999999999
            Dim pkEndDirect As Double = -999999999
            Dim pkStartReverse As Double = -999999999
            Dim pkEndReverse As Double = -999999999
            Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, pt1.Pos, pt2.Pos)
            If pointIntersectCollection1.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection1.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkStartDirect, off)
            End If
            Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, pt1.Pos, pt2.Pos)
            If pointIntersectCollection2.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection2.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkStartReverse, off)
            End If
            Dim pointIntersectCollection3 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, pt3.Pos, pt4.Pos)
            If pointIntersectCollection3.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection3.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkEndDirect, off)
            End If
            Dim pointIntersectCollection4 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, pt3.Pos, pt4.Pos)
            If pointIntersectCollection4.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection4.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkEndReverse, off)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'обрезаем полилинии
            If pkStartDirect <> -999999999 And pkStartReverse <> -999999999 And pkEndDirect <> -999999999 And pkEndReverse <> -999999999 Then
                Dim pline2dCurv1 As Polyline2DCurve = New Polyline2DCurve()
                'pline2dCurv1.Add(New BugleVector2D(pointIntersectCollection1.ElementAt(0), 0))
                For i As Double = pkStartDirect To pkEndDirect Step 1
                    Dim pt As Vector2D = axisPline3DDirect.StaOffsetToPos(i, 0)
                    pline2dCurv1.Add(New BugleVector2D(pt, 0))
                Next
                pline2dCurv1.Add(New BugleVector2D(pointIntersectCollection3.ElementAt(0), 0))

                Dim pline2dCurv2 As Polyline2DCurve = New Polyline2DCurve()
                For i As Double = pkEndReverse To pkStartReverse Step 1
                    Dim pt As Vector2D = axisPline3DReverse.StaOffsetToPos(i, 0)
                    pline2dCurv2.Add(New BugleVector2D(pt, 0))
                Next
                pline2dCurv2.Add(New BugleVector2D(pointIntersectCollection2.ElementAt(0), 0))

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ось трассы
                Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
                For i As Integer = 0 To axisCentrePline.Count - 1
                    pline2dCurv.Add(axisCentrePline.Item(i))
                Next i
                Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(pkStartPoint, pkEndPoint)
                Dim pline2dCurv3 As Polyline2DCurve = New Polyline2DCurve()
                If IsArray(arrayPlineCurve) = True Then
                    For i As Integer = 0 To arrayPlineCurve.Length - 1
                        Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                        If Math.Abs(plineCurv.Length - (pkEndPoint - pkStartPoint)) <= 0.01 Then
                            pline2dCurv3 = plineCurv
                        End If
                    Next i
                End If
                If pline2dCurv3.Length > 0 Then
                    If IsNothing(axisBridge) = True Then
                        axisBridge = New DwgPolyline()
                        ActivDocument.ActiveSpace.Entities.Add(axisBridge)
                    ElseIf axisBridge.Length = 0 Then
                        If ActivDocument.ActiveSpace.Entities.Contains(axisBridge) = False Then
                            ActivDocument.ActiveSpace.Entities.Add(axisBridge)
                        End If
                    Else
                        axisBridge.Clear()
                    End If

                    For i As Integer = 0 To pline2dCurv3.Count - 1
                        axisBridge.Add(pline2dCurv3.Item(i))
                    Next i
                End If

                If pline2dCurv2.Length > 0 And pline2dCurv1.Length > 0 Then
                    If IsNothing(boundStructure) = True Then
                        boundStructure = New DwgPolyline()
                        ActivDocument.ActiveSpace.Entities.Add(boundStructure)
                    ElseIf boundStructure.Length = 0 Then
                        ActivDocument.ActiveSpace.Entities.Add(boundStructure)
                    Else
                        boundStructure.Clear()
                    End If

                    For i As Integer = 0 To pline2dCurv1.Count - 1
                        boundStructure.Add(pline2dCurv1.Item(i))
                    Next i
                    Dim count As Integer = pline2dCurv2.Count - 1
                    For i As Integer = 0 To pline2dCurv2.Count - 1
                        Dim bulg As Double = 0
                        If i <> 0 Then
                            bulg = pline2dCurv2.Item(i - 1).Bugle
                        End If
                        boundStructure.Add(pline2dCurv2.Item(i))
                    Next i
                    boundStructure.Closed = True
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As System.Exception
        Finally
            If ActivDocument.ActiveSpace.Entities.Contains(axisPlineDirect) = True Then
                ActivDocument.ActiveSpace.Entities.Remove(axisPlineDirect)
            End If
            If ActivDocument.ActiveSpace.Entities.Contains(axisPlineReverse) = True Then
                ActivDocument.ActiveSpace.Entities.Remove(axisPlineReverse)
            End If
            If ActivDocument.ActiveSpace.Entities.Contains(axisCentrePline) = True Then
                ActivDocument.ActiveSpace.Entities.Remove(axisCentrePline)
            End If
        End Try
    End Function
    '=========================================================================================================
    'создать участки омоличивания для балок
    Public Function CreateMonolithingBeams(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByRef dictionaryBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal idBridge As String, ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional boolDeleteSiteMonolit As Boolean = False, Optional templateXML As String = "", Optional fullMonolitBeams As Boolean = False, Optional numberProlet As Integer = 0)
        Dim nameHatch As String = "ANSI31"
        If dictionaryBeams.Count = 0 Then
            MsgBox("Словарь с марками балок пуст. Сооружение не построено!!!")
            Return False
        End If
        '1. Оформление
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim deltaAxisPillar As Double = 2 'величина выпуска осей опоры за границы сооружения
        Dim categoryTables As String = "Искусственные сооружения"
        Dim nameTableBeams As String = "Балки мостовых сооружений"
        Dim funcBridges As Bridges = New Bridges
        'ось балки
        Dim drawingMonolitSites As Topomatic.Dwg.Drawing = ActivDocument
        Dim drawObject As CreateDwgObject = New CreateDwgObject(drawingMonolitSites)
        If IsNothing(drawingMonolitSites) = True Then
            MsgBox("Активный проект для создания участков омоноличивания не найден.")
            Return False
        End If
        'ось 
        Dim styleAxisMonolitSiteBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingMonolitSites)
        styleAxisMonolitSiteBeams.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Учасок омоноличивания балок (ось)")
        'контцр (верх)
        Dim styleCounterMonolitSiteBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingMonolitSites)
        styleCounterMonolitSiteBeams.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Участок омоноличивания балок (верх контура)")
        'контцр (низ)
        Dim styleBottomCounterMonolitSiteBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingMonolitSites)
        styleBottomCounterMonolitSiteBeams.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Участок омоноличивания балок (низ контура)")
        'модель
        Dim styleModelMonolitSiteBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(drawingMonolitSites)
        styleModelMonolitSiteBeams.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Участок омоноличивания балок (модель)")

        If dictionaryBeams.Count > 0 Then
            Dim listAxisSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
            If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.axisSiteMonolitBeams) = True Then
                listAxisSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.axisSiteMonolitBeams)
            End If
            Dim listTopCounterSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
            If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.counterSiteMonolitBeamsTop) = True Then
                listTopCounterSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.counterSiteMonolitBeamsTop)
            End If
            Dim listBottomCounterSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
            If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.counterSiteMonolitBeamsBottom) = True Then
                listBottomCounterSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.counterSiteMonolitBeamsBottom)
            End If
            Dim listHatchSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
            If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.hatchSiteMonolitPillar) = True Then
                listHatchSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.hatchSiteMonolitPillar)
            End If
            Dim listModelSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
            If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.modelSiteMonolitBeams) = True Then
                listModelSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.modelSiteMonolitBeams)
            End If
            'идем по пролетам
            For i As Integer = 0 To dictionaryBeams.Count - 1
                'номер пролета
                Dim numbProlet As Integer = dictionaryBeams.ElementAt(i).Key
                'список балок в пролете
                Dim dictBeams As Dictionary(Of Integer, StructureElement) = dictionaryBeams.ElementAt(i).Value
                If IsNothing(dictBeams) = False Then
                    If dictBeams.Count > 1 Then
                        'удаляем из словаря несуществующие балки
                        For Each keyElement As Integer In dictBeams.Keys.ToList()
                            Dim dataStructureBeam As StructureElement = dictBeams(keyElement)
                            If dataStructureBeam Is Nothing OrElse dataStructureBeam.DWGEntity Is Nothing Then
                                dictBeams.Remove(keyElement)
                            End If
                        Next
                        'перебираем оси балок
                        For j As Integer = 0 To dictBeams.Count - 2
                            'первая балка
                            Dim dataStructureBeam1 As StructureElement = dictBeams.ElementAt(j).Value
                            Dim userBeam1 As BeamI = dataStructureBeam1.getBeamI()
                            Dim firstPlateTopBeam As Dictionary(Of Integer, PointStructure) = userBeam1._elementBridgePoint.ListPointModel
                            Dim rightTopLine As DwgLine = New DwgLine()
                            rightTopLine.StartPoint = userBeam1.getPointBeam("rightPt1")
                            rightTopLine.EndPoint = userBeam1.getPointBeam("rightPt2")
                            Dim rightBottomLine As DwgLine = New DwgLine()
                            rightBottomLine.StartPoint = userBeam1.getPointBeam("rightPt1", False)
                            rightBottomLine.EndPoint = userBeam1.getPointBeam("rightPt2", False)
                            'вторая балка
                            Dim dataStructureBeam2 As StructureElement = dictBeams.ElementAt(j + 1).Value
                            Dim userBeam2 As BeamI = dataStructureBeam2.getBeamI()
                            Dim secondPlateTopBeam As Dictionary(Of Integer, PointStructure) = userBeam2._elementBridgePoint.ListPointModel
                            Dim leftTopLine As DwgLine = New DwgLine()
                            leftTopLine.StartPoint = userBeam2.getPointBeam("leftPt1")
                            leftTopLine.EndPoint = userBeam2.getPointBeam("leftPt2")
                            Dim leftBottomLine As DwgLine = New DwgLine()
                            leftBottomLine.StartPoint = userBeam2.getPointBeam("leftPt1", False)
                            leftBottomLine.EndPoint = userBeam2.getPointBeam("leftPt2", False)
                            'номер пролета учпстка омоноличивания
                            Dim numberProletSiteMonolit As Integer = numbProlet
                            'номер ряда участка омоноличивания
                            Dim numberRowsSiteMonolit As Integer = j + 1
                            'ищем уже существующий участок омоноличивания балок
                            Dim dataSiteMonolit As StructureElement = SiteMonolitBeams.getAxisMonolitSitesBeam(dictinaryAllObjectBridge, numberProletSiteMonolit, userBeam1.numberRow, userBeam2.numberRow)
                            Dim userSiteMonolit As SiteMonolitBeams = New SiteMonolitBeams
                            If IsNothing(dataSiteMonolit) = True Then
                                dataSiteMonolit = SiteMonolitBeams.createAxis(idBridge)
                            Else
                                userSiteMonolit = dataSiteMonolit.getMonolitSiteBeams
                                numberRowsSiteMonolit = userSiteMonolit.numberRow
                            End If
                            If IsNothing(dataSiteMonolit) = True Then Continue For
                            Dim axisLineSiteMonolit As DwgLine = dataSiteMonolit.DWGEntity
                            'ищем контур по верху
                            Dim dataTopSiteMonolit As StructureElement = CounterSiteMonolitBeams.getCounterMonolitSitesBeam(dictinaryAllObjectBridge, numberProletSiteMonolit, numberRowsSiteMonolit, StructureElement.typeObject.counterSiteMonolitBeamsTop)
                            Dim userTopSiteMonolit As CounterSiteMonolitBeams = New CounterSiteMonolitBeams
                            If IsNothing(dataTopSiteMonolit) = True Then
                                dataTopSiteMonolit = CounterSiteMonolitBeams.createAxis(idBridge, StructureElement.typeObject.counterSiteMonolitBeamsTop)
                            Else
                                userTopSiteMonolit = dataTopSiteMonolit.getCounterMonolitSiteBeams
                            End If
                            If IsNothing(dataTopSiteMonolit) = True Then Continue For
                            Dim poly3dTopSiteMonolit As DwgPolyline3D = dataTopSiteMonolit.DWGEntity
                            'ищем контур по низу
                            Dim dataBottomSiteMonolit As StructureElement = CounterSiteMonolitBeams.getCounterMonolitSitesBeam(dictinaryAllObjectBridge, numberProletSiteMonolit, numberRowsSiteMonolit, StructureElement.typeObject.counterSiteMonolitBeamsBottom)
                            Dim userBottomSiteMonolit As CounterSiteMonolitBeams = New CounterSiteMonolitBeams
                            If IsNothing(dataBottomSiteMonolit) = True Then
                                dataBottomSiteMonolit = CounterSiteMonolitBeams.createAxis(idBridge, StructureElement.typeObject.counterSiteMonolitBeamsBottom)
                            Else
                                userBottomSiteMonolit = dataBottomSiteMonolit.getCounterMonolitSiteBeams
                            End If
                            If IsNothing(dataBottomSiteMonolit) = True Then Continue For
                            Dim poly3dBottomSiteMonolit As DwgPolyline3D = dataBottomSiteMonolit.DWGEntity
                            'ищем штриховку
                            Dim dataHatchSiteMonolit As StructureElement = HatchSiteMonolitBeams.getHatchMonolitSitesBeam(dictinaryAllObjectBridge, numberProletSiteMonolit, numberRowsSiteMonolit)
                            Dim userHatchSiteMonolit As HatchSiteMonolitBeams = New HatchSiteMonolitBeams
                            If IsNothing(dataHatchSiteMonolit) = True Then
                                dataHatchSiteMonolit = HatchSiteMonolitBeams.createAxis(idBridge)
                            Else
                                userHatchSiteMonolit = dataHatchSiteMonolit.getHatchMonolitSiteBeams
                            End If
                            If IsNothing(dataHatchSiteMonolit) = True Then Continue For
                            Dim hatchSiteMonolit As DwgHatch = dataHatchSiteMonolit.DWGEntity
                            'ищем существующую модель
                            Dim dataModelSiteMonolit As StructureElement = ModelSiteMonolitBeams.getModelMonolitSitesBeam(dictinaryAllObjectBridge, numberProletSiteMonolit, numberRowsSiteMonolit)
                            Dim userModelSiteMonolit As ModelSiteMonolitBeams = New ModelSiteMonolitBeams
                            If IsNothing(dataModelSiteMonolit) = True Then
                                dataModelSiteMonolit = ModelSiteMonolitBeams.createModel(idBridge)
                            Else
                                userModelSiteMonolit = dataModelSiteMonolit.getModelSiteMonolitBeams
                            End If
                            If IsNothing(dataModelSiteMonolit) = True Then Continue For
                            Dim modelSiteMonolit As DwgModel3DElement = dataModelSiteMonolit.DWGEntity

                            '=======================================================================================================================
                            'рисуем полилинию
                            Dim userTopListVertex As List(Of Vector3D) = New List(Of Vector3D)
                            userTopListVertex.Add(rightTopLine.StartPoint)
                            userTopListVertex.Add(rightTopLine.EndPoint)
                            userTopListVertex.Add(leftTopLine.EndPoint)
                            userTopListVertex.Add(leftTopLine.StartPoint)
                            '3d по верху
                            Dim userBottomListVertex As List(Of Vector3D) = New List(Of Vector3D)
                            userBottomListVertex.Add(rightBottomLine.StartPoint)
                            userBottomListVertex.Add(rightBottomLine.EndPoint)
                            userBottomListVertex.Add(leftBottomLine.EndPoint)
                            userBottomListVertex.Add(leftBottomLine.StartPoint)
                            'если омоноличивание идет под прямым углом
                            If fullMonolitBeams = False Then
                                Dim angleNLine4 As Double = leftTopLine.Rotation - Math.PI / 2
                                If angleNLine4 < 0 Then
                                    angleNLine4 += 2 * Math.PI
                                End If
                                'в противоположном напрпвлении
                                Dim angleNLine1 As Double = rightTopLine.Rotation + Math.PI / 2
                                If angleNLine1 > 2 * Math.PI Then
                                    angleNLine1 -= 2 * Math.PI
                                End If
                                'начало балки
                                Dim pt4 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightTopLine.StartPoint.Pos, angleNLine4, 10)
                                Dim intersectPt4 As Vector2D = Nothing
                                Dim boolIntersect4 As Boolean = MathFunction.FuncIntersectionTwoSegments(leftTopLine.EndPoint.Pos, leftTopLine.StartPoint.Pos, rightTopLine.StartPoint.Pos, pt4, intersectPt4)
                                If boolIntersect4 = True Then
                                    Dim hPt4 As Double = MathFunction.FuncCalcElevationByLine(leftTopLine.StartPoint, leftTopLine.EndPoint, intersectPt4)
                                    userTopListVertex.Item(3) = New Vector3D(intersectPt4, hPt4)
                                    Dim pointDown As Vector3D = BridgeGeometry.calculatePointP(userTopListVertex.Item(3), leftTopLine.StartPoint, 0, -1 * userBeam2.heightTopPlate)
                                    userBottomListVertex.Item(3) = pointDown
                                Else
                                    Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftTopLine.StartPoint.Pos, angleNLine1, 10)
                                    Dim intersectPt1 As Vector2D = Nothing
                                    Dim boolIntersect1 As Boolean = MathFunction.FuncIntersectionTwoSegments(rightTopLine.EndPoint.Pos, rightTopLine.StartPoint.Pos, leftTopLine.StartPoint.Pos, pt1, intersectPt1)
                                    If boolIntersect1 = True Then
                                        Dim hPt1 As Double = MathFunction.FuncCalcElevationByLine(rightTopLine.StartPoint, rightTopLine.EndPoint, intersectPt1)
                                        userTopListVertex.Item(0) = New Vector3D(intersectPt1, hPt1)
                                        Dim pointDown As Vector3D = BridgeGeometry.calculatePointP(userTopListVertex.Item(0), rightTopLine.EndPoint, 0, -1 * userBeam1.heightTopPlate)
                                        userBottomListVertex.Item(0) = pointDown
                                    End If
                                End If
                                'конец балки
                                Dim pt3 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightTopLine.EndPoint.Pos, angleNLine4, 10)
                                Dim intersectPt3 As Vector2D = Nothing
                                Dim boolIntersect3 As Boolean = MathFunction.FuncIntersectionTwoSegments(leftTopLine.StartPoint.Pos, leftTopLine.EndPoint.Pos, rightTopLine.EndPoint.Pos, pt3, intersectPt3)
                                If boolIntersect3 = True Then
                                    Dim hPt3 As Double = MathFunction.FuncCalcElevationByLine(leftTopLine.StartPoint, leftTopLine.EndPoint, intersectPt3)
                                    userTopListVertex.Item(2) = New Vector3D(intersectPt3, hPt3)
                                    Dim pointDown As Vector3D = BridgeGeometry.calculatePointP(userTopListVertex.Item(2), leftTopLine.StartPoint, 0, -1 * userBeam2.heightTopPlate)
                                    userBottomListVertex.Item(2) = pointDown
                                Else
                                    Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftTopLine.EndPoint.Pos, angleNLine1, 10)
                                    Dim intersectPt2 As Vector2D = Nothing
                                    Dim boolIntersect2 As Boolean = MathFunction.FuncIntersectionTwoSegments(rightTopLine.StartPoint.Pos, rightTopLine.EndPoint.Pos, leftTopLine.EndPoint.Pos, pt2, intersectPt2)
                                    If boolIntersect2 = True Then
                                        Dim hPt2 As Double = MathFunction.FuncCalcElevationByLine(rightTopLine.StartPoint, rightTopLine.EndPoint, intersectPt2)
                                        userTopListVertex.Item(1) = New Vector3D(intersectPt2, hPt2)
                                        Dim pointDown As Vector3D = BridgeGeometry.calculatePointP(userTopListVertex.Item(1), rightTopLine.StartPoint, 0, -1 * userBeam1.heightTopPlate)
                                        userBottomListVertex.Item(1) = pointDown
                                    End If
                                End If
                            End If
                            'рисуем ось
                            Dim centerStart As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(userTopListVertex.Item(0), userTopListVertex.Item(3))
                            Dim centerEnd As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(userTopListVertex.Item(1), userTopListVertex.Item(2))
                            axisLineSiteMonolit.StartPoint = centerStart
                            axisLineSiteMonolit.EndPoint = centerEnd
                            userSiteMonolit.numberProlet = numberProletSiteMonolit
                            userSiteMonolit.numberLeftBeam = userBeam1.numberRow
                            userSiteMonolit.numberRightBeam = userBeam2.numberRow
                            userSiteMonolit.numberRow = numberRowsSiteMonolit
                            userSiteMonolit.thickness = Math.Round((userBeam2.heightTopPlate + userBeam1.heightTopPlate) / 2, 3)
                            userSiteMonolit._elementBridgePoint.StartAxisPoint = centerStart
                            userSiteMonolit._elementBridgePoint.EndAxisPoint = centerEnd
                            Dim listModelPoint As New Dictionary(Of Integer, PointStructure)
                            Dim x As Double = Math.Round(userTopListVertex.Item(0).X, 3)
                            Dim y As Double = Math.Round(userTopListVertex.Item(0).Y, 3)
                            Dim z As Double = Math.Round(userTopListVertex.Item(0).Z, 3)
                            Dim dx As Double = 0
                            Dim dy As Double = 0
                            Dim dz As Double = Math.Round(userBeam2.heightTopPlate, 3)
                            Dim code As String = "leftPt1"
                            Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                            listModelPoint.Add(1, pointModel)

                            x = Math.Round(userTopListVertex.Item(1).X, 3)
                            y = Math.Round(userTopListVertex.Item(1).Y, 3)
                            z = Math.Round(userTopListVertex.Item(1).Z, 3)
                            dx = 0
                            dy = 0
                            dz = Math.Round(userBeam1.heightTopPlate, 3)
                            code = "leftPt2"
                            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                            listModelPoint.Add(2, pointModel)

                            x = Math.Round(userTopListVertex.Item(2).X, 3)
                            y = Math.Round(userTopListVertex.Item(2).Y, 3)
                            z = Math.Round(userTopListVertex.Item(2).Z, 3)
                            dx = 0
                            dy = 0
                            dz = Math.Round(userBeam1.heightTopPlate, 3)
                            code = "rightPt2"
                            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                            listModelPoint.Add(3, pointModel)

                            x = Math.Round(userTopListVertex.Item(3).X, 3)
                            y = Math.Round(userTopListVertex.Item(3).Y, 3)
                            z = Math.Round(userTopListVertex.Item(3).Z, 3)
                            dx = 0
                            dy = 0
                            dz = Math.Round(userBeam2.heightTopPlate, 3)
                            code = "rightPt1"
                            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                            listModelPoint.Add(4, pointModel)
                            userSiteMonolit._elementBridgePoint.ListPointModel = listModelPoint

                            Dim strGSONModelBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userSiteMonolit)
                            If drawingMonolitSites.ActiveSpace.Entities.Contains(axisLineSiteMonolit) = False Then
                                drawingMonolitSites.ActiveSpace.Add(axisLineSiteMonolit)
                            End If
                            dataSiteMonolit.KeyParameter = strGSONModelBeam
                            dataSiteMonolit.DWGEntity = axisLineSiteMonolit
                            Dim boolRecDataModelBeam = FuncXRecords.setXRecords(axisLineSiteMonolit, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataSiteMonolit)
                            Dim boolAxisStyle As Boolean = styleAxisMonolitSiteBeams.setObjectStyle(axisLineSiteMonolit)
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'рисуем контур по верху
                            If drawingMonolitSites.ActiveSpace.Entities.Contains(poly3dTopSiteMonolit) = True Then
                                Dim boolRedraw As Boolean = drawObject.reDrawPolyline3D(poly3dTopSiteMonolit, userTopListVertex)
                            Else
                                userTopSiteMonolit.numberProlet = numberProletSiteMonolit
                                userTopSiteMonolit.numberRow = numberRowsSiteMonolit
                                userTopSiteMonolit.TypeCounter = StructureElement.typeObject.counterSiteMonolitBeamsTop
                                poly3dTopSiteMonolit = drawObject.createPolyline3D(userTopListVertex, True)
                            End If
                            strGSONModelBeam = Newtonsoft.Json.JsonConvert.SerializeObject(userTopSiteMonolit)
                            dataTopSiteMonolit.KeyParameter = strGSONModelBeam
                            dataTopSiteMonolit.DWGEntity = poly3dTopSiteMonolit
                            Dim boolRecDataTopCounterBeam = FuncXRecords.setXRecords(poly3dTopSiteMonolit, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopSiteMonolit)
                            Dim boolCreateStyle As Boolean = styleCounterMonolitSiteBeams.setObjectStyle(poly3dTopSiteMonolit)
                            'рисуем контур по низу
                            If drawingMonolitSites.ActiveSpace.Entities.Contains(poly3dBottomSiteMonolit) = True Then
                                Dim boolRedraw As Boolean = drawObject.reDrawPolyline3D(poly3dBottomSiteMonolit, userBottomListVertex)
                            Else
                                userBottomSiteMonolit.numberProlet = numberProletSiteMonolit
                                userBottomSiteMonolit.numberRow = numberRowsSiteMonolit
                                userBottomSiteMonolit.TypeCounter = StructureElement.typeObject.counterSiteMonolitBeamsBottom
                                poly3dBottomSiteMonolit = drawObject.createPolyline3D(userBottomListVertex, True)
                            End If
                            strGSONModelBeam = Newtonsoft.Json.JsonConvert.SerializeObject(userBottomSiteMonolit)
                            dataBottomSiteMonolit.KeyParameter = strGSONModelBeam
                            dataBottomSiteMonolit.DWGEntity = poly3dBottomSiteMonolit
                            boolCreateStyle = styleBottomCounterMonolitSiteBeams.setObjectStyle(poly3dBottomSiteMonolit)
                            Dim boolRecData = FuncXRecords.setXRecords(poly3dBottomSiteMonolit, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomSiteMonolit)
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'рисуем штриховку
                            If ActivDocument.ActiveSpace.Entities.Contains(hatchSiteMonolit) = True Then
                                Dim boolRedraw As Boolean = drawObject.reDrawHatchByPolyline3d(poly3dTopSiteMonolit, hatchSiteMonolit)
                            Else
                                hatchSiteMonolit = drawObject.createHatchToPolyline3d(poly3dTopSiteMonolit, nameHatch)
                                userHatchSiteMonolit.numberProlet = numberProletSiteMonolit
                                userHatchSiteMonolit.numberRow = numberRowsSiteMonolit
                            End If
                            strGSONModelBeam = Newtonsoft.Json.JsonConvert.SerializeObject(userHatchSiteMonolit)
                            dataHatchSiteMonolit.KeyParameter = strGSONModelBeam
                            dataHatchSiteMonolit.DWGEntity = hatchSiteMonolit
                            boolRecData = FuncXRecords.setXRecords(hatchSiteMonolit, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataHatchSiteMonolit)
                            boolCreateStyle = styleModelMonolitSiteBeams.setObjectStyle(hatchSiteMonolit)
                            'рисуем модель
                            Dim shell As Shell = FuncModeling3d.FuncCreateSolidByTwoNPolyline3d(drawingMonolitSites, poly3dTopSiteMonolit, poly3dBottomSiteMonolit)
                            If IsNothing(shell) = False Then
                                Dim originPoint = shell.Vertices.First.Position
                                Dim elementModel = New StaticSolidElement("Участок омоноличивания балок (модель)", "SmdxElement", New ImProperties(), shell, New ImDocuments())
                                elementModel.Origin = originPoint
                                modelSiteMonolit.Position = elementModel.Origin
                                modelSiteMonolit.Element = elementModel
                                If ActivDocument.ActiveSpace.Entities.Contains(modelSiteMonolit) = False Then
                                    ActivDocument.ActiveSpace.Add(modelSiteMonolit)
                                End If
                                If ActivDocument.ActiveSpace.Entities.Contains(modelSiteMonolit) = True Then
                                    userModelSiteMonolit.numberProlet = numberProletSiteMonolit
                                    userModelSiteMonolit.numberRow = numberRowsSiteMonolit
                                    strGSONModelBeam = Newtonsoft.Json.JsonConvert.SerializeObject(userModelSiteMonolit)
                                    dataModelSiteMonolit.KeyParameter = strGSONModelBeam
                                    dataModelSiteMonolit.DWGEntity = modelSiteMonolit
                                    boolRecData = FuncXRecords.setXRecords(modelSiteMonolit, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModelSiteMonolit)
                                    boolCreateStyle = styleModelMonolitSiteBeams.setObjectStyle(modelSiteMonolit)
                                End If
                            End If
                        Next j
                    End If
                End If
            Next i
        End If
        Return True
    End Function
    '========================================================================================================
    'функция водвращает ось объекта по его 3д модели
    Public Shared Function getAxisElementsByModel3d(ByVal acModel3d As DwgModel3DElement) As DwgEntity
        getAxisElementsByModel3d = Nothing
        If IsNothing(acModel3d) = False Then
            Dim dataElements As StructureElement = Nothing
            Dim boolDataPs1 As Boolean = FuncXRecords.getXRecords(acModel3d, dataElements)
            If IsNothing(dataElements) Then
                Dim localId As String = dataElements.IdElement
                Dim ActivDocument As Topomatic.Dwg.Drawing = acModel3d.Drawing
                If IsNothing(ActivDocument) = False Then
                    For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                        If TypeOf (acEnt) Is DwgEntity Then
                            Dim dataEntity As StructureElement = Nothing
                            Dim boolDataEnt As Boolean = FuncXRecords.getXRecords(acEnt, dataEntity)
                            If dataEntity.IdElement Like localId Then
                                Return acEnt
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает имена мостов
    Public Shared Function getBridgeObject(ByVal activeDoc As Topomatic.Dwg.Drawing, ByRef bridgeGeneralAxisDictionary As Dictionary(Of String, String())) As Boolean
        bridgeGeneralAxisDictionary = New Dictionary(Of String, String())
        getBridgeObject = False
        If IsNothing(activeDoc) = False Then
            For Each acEnt As DwgEntity In activeDoc.ActiveSpace.Entities
                Dim dataBridge As StructureElement = Nothing
                Dim readData As Boolean = FuncXRecords.getXRecords(acEnt, dataBridge, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                If IsNothing(dataBridge) = False And readData = True Then
                    If dataBridge.Name = StructureElement.typeObject.axisBridge Then
                        Dim userBridge As Bridges = dataBridge.getBridge()
                        Dim arrayWrite As String() = {}
                        ReDim arrayWrite(4)
                        arrayWrite(0) = userBridge.NameBridge
                        arrayWrite(1) = dataBridge.KeyParameter
                        arrayWrite(2) = dataBridge.IdElement
                        arrayWrite(3) = dataBridge.IdStructure
                        arrayWrite(4) = dataBridge.DWGEntity.ObjectID
                        If bridgeGeneralAxisDictionary.ContainsKey(dataBridge.IdStructure) = False Then
                            bridgeGeneralAxisDictionary.Add(dataBridge.IdStructure, arrayWrite)
                        End If
                    End If
                End If
            Next
        End If
        If bridgeGeneralAxisDictionary.Count > 0 Then
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает словарь с номерами рядов балок
    Public Function getConditionalRows() As Dictionary(Of Integer, String)
        Dim result As Dictionary(Of Integer, String) = New Dictionary(Of Integer, String)
        Dim countLeftRow As Integer = LeftRowsCount
        If countLeftRow > 0 Then
            For i As Integer = countLeftRow To 1 Step -1
                result.Add(-1 * i, "Л" & i)
            Next
        End If
        If centerAxis = True Then
            result.Add(0, "Ось")
        End If
        Dim countRightRow As Integer = RightRowsCount
        If countRightRow > 0 Then
            For i As Integer = 1 To countRightRow
                result.Add(i, "П" & i)
            Next
        End If
        Return result
    End Function
    'функция рисует ось сооружения
    Public Shared Function drawAxisBridge(ByRef drawingDocument As Topomatic.Dwg.Drawing, ByVal dataBridge As StructureElement, ByVal axisPillars As Dictionary(Of Integer, List(Of StructureElement)), ByVal projectAlignment As Alignment, Optional styleAxisBridge As ProjectCivilStructuresStyle = Nothing, Optional offsetAlign As Double = 5) As Boolean
        Dim result As Boolean = False
        If IsNothing(drawingDocument) = True Then
            Return False
        End If
        If IsNothing(dataBridge) = True Then
            Return False
        End If
        If IsNothing(axisPillars) = True Then
            Return False
        End If
        If axisPillars.Count < 2 Then
            Return False
        End If
        If IsNothing(projectAlignment) = True Then
            Return False
        End If
        Dim userBridge As Bridges = dataBridge.getBridge()
        If IsNothing(userBridge) = True Then
            Return False
        End If
        Dim axisStructure As DwgPolyline = New DwgPolyline
        If IsNothing(dataBridge.DWGEntity) = False Then
            axisStructure = dataBridge.DWGEntity
        End If
        'преобразуем трассу в полилинию
        Dim axisPlineAlign As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, 0)
        Dim axisPlineAlign3D = New Polyline3D()
        axisPlineAlign.GetPolyline(axisPlineAlign3D)
        If axisPlineAlign3D.Length2D > 0 Then
            'находим первую и последнюю опору
            Dim dataFirstPillar As StructureElement = axisPillars.First.Value.Item(1)
            Dim dataLastPillar As StructureElement = axisPillars.Last.Value.Item(1)
            If IsNothing(dataFirstPillar) = False And IsNothing(dataLastPillar) = False Then
                Dim firstPillar As DwgLine = dataFirstPillar.DWGEntity
                Dim lastPillar As DwgLine = dataLastPillar.DWGEntity
                If IsNothing(firstPillar) = False And IsNothing(lastPillar) = False Then
                    If firstPillar.Length > 0 And lastPillar.Length > 0 Then
                        Dim firstPoint As Vector2D = axisPlineAlign.Item(0).Vertex
                        Dim pointIntersectCollectionStart As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlineAlign3D, firstPillar.StartPoint.Pos, firstPillar.EndPoint.Pos)
                        If pointIntersectCollectionStart.Count > 0 Then
                            firstPoint = pointIntersectCollectionStart(0)
                        End If
                        Dim lastPoint As Vector2D = axisPlineAlign.Item(axisPlineAlign.Count - 1).Vertex
                        Dim pointIntersectCollectionEnd As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlineAlign3D, lastPillar.StartPoint.Pos, lastPillar.EndPoint.Pos)
                        If pointIntersectCollectionEnd.Count > 0 Then
                            lastPoint = pointIntersectCollectionEnd(0)
                        End If
                        Dim startPK As Double = 0
                        Dim startOff As Double = 0
                        Dim boolFindPk As Boolean = projectAlignment.Plan.CompoundLine.PosToStaOffset(firstPoint, startPK, startOff)
                        Dim endPK As Double = projectAlignment.Plan.CompoundLine.Length
                        Dim endOff As Double = 0
                        Dim boolFindPkEnd As Boolean = projectAlignment.Plan.CompoundLine.PosToStaOffset(lastPoint, endPK, endOff)
                        '===============================================================================================================
                        Dim newStartPk As Double = startPK - offsetAlign
                        Dim newEndPk As Double = endPK + offsetAlign
                        Dim newStartPoint As Vector2D = axisPlineAlign.Item(0).Vertex
                        Dim boolFindStartPoint As Boolean = projectAlignment.Plan.CompoundLine.StaOffsetToPos(newStartPk, 0, newStartPoint)
                        If boolFindStartPoint = False Then
                            newStartPoint = axisPlineAlign.Item(0).Vertex
                        End If
                        Dim newEndPoint As Vector2D = axisPlineAlign.Item(axisPlineAlign.Count - 1).Vertex
                        Dim boolFindEndPoint As Boolean = projectAlignment.Plan.CompoundLine.StaOffsetToPos(newEndPk, 0, newEndPoint)
                        If boolFindEndPoint = False Then
                            newEndPoint = axisPlineAlign.Item(axisPlineAlign.Count - 1).Vertex
                        End If
                        '=============================================================================================================
                        'рисуем ось
                        Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
                        For i As Integer = 0 To axisPlineAlign.Count - 1
                            pline2dCurv.Add(axisPlineAlign.Item(i))
                        Next i
                        Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(newStartPk, newEndPk)
                        Dim pline2dCurv3 As Polyline2DCurve = New Polyline2DCurve()
                        If IsArray(arrayPlineCurve) = True Then
                            For i As Integer = 0 To arrayPlineCurve.Length - 1
                                Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                                If Math.Abs(plineCurv.Length - (newEndPk - newStartPk)) <= 0.01 Then
                                    pline2dCurv3 = plineCurv
                                End If
                            Next i
                        End If
                        If pline2dCurv3.Length > 0 Then
                            If IsNothing(axisStructure) = True Then
                                axisStructure = New DwgPolyline()
                                drawingDocument.ActiveSpace.Entities.Add(axisStructure)
                                styleAxisBridge.setObjectStyle(axisStructure)
                            ElseIf axisStructure.Length = 0 Then
                                If drawingDocument.ActiveSpace.Entities.Contains(axisStructure) = False Then
                                    drawingDocument.ActiveSpace.Entities.Add(axisStructure)
                                    styleAxisBridge.setObjectStyle(axisStructure)
                                End If
                            Else
                                axisStructure.Clear()
                            End If
                            For i As Integer = 0 To pline2dCurv3.Count - 1
                                axisStructure.Add(pline2dCurv3.Item(i))
                            Next i
                            dataBridge.DWGEntity = axisStructure
                        End If
                    End If
                End If
            End If
            '===================================================================================================================================
            'записываем пикет начала раскладки балок
            If userBridge.startPlacementPosition = 0 Then
                If axisPillars.Count > 0 Then
                    For i As Integer = 0 To axisPillars.Count - 1
                        Dim listPillar As List(Of StructureElement) = axisPillars.ElementAt(i).Value
                        Dim dataPillar As StructureElement = listPillar.Item(1)
                        Dim userPillar As Pillar = dataPillar.getPillar()
                        If userPillar.Defining = True Then
                            Dim axisLinePillar As DwgLine = dataPillar.DWGEntity
                            Dim pointIntersectCollectionStart As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlineAlign3D, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos)
                            If pointIntersectCollectionStart.Count > 0 Then
                                Dim pointIntersect As Vector2D = pointIntersectCollectionStart(0)
                                Dim station As Double = 0
                                Dim offset As Double = 0
                                Dim boolStation As Boolean = projectAlignment.Plan.CompoundLine.PosToStaOffset(pointIntersect, station, offset)
                                If boolStation = True Then
                                    userBridge.startPlacementPosition = Math.Round(station, 3)
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisStructure, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBridge)
            If boolRecData = True Then
                result = True
            End If
        End If
        Return result
    End Function
End Class





