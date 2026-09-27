Imports System.ComponentModel
Imports System.Drawing
Imports System.Net
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Media.Media3D
Imports CivilEnginStructures.Pillar
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg.Bridges
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization.Geometry
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D

Public Class HandPillar
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _numberSubPillar As Integer               ' Номер подопоры
    Private _sideHand As SidePillarElement            ' Тип (Left/Right)
    Private _width As Double                          ' Толщина крыла без учета карниза
    Private _lengthTop As Double                      ' Полная длина крыла по верху (l)
    Private _lengthBottom As Double                   ' Длина горизонтальной части крыла по низу (c)
    Private _heightTop As Double                      ' Высота крыла от верха насадки (a)
    Private _heightBottom As Double                   ' Выпуск крыла вниз по торцу насадки (b)
    Private _heightTopFace As Double                  ' Высота крыла по фасаду (g)
    Private _heightBottomFace As Double               ' Длина по торцу крыла сзади (вертикальная линия по дальнему концу) f
    Private _heightCornice As Double                  ' Высота карниза h
    Private _widthCornice As Double                   ' Ширина карниза
    Private _deltaElevationSurface As Double          ' возвышение крыла над верхом проектной поверхности
    Private _boolHeightBottom As Boolean              ' вертикальная нижняя линия по торцу насадки вычисляется автоматически
    Private _deltaElevationPoint1 As Double           ' отметка верха откосного крыла у шкафной стенки
    Private _deltaElevationPoint2 As Double           ' отметка верха откосного крыла у шкафной стенки
    Private _model As String                          ' Имя модели
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _numberSubPillar = 0
        _sideHand = SidePillarElement.None
        _width = 0.0
        _lengthTop = 0.0
        _lengthBottom = 0.0
        _heightTop = 0.0
        _heightBottom = 0.0
        _heightTopFace = 0.0
        _heightBottomFace = 0.0
        _heightCornice = 0.0
        _widthCornice = 0.0
        _deltaElevationSurface = 0
        _deltaElevationPoint1 = 0
        _deltaElevationPoint1 = 0
        _elementBridgePoint = New PointsCollections
        _model = String.Empty
    End Sub
    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property NumberPillar() As Integer
        Get
            Return _numberPillar
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberPillar = value
            End If
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер элемента")>
    <Category("Свойства")>
    <DisplayName("Номер элемента")>
    <[ReadOnly](True)>
    Public Property NumberSubPillar() As Integer
        Get
            Return _numberSubPillar
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberSubPillar = value
            End If
        End Set
    End Property
    <Browsable(True)>
    <Description("Тип элемента")>
    <Category("Свойства")>
    <DisplayName("Тип элемента")>
    <[ReadOnly](True)>
    Public Property SideHand() As SidePillarElement
        Get
            Return _sideHand
        End Get
        Set(value As SidePillarElement)
            _sideHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина откосного крыла, м")>
    <Category("Свойства")>
    <DisplayName("Ширина")>
    Public Property Width() As Double
        Get
            Return _width
        End Get
        Set(value As Double)
            If value >= 0 Then
                _width = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Полная длина откосного крыла, м")>
    <Category("Свойства")>
    <DisplayName("Длина по верху")>
    Public Property LengthTop() As Double
        Get
            Return _lengthTop
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthTop = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина горизонтальной стороны крыла по низу, м")>
    <Category("Свойства")>
    <DisplayName("Длина по низу")>
    Public Property LengthBottom() As Double
        Get
            Return _lengthBottom
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthBottom = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота от верха насадки по линии шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Высота по шк.стенке")>
    Public Property HeightTop() As Double
        Get
            Return _heightTop
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightTop = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина стороны от верха насадки вниз по линии насадки, м")>
    <Category("Свойства")>
    <DisplayName("Высота по насадке")>
    Public Property HeightBottom() As Double
        Get
            Return _heightBottom
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightBottom = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина вертикальной линии торца крыла, м")>
    <Category("Свойства")>
    <DisplayName("Высота по торцу крыла")>
    Public Property HeightTopFace() As Double
        Get
            Return _heightTopFace
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightTopFace = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Превышение верха крыла над проектной поверхностью, м")>
    <Category("Свойства")>
    <DisplayName("Превышение над проектной поверхностью")>
    Public Property DeltaElevationSurface() As Double
        Get
            Return _deltaElevationSurface
        End Get
        Set(value As Double)
            _deltaElevationSurface = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Превышение над проектной поверхностью в начале крыла, м")>
    <Category("Свойства")>
    <DisplayName("Превышение в начале крыла")>
    Public Property DeltaElevationTopStartPoint() As Double
        Get
            Return _deltaElevationPoint1
        End Get
        Set(value As Double)
            _deltaElevationPoint1 = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Превышение над проектной поверхностью в конце крыла, м")>
    <Category("Свойства")>
    <DisplayName("Превышение в конце крыла")>
    Public Property DeltaElevationTopEndPoint() As Double
        Get
            Return _deltaElevationPoint2
        End Get
        Set(value As Double)
            _deltaElevationPoint2 = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Вертикальная нижняя линия по торцу насадки вычисляется автоматически")>
    <Category("Свойства")>
    <DisplayName("Вычислить торец насадки")>
    Public Property BoolHeightBottom() As Boolean
        Get
            Return _boolHeightBottom
        End Get
        Set(value As Boolean)
            _boolHeightBottom = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота карниза, м")>
    <Category("Свойства")>
    <DisplayName("Высота карниза")>
    Public Property HeightCornice() As Double
        Get
            Return _heightCornice
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightCornice = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина карниза, м")>
    <Category("Свойства")>
    <DisplayName("Ширина карниза")>
    Public Property WidthCornice() As Double
        Get
            Return _widthCornice
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthCornice = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Превышение по верху крыла, м")>
    <Category("Свойства")>
    <DisplayName("Превышение по верху крыла")>
    Public Property HeightBottomFace() As Double
        Get
            Return _heightBottomFace
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightBottomFace = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Имя модели")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    Public Property NameModel() As String
        Get
            Return _model
        End Get
        Set(value As String)
            _model = value
        End Set
    End Property

    Public Shared Function createAxisHandPillar(ByVal idBridge As String, ByVal sideElement As SidePillarElement) As StructureElement
        Dim elementHand As StructureElement = New StructureElement()
        elementHand.Label = "Мосты и путепроводы"
        elementHand.ClassBridgeObject = StructureElement.classBridge.Pillars
        If sideElement = SidePillarElement.Left Then
            elementHand.ClassObject = StructureElement.classStructure.HandLeftPillar
            elementHand.Name = StructureElement.typeObject.axisLeftHand
        Else
            elementHand.ClassObject = StructureElement.classStructure.HandRightPillar
            elementHand.Name = StructureElement.typeObject.axisRightHand
        End If
        elementHand.Description = StructureElement.GetDescription(elementHand.Name)
        elementHand.KeyParameter = ""
        elementHand.IdElement = Guid.NewGuid.ToString
        elementHand.IdStructure = idBridge
        elementHand.Note = ""
        elementHand.DWGEntity = New DwgLine()
        Return elementHand
    End Function
    'ищет откосное крыло
    Public Shared Function getHandPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal sideHand As SidePillarElement, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim dataHand As New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        Dim listAxisHand As List(Of StructureElement) = Nothing
        If sideHand = SidePillarElement.Left Then
            If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisLeftHand) = True Then
                listAxisHand = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisLeftHand)
            End If
        ElseIf sideHand = SidePillarElement.Right Then
            If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisRightHand) = True Then
                listAxisHand = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisRightHand)
            End If
        End If
        If IsNothing(listAxisHand) = False Then
            If listAxisHand.Count > 0 Then
                For k As Integer = 0 To listAxisHand.Count - 1
                    Dim tempData As StructureElement = listAxisHand.Item(k)
                    If IsNothing(tempData) = False Then
                        Dim userAxisHand As HandPillar = tempData.getHandPillar
                        If IsNothing(userAxisHand) = False Then
                            If numberPillar = userAxisHand.NumberPillar Then
                                If numberSubPillar = 0 Then numberSubPillar = userAxisHand.NumberSubPillar
                                If numberSubPillar = userAxisHand.NumberSubPillar Then
                                    dataHand.Add(tempData)
                                End If
                            End If
                        End If
                    End If
                Next k
            End If
        End If

        Return dataHand
    End Function
    'чтение данных из датагрид
    Public Shared Function readPropertiesHand(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Hand As DataGridView, Optional leftHand As Boolean = True, Optional boolBottomHeightNozzle As Boolean = False) As HandPillar
        Dim result As HandPillar = New HandPillar
        result.NumberPillar = numbPillar
        result.NumberSubPillar = numbSubPillar
        If leftHand = True Then
            result.SideHand = SidePillarElement.Left
        Else
            result.SideHand = SidePillarElement.Right
        End If
        result.BoolHeightBottom = boolBottomHeightNozzle
        If DGV_Hand.RowCount > 1 Then
            For i As Integer = 0 To DGV_Hand.RowCount - 1
                Dim tag As String = DGV_Hand.Rows(i).Tag
                Dim value As String = DGV_Hand.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_lefthand_length" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LengthTop = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_length" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LengthTop = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightTop = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightTop = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_lk" Then 'низ крыла
                        If IsNumeric(value) = True Then
                            result.LengthBottom = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_lk" Then 'низ крыла
                        If IsNumeric(value) = True Then
                            result.LengthBottom = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_h1" Then 'превышение по верху крыла
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightBottomFace = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_h1" Then 'превышение по верху крыла
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightBottomFace = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_h2" Then 'длина по торцу крыла
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightTopFace = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_h2" Then 'длина по торцу крыла
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightTopFace = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "heightBottomNozzle" Then 'длина по высоте насадки
                        If IsNumeric(value) = True Then
                            result.HeightBottom = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_lefthand_caplength" Then
                        If IsNumeric(value) = True Then
                            result.HeightCornice = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_righthand_caplength" Then
                        If IsNumeric(value) = True Then
                            result.HeightCornice = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "widthCornice" Then
                        If IsNumeric(value) = True Then
                            result.WidthCornice = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "DeltaElevationSurface" Then
                        If IsNumeric(value) = True Then
                            result.DeltaElevationSurface = Math.Round(Val(value), 3)
                        End If
                    End If
                End If
            Next i
        End If
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesHand(ByRef DGV_Hand As DataGridView) As Boolean
        If IsNothing(DGV_Hand) = True Then Return False
        If DGV_Hand.RowCount > 1 Then
            For j As Integer = 0 To DGV_Hand.RowCount - 1
                Dim tag As String = DGV_Hand.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_lefthand_length" Then
                        DGV_Hand.Rows(j).Cells(1).Value = LengthTop
                    ElseIf tag Like "bridge_lefthand_width" Then
                        DGV_Hand.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "bridge_lefthand_height" Then
                        DGV_Hand.Rows(j).Cells(1).Value = HeightTop
                    ElseIf tag Like "bridge_lefthand_lk" Then 'низ крыла
                        DGV_Hand.Rows(j).Cells(1).Value = LengthBottom
                    ElseIf tag Like "bridge_lefthand_h1" Then 'превышение по верху крыла
                        DGV_Hand.Rows(j).Cells(1).Value = HeightBottomFace
                    ElseIf tag Like "bridge_lefthand_h2" Then 'длина по торцу крыла
                        DGV_Hand.Rows(j).Cells(1).Value = HeightTopFace
                    ElseIf tag Like "heightBottomNozzle" Then 'длина по высоте насадки
                        DGV_Hand.Rows(j).Cells(1).Value = HeightBottom
                    ElseIf tag Like "bridge_lefthand_caplength" Then
                        DGV_Hand.Rows(j).Cells(1).Value = HeightCornice
                    ElseIf tag Like "widthCornice" Then
                        DGV_Hand.Rows(j).Cells(1).Value = WidthCornice
                    ElseIf tag Like "DeltaElevationSurface" Then
                        DGV_Hand.Rows(j).Cells(1).Value = DeltaElevationSurface
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Hand As DataGridView) As Boolean
        If IsNothing(DGV_Hand) = True Then Return False
        If DGV_Hand.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation1 As Boolean = False
                Dim boolTopElevation2 As Boolean = False
                Dim booldeltaH1 As Boolean = False
                Dim booldeltaH2 As Boolean = False
                For i As Integer = 0 To DGV_Hand.RowCount - 1
                    Dim oldTag As String = DGV_Hand.Rows(i).Tag
                    If oldTag Like "calc-ElevationTopStartPoint" Then
                        DGV_Hand.Rows(i).Cells(1).Value = _elementBridgePoint.StartAxisPoint.Z
                        boolTopElevation1 = True
                    ElseIf oldTag Like "calc-ElevationTopEndPoint" Then
                        DGV_Hand.Rows(i).Cells(1).Value = _elementBridgePoint.EndAxisPoint.Z
                        boolTopElevation2 = True
                    ElseIf oldTag Like "calc-DeltaElevationTopStartPoint" Then
                        DGV_Hand.Rows(i).Cells(1).Value = DeltaElevationTopStartPoint
                        booldeltaH1 = True
                    ElseIf oldTag Like "calc-DeltaElevationTopEndPoint" Then
                        DGV_Hand.Rows(i).Cells(1).Value = DeltaElevationTopEndPoint
                        booldeltaH2 = True
                    End If
                Next i
                If boolTopElevation1 = False Then
                    Dim numberRow As Integer = DGV_Hand.RowCount - 1
                    DGV_Hand.Rows.Insert(numberRow)
                    DGV_Hand.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Hand.Rows(numberRow).Cells(0).Value = "Отметка верха крыла у шкафной стенки, м"
                    DGV_Hand.Rows(numberRow).Cells(1).Value = _elementBridgePoint.StartAxisPoint.Z
                    DGV_Hand.Rows(numberRow).Tag = "calc-ElevationTopStartPoint"
                End If
                If boolTopElevation1 = False Then
                    Dim numberRow As Integer = DGV_Hand.RowCount - 1
                    DGV_Hand.Rows.Insert(numberRow)
                    DGV_Hand.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Hand.Rows(numberRow).Cells(0).Value = "Отметка верха крыла, м"
                    DGV_Hand.Rows(numberRow).Cells(1).Value = _elementBridgePoint.EndAxisPoint.Z
                    DGV_Hand.Rows(numberRow).Tag = "calc-ElevationTopEndPoint"
                End If
                If booldeltaH1 = False Then
                    Dim numberRow As Integer = DGV_Hand.RowCount - 1
                    DGV_Hand.Rows.Insert(numberRow)
                    DGV_Hand.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Hand.Rows(numberRow).Cells(0).Value = "Превышение крыла над поверхностью у шкафной стенки, м"
                    DGV_Hand.Rows(numberRow).Cells(1).Value = DeltaElevationTopStartPoint
                    DGV_Hand.Rows(numberRow).Tag = "calc-DeltaElevationTopStartPoint"
                End If
                If booldeltaH2 = False Then
                    Dim numberRow As Integer = DGV_Hand.RowCount - 1
                    DGV_Hand.Rows.Insert(numberRow)
                    DGV_Hand.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Hand.Rows(numberRow).Cells(0).Value = "Превышение крыла над поверхностью, м"
                    DGV_Hand.Rows(numberRow).Cells(1).Value = DeltaElevationTopEndPoint
                    DGV_Hand.Rows(numberRow).Tag = "calc-DeltaElevationTopEndPoint"
                End If
            End If
        End If
        Return True
    End Function
    'расчет откосного крыла
    Public Function calculateHand(ByVal userNozzle As NozzlePillar, ByVal userCabinetWall As CabinetWallPillar, Optional projectSurface As Surface = Nothing) As Boolean
        Dim pointModelNozzle As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        Dim pointModelCabinetWall As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        Dim ListPointModel As New Dictionary(Of Integer, PointStructure)
        If IsNothing(userNozzle) = False Then
            If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
                pointModelNozzle = userNozzle._elementBridgePoint.ListPointModel
            End If
        End If
        If IsNothing(userCabinetWall) = False Then
            If userCabinetWall._elementBridgePoint.ListPointModel.Count > 3 Then
                pointModelCabinetWall = userCabinetWall._elementBridgePoint.ListPointModel
            End If
        End If
        If pointModelNozzle.Count = 0 Then Return False
        If pointModelCabinetWall.Count = 0 Then Return False
        'линиии построения
        Dim shortLineNozzle As DwgLine = New DwgLine 'короткая сторона (по ширине насадки)
        Dim shortLineCabinetWall As DwgLine = New DwgLine 'короткая сторона (по ширина шкафной стенки)
        Dim bottomLongLeftLine As DwgLine = New DwgLine() 'линия по задней нижней стороне насадки
        Dim longLineLeftNozzle As DwgLine = New DwgLine 'длинная сторона (сторона насадки передняя)
        Dim longLineRightNozzle As DwgLine = New DwgLine 'длинная сторона (сторона насадки по линии начала шкафной стенки)
        'точки насадки по низу
        Dim bottomPointNozzle1 As Vector3D = New Vector3D
        Dim bottomPointNozzle2 As Vector3D = New Vector3D
        'выбираем точки для левого открылка
        If SideHand = Pillar.SidePillarElement.Left Then
            'берем 4 точки насадки
            For i As Integer = 0 To pointModelNozzle.Count - 1
                Dim tempPoint As PointStructure = pointModelNozzle.ElementAt(i).Value
                If tempPoint.Code Like "middlePt1" Then
                    shortLineNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineRightNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt1" Then
                    shortLineNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineLeftNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    bottomLongLeftLine.StartPoint = New Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                ElseIf tempPoint.Code Like "middlePt2" Then
                    longLineRightNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt2" Then
                    longLineLeftNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    bottomLongLeftLine.EndPoint = New Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                End If
            Next
        ElseIf SideHand = Pillar.SidePillarElement.Right Then
            'берем 4 точки насадки
            For i As Integer = 0 To pointModelNozzle.Count - 1
                Dim tempPoint As PointStructure = pointModelNozzle.ElementAt(i).Value
                If tempPoint.Code Like "middlePt2" Then
                    shortLineNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineRightNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt2" Then
                    shortLineNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineLeftNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    bottomLongLeftLine.StartPoint = New Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                ElseIf tempPoint.Code Like "middlePt1" Then
                    longLineRightNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt1" Then
                    longLineLeftNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    bottomLongLeftLine.EndPoint = New Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                End If
            Next
        End If
        '========================================================================================================
        If SideHand = Pillar.SidePillarElement.Left Then
            For i As Integer = 0 To pointModelCabinetWall.Count - 1
                Dim tempPoint As PointStructure = pointModelCabinetWall.ElementAt(i).Value
                If tempPoint.Code Like "rightPt1" Then
                    shortLineCabinetWall.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt1" Then
                    shortLineCabinetWall.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                End If
            Next
        ElseIf SideHand = Pillar.SidePillarElement.Right Then
            For i As Integer = 0 To pointModelCabinetWall.Count - 1
                Dim tempPoint As PointStructure = pointModelCabinetWall.ElementAt(i).Value
                If tempPoint.Code Like "rightPt2" Then
                    shortLineCabinetWall.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "leftPt2" Then
                    shortLineCabinetWall.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                End If
            Next
        End If
        If shortLineCabinetWall.Length = 0 Then Return False
        Dim boolRez As Boolean = False
        '1 точка (начало шкафной стенки
        Dim pointLeft1 As Vector3D = shortLineNozzle.StartPoint
        Dim pointRight1 As Vector3D = MathFunction.FuncCalcPointInLine(shortLineNozzle.StartPoint, longLineRightNozzle.EndPoint, Width)
        '2 точка (конец шкафной стенки, край насадки
        Dim pointLeft2 As Vector3D = shortLineNozzle.EndPoint
        Dim pointRight2 As Vector3D = MathFunction.FuncCalcPointInLine(shortLineNozzle.EndPoint, longLineLeftNozzle.EndPoint, Width)
        '22 точка по низу
        'определяем положение точки 22
        Dim LeftPoint22 As Vector3D = New Vector3D(-1, -1, -1)
        Dim rightPoint22 As Vector3D = New Vector3D(-1, -1, 1)
        ' определяем вторую высоту 2 точки если она есть
        Dim heightBottomLine As Double = HeightBottom
        If BoolHeightBottom = True Then
            HeightBottom = userNozzle.SecondHeight
        End If
        If HeightBottom > 0 Then
            LeftPoint22 = MathFunction.FuncCalcPointInLine(shortLineNozzle.EndPoint, bottomLongLeftLine.StartPoint, HeightBottom)
            Dim tempRightPoint22 As Vector3D = MathFunction.FuncCalcPointInLine(bottomLongLeftLine.StartPoint, bottomLongLeftLine.EndPoint, Width)
            rightPoint22 = MathFunction.FuncCalcPointInLine(pointRight2, tempRightPoint22, HeightBottom)
        End If
        '3 точка (ее может и не быть)
        Dim pointLeft3 As Vector3D = New Vector3D(-1, -1, -1)
        Dim pointRight3 As Vector3D = New Vector3D(-1, -1, 1)
        If LengthBottom > 0 And HeightBottom > 0 Then
            pointLeft3 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeft2.Pos, shortLineNozzle.Rotation, LengthBottom)
            pointRight3 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeft3, longLineLeftNozzle.Rotation, Width)
            pointLeft3 = New Vector3D(pointLeft3, LeftPoint22.Z)
            pointRight3 = New Vector3D(pointRight3, rightPoint22.Z)
        End If
        '4 крайняя точка откосного крыла (делаем перпендикуляр от левой точки, эта линия не паралельна длинной стороне насадки)
        Dim pointLeft4 As Vector3D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeft1, shortLineNozzle.Rotation, LengthTop)
        pointLeft4 = New Vector3D(pointLeft4, pointLeft2.Z - HeightBottom + HeightBottomFace)
        Dim tempPointRight4 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointRight1, shortLineNozzle.Rotation, LengthTop)
        Dim rot As Double = shortLineNozzle.Rotation + Math.PI / 2
        If rot > 2 * Math.PI Then
            rot -= 2 * Math.PI
        End If
        Dim tempPointLeft4 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeft4, rot, 10)
        Dim pointRight4 As Vector3D = MathFunction.FuncFindLineIntersection(pointRight1, tempPointRight4, pointLeft4, tempPointLeft4, boolRez)
        pointRight4 = New Vector3D(pointRight4, pointRight2.Z - HeightBottom + HeightBottomFace)
        If boolRez = False Then
            Return False
        End If
        'назначаем высоты верхним точкам
        Dim startElev As Double = -9999
        Dim endElev As Double = -9999
        If IsNothing(projectSurface) = False Then
            If projectSurface.Triangles.Count > 0 Then
                If IsNothing(projectSurface) = False Then
                    Try
                        startElev = projectSurface.GetElevation(pointLeft1.Pos)
                    Catch ex As ArgumentOutOfRangeException
                    End Try
                    Try
                        endElev = projectSurface.GetElevation(pointLeft4.Pos)
                    Catch ex As ArgumentOutOfRangeException
                    End Try
                End If
            End If
        End If
        If BoolHeightBottom = True Then
            HeightBottom = userNozzle.SecondHeight
        End If
        '1 точка
        If HeightTop = 0 Then
            If startElev <> -9999 Then
                Dim HBottom As Double = pointLeft1.Z
                Dim HTop As Double = startElev + DeltaElevationSurface
                Dim tempHeightTop As Double = HTop - HBottom - HeightCornice
                If tempHeightTop < 0 Then
                    MsgBox("Параметры крыла не верны")
                Else
                    HeightTop = tempHeightTop
                End If
            End If
        End If
        Dim topElevPoint1 As Double = pointLeft1.Z + HeightTop
        If HeightTopFace = 0 Then
            If endElev <> 0 Then
                Dim HBottom As Double = shortLineNozzle.EndPoint.Z - HeightBottom
                Dim HTop As Double = endElev + DeltaElevationSurface
                Dim tempHeightTopFace As Double = HTop - HBottom - HeightCornice - HeightBottomFace
                If tempHeightTopFace <= 0 Then
                    MsgBox("Параметры крыла не верны")
                Else
                    HeightTopFace = tempHeightTopFace
                End If
            End If
        End If
        '2. определаем высоту последней точки
        Dim topElevPoint4 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointLeft4, shortLineNozzle.EndPoint.Z - HeightBottom + HeightBottomFace + HeightTopFace)
        '3. определяем высоту 2 точки (интерполячия межлу 1 и 4
        Dim tempTopPoint1 As Vector3D = New Vector3D(pointLeft1, topElevPoint1)
        Dim topElevPoint2 As Double = MathFunction.FuncCalcElevationByLine(tempTopPoint1, topElevPoint4, pointLeft2) 'по верху
        'определяем высоты 3 точки пересечения
        Dim topElevPoint3 As Double = 0
        If LengthBottom > 0 Then
            topElevPoint3 = MathFunction.FuncCalcElevationByLine(tempTopPoint1, topElevPoint4, pointLeft3)
        End If
        '=================================================================================================
        'записываем в библиотеку координаты точек
        Dim pointBridgeHand As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        '1. точка
        Dim countLeft As Integer = 1
        Dim x As Double = Math.Round(pointLeft1.X, 3)
        Dim y As Double = Math.Round(pointLeft1.Y, 3)
        Dim z As Double = Math.Round(topElevPoint1, 3)
        Dim h1 As Double = Math.Round(HeightTop, 3)
        Dim code As String = "leftPt1"
        Dim pointBridge As PointStructure = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '2 точка
        x = Math.Round(pointLeft2.X, 3)
        y = Math.Round(pointLeft2.Y, 3)
        z = Math.Round(topElevPoint2, 3)
        h1 = Math.Round(topElevPoint2 - pointLeft2.Z, 3)
        Dim h2 As Double = 0
        Dim dx As Double = 0
        Dim dy As Double = 0
        If HeightBottom > 0 Then
            h2 = Math.Round(topElevPoint2 - LeftPoint22.Z, 3)
        End If
        code = "leftPt01"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, -1 * h2, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '3 точка
        If LengthBottom > 0 And HeightBottom > 0 Then
            x = Math.Round(pointLeft3.X, 3)
            y = Math.Round(pointLeft3.Y, 3)
            z = Math.Round(topElevPoint3, 3)
            h1 = Math.Round(topElevPoint3 - pointLeft3.Z, 3)
            code = "leftPt02"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(countLeft, pointBridge)
            countLeft += 1
        End If
        '4 точка
        x = Math.Round(pointLeft4.X, 3)
        y = Math.Round(pointLeft4.Y, 3)
        z = Math.Round(topElevPoint4.Z, 3)
        h1 = Math.Round(HeightTopFace, 3)
        code = "leftPt2"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'в противоположном направлении
        '4 точка
        x = Math.Round(pointRight4.X, 3)
        y = Math.Round(pointRight4.Y, 3)
        z = Math.Round(topElevPoint4.Z, 3)
        h1 = Math.Round(HeightTopFace, 3)
        code = "rightPt2"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '3 точка
        If LengthBottom > 0 And HeightBottom > 0 Then
            x = Math.Round(pointRight3.X, 3)
            y = Math.Round(pointRight3.Y, 3)
            z = Math.Round(topElevPoint3, 3)
            h1 = Math.Round(topElevPoint3 - pointRight3.Z, 3)
            code = "rightPt02"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(countLeft, pointBridge)
            countLeft += 1
        End If
        '2 точка
        x = Math.Round(pointRight2.X, 3)
        y = Math.Round(pointRight2.Y, 3)
        z = Math.Round(topElevPoint2, 3)
        h1 = Math.Round(topElevPoint2 - pointRight2.Z, 3)
        h2 = 0
        dx = 0
        dy = 0
        If HeightBottom > 0 Then
            h2 = Math.Round(topElevPoint2 - rightPoint22.Z, 3)
        End If
        code = "rightPt01"
        pointBridge = New PointStructure(x, y, z, dx, dy, -1 * h1, -1 * h2, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '1 точка
        x = Math.Round(pointRight1.X, 3)
        y = Math.Round(pointRight1.Y, 3)
        z = Math.Round(topElevPoint1, 3)
        h1 = Math.Round(HeightTop, 3)
        code = "rightPt1"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        _elementBridgePoint.ListPointModel = pointBridgeHand
        'рисуем карниз
        If WidthCornice > 0 And HeightCornice > 0 Then
            Dim reverseAngle As Double = longLineLeftNozzle.Rotation + Math.PI
            If reverseAngle > Math.PI * 2 Then
                reverseAngle -= Math.PI * 2
            End If
            Dim pointCornice1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(pointRight1, pointLeft1, WidthCornice)
            Dim pointCornice2 As Vector2D = MathFunction.FuncCalcPoint2DInLine(pointRight4, pointLeft4, WidthCornice)
            'находим высоту карнизк выступающей части
            Dim startElevation As Double = topElevPoint1 + HeightCornice
            Dim endElevation As Double = topElevPoint4.Z + HeightCornice
            pointBridgeHand = New Dictionary(Of Integer, PointStructure)
            x = Math.Round(pointCornice1.X, 3)
            y = Math.Round(pointCornice1.Y, 3)
            z = Math.Round(startElevation, 3)
            h1 = Math.Round(HeightCornice, 3)
            code = "leftPt1"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(4, pointBridge)

            x = Math.Round(pointCornice2.X, 3)
            y = Math.Round(pointCornice2.Y, 3)
            z = Math.Round(endElevation, 3)
            h1 = Math.Round(HeightCornice, 3)
            code = "leftPt2"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(3, pointBridge)

            x = Math.Round(pointRight4.X, 3)
            y = Math.Round(pointRight4.Y, 3)
            z = Math.Round(endElevation, 3)
            h1 = Math.Round(HeightCornice, 3)
            code = "rightPt2"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(2, pointBridge)

            x = Math.Round(pointRight1.X, 3)
            y = Math.Round(pointRight1.Y, 3)
            z = Math.Round(startElevation, 3)
            h1 = Math.Round(HeightCornice, 3)
            code = "rightPt1"
            pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointBridgeHand.Add(1, pointBridge)

            _elementBridgePoint.ListPointSecondModel = pointBridgeHand
        End If
        '=============================================================================================================
        Dim middleStartPoint As Cad.Foundation.Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(pointLeft1, pointRight1)
        Dim middleEndPoint As Cad.Foundation.Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(pointLeft4, pointRight4)
        _elementBridgePoint.StartAxisPoint = New Vector3D(middleStartPoint, topElevPoint1)
        _elementBridgePoint.EndAxisPoint = New Vector3D(middleEndPoint, topElevPoint4.Z)
        If startElev <> -9999 Then
            DeltaElevationTopStartPoint = Math.Round(_elementBridgePoint.StartAxisPoint.Z - startElev, 3)
        End If
        If endElev <> -9999 Then
            DeltaElevationTopEndPoint = Math.Round(_elementBridgePoint.StartAxisPoint.Z - endElev, 3)
        End If
        _elementBridgePoint.CenterTopPoint = _elementBridgePoint.StartAxisPoint
        Return True
    End Function
    'рисование оси откосного крыла
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineHand As DwgLine = Nothing
        Dim dataStructureHand As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getHandPillar(dictionaryBridgeElements, NumberPillar, SideHand)
        If listAxis.Count = 0 Then
            dataStructureHand = createAxisHandPillar(idBridge, SideHand)
        ElseIf listAxis.Count = 1 Then
            dataStructureHand = listAxis.Item(0)
        Else
            dataStructureHand = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataStructureHand) Then Return dataStructureHand
        axisLineHand = dataStructureHand.DWGEntity
        If IsNothing(axisLineHand) = True Then axisLineHand = New DwgLine
        If axisLineHand.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Обратный открылок (ось)")
            styleAxisNozzle.setObjectStyle(axisLineHand)
        End If
        'ось обратного открылка
        axisLineHand.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineHand.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLineHand.StartPoint.Pos - axisLineHand.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось насадки имеет нулевое значение. Насадка не построена.")
            Return dataStructureHand
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineHand) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineHand)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureHand.KeyParameter = strJson
        dataStructureHand.DWGEntity = axisLineHand
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineHand, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureHand)
        Return dataStructureHand
    End Function
    'функция получает все точки откосного крыла
    Public Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim ListPointModelHand As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
        If IsNothing(ListPointModelHand) = True Then
            Return result
        End If
        If ListPointModelHand.Count = 0 Then
            Return result
        End If
        'записываем результат
        For i As Integer = 0 To ListPointModelHand.Count - 1
            Dim pointStructure As PointStructure = ListPointModelHand.ElementAt(i).Value
            Dim code As String = pointStructure.Code
            Dim pointTop As Vector3D = New Vector3D(pointStructure.X, pointStructure.Y, pointStructure.Z)
            Dim pointBottom As Vector3D = New Vector3D(pointStructure.X - pointStructure.dx, pointStructure.Y - pointStructure.dy, pointStructure.Z + pointStructure.dz)
            Dim transformedBottomPoint As Vector3D = matrixTransform.transformUserPoint(pointBottom)
            If transformBottom.ContainsKey(code) = False Then
                Dim h2 As Double = pointStructure.dz2
                If h2 <> 0 Then
                    If code.IndexOf("left") > -1 Then
                        Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointBottom
                        projectPoint.projectPoint = transformedBottomPoint
                        transformBottom.Add(code, projectPoint)

                        Dim pointBottom2 As Vector3D = New Vector3D(pointStructure.X - pointStructure.dx, pointStructure.Y - pointStructure.dy, pointStructure.Z + pointStructure.dz2)
                        transformedBottomPoint = matrixTransform.transformUserPoint(pointBottom2)
                        projectPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointBottom2
                        projectPoint.projectPoint = transformedBottomPoint
                        transformBottom.Add(code & "2", projectPoint)
                    ElseIf code.IndexOf("right") > -1 Then
                        Dim pointBottom2 As Vector3D = New Vector3D(pointStructure.X - pointStructure.dx, pointStructure.Y - pointStructure.dy, pointStructure.Z + pointStructure.dz2)
                        Dim transformedBottomPoint2 As Vector3D = matrixTransform.transformUserPoint(pointBottom2)
                        Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointBottom2
                        projectPoint.projectPoint = transformedBottomPoint2
                        transformBottom.Add(code & "2", projectPoint)

                        projectPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointBottom
                        projectPoint.projectPoint = transformedBottomPoint
                        transformBottom.Add(code, projectPoint)
                    End If
                Else
                    Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                    projectPoint.originPoint = pointBottom
                    projectPoint.projectPoint = transformedBottomPoint
                    transformBottom.Add(code, projectPoint)
                End If
            End If
            Dim transformedTopPoint As Vector3D = matrixTransform.transformUserPoint(pointTop)
            If transformTop.ContainsKey(code) = False Then
                Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                projectPoint.originPoint = pointTop
                projectPoint.projectPoint = transformedTopPoint
                transformTop.Add(code, projectPoint)
            End If
        Next i
        result.Add(transformTop)
        result.Add(transformBottom)
        Return result
    End Function
    'функция возвращает все точки карниза
    Public Function getProjectionPointCornice(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim ListPointModelCornice As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointSecondModel
        If IsNothing(ListPointModelCornice) = True Then
            Return result
        End If
        If ListPointModelCornice.Count = 0 Then
            Return result
        End If
        'записываем результат
        For i As Integer = 0 To ListPointModelCornice.Count - 1
            Dim pointStructure As PointStructure = ListPointModelCornice.ElementAt(i).Value
            Dim code As String = pointStructure.Code
            Dim pointTop As Vector3D = New Vector3D(pointStructure.X, pointStructure.Y, pointStructure.Z)
            Dim pointBottom As Vector3D = New Vector3D(pointStructure.X - pointStructure.dx, pointStructure.Y - pointStructure.dy, pointStructure.Z + pointStructure.dz)
            Dim transformedBottomPoint As Vector3D = matrixTransform.transformUserPoint(pointBottom)
            If transformBottom.ContainsKey(code) = False Then
                Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                projectPoint.originPoint = pointBottom
                projectPoint.projectPoint = transformedBottomPoint
                transformBottom.Add(code, projectPoint)
            End If
            Dim transformedTopPoint As Vector3D = matrixTransform.transformUserPoint(pointTop)
            If transformTop.ContainsKey(code) = False Then
                Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                projectPoint.originPoint = pointTop
                projectPoint.projectPoint = transformedTopPoint
                transformTop.Add(code, projectPoint)
            End If
        Next i
        result.Add(transformTop)
        result.Add(transformBottom)
        Return result
    End Function
    'функция возвращает координаты точки по ее коду
    Public Function getPointByCode(ByVal code As String) As Vector3D
        Dim result As Vector3D = New Vector3D
        If IsNothing(code) = False Then
            If code.Trim.Length > 0 Then
                Dim ListPointModelHand As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
                If ListPointModelHand.Count > 0 Then
                    For i As Integer = 0 To ListPointModelHand.Count - 1
                        Dim ptStructure As PointStructure = ListPointModelHand.ElementAt(i).Value
                        If ptStructure.Code Like code Then
                            result = New Vector3D(ptStructure.X - ptStructure.dx, ptStructure.Y - ptStructure.dy, ptStructure.Z + ptStructure.dz)
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
        Return result
    End Function
End Class
