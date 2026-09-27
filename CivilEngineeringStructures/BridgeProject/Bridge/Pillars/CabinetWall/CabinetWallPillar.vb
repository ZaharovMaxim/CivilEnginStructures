Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc

Public Class CabinetWallPillar
    ' Приватные поля класса
    Private _numberPillar As Integer                              ' Номер опоры
    Private _number As Integer                                    ' Номер шкафной стенки
    Private _width As Double                                      ' Ширина шкафной стенки
    Private _leftHeight As Double                                 ' Высота шкафной стенки (по умолчанию)
    Private _rightHeight As Double                                ' Высота шкафной стенки (по умолчанию)
    Private _centerHeight As Double                               ' Высота шкафной стенки (по умолчанию)
    Private _lenght As Double                                     ' длина шкафной стенки
    Private _elevationOffsetProjectSurface As Double              ' Заглубление над проектной поверхностью

    Private _heightTopPl As Double                                ' Высота от низа насадки до верха плиты зуба упора
    Private _fullLengthPl As Double                               ' Полная высота зуба упора по шкафной стенке
    Private _lengthPl As Double                                   ' Высота зуба упора с противоположной стороны шкафной стенки
    Private _widthPl As Double                                    ' Ширина зуба упора
    Private _widthUPl As Double                                   ' Ширина скоса горизонтальной площадки зуба упора
    Private _heightUPl As Double                                  ' Высота скоса горизонтальной площадки зуба упора
    Private _heightPl As Double                                   ' Высота скоса низа зуба упора
    Private _fixedHeight As Boolean                               ' Зафиксировать высоту шкафной стенки
    Private _elevationTopPlate As Double                          ' Отметка верха плиты
    Private _model As String
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _number = 0
        _width = 0.0
        _leftHeight = 0.0
        _rightHeight = 0.0
        _centerHeight = 0.0
        _lenght = 0
        _elevationOffsetProjectSurface = 0.0
        _elevationTopPlate = 0.0
        _heightTopPl = 0.0
        _fullLengthPl = 0.0
        _lengthPl = 0.0
        _widthPl = 0.0
        _widthUPl = 0
        _heightUPl = 0
        _heightPl = 0
        _fixedHeight = False
        _model = ""
        _elementBridgePoint = New PointsCollections
    End Sub

    <Browsable(False)>
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

    <Browsable(False)>
    <Description("Номер элемента")>
    <Category("Свойства")>
    <DisplayName("Номер элемента")>
    <[ReadOnly](True)>
    Public Property Number() As Integer
        Get
            Return _number
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _number = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина шкафной стенки, м")>
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
    <Description("Высота шкафной стенки слева, м")>
    <Category("Свойства")>
    <DisplayName("Высота слева")>
    Public Property LeftHeight() As Double
        Get
            Return _leftHeight
        End Get
        Set(value As Double)
            If value >= 0 Then
                _leftHeight = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота шкафной стенки справа, м")>
    <Category("Свойства")>
    <DisplayName("Высота справа")>
    Public Property RightHeight() As Double
        Get
            Return _rightHeight
        End Get
        Set(value As Double)
            If value >= 0 Then
                _rightHeight = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота шкафной стенки по центру, м")>
    <Category("Свойства")>
    <DisplayName("Высота по центру")>
    Public Property CenterHeight() As Double
        Get
            Return _centerHeight
        End Get
        Set(value As Double)
            If value >= 0 Then
                _centerHeight = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property Lenght() As Double
        Get
            Return _lenght
        End Get
        Set(value As Double)
            _lenght = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Заглубление над проектной поверхностью, м")>
    <Category("Свойства")>
    <DisplayName("Заглубление")>
    Public Property ElevationOffsetProjectSurface() As Double
        Get
            Return _elevationOffsetProjectSurface
        End Get
        Set(value As Double)
            _elevationOffsetProjectSurface = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота от верха насадки до верха зуба упора, м")>
    <Category("Свойства")>
    <DisplayName("Высота до зуба упора")>
    Public Property HeightTopPl() As Double
        Get
            Return _heightTopPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightTopPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Весота зуба упора по линии шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Прлная высота зуба упора")>
    Public Property FullLengthPl() As Double
        Get
            Return _fullLengthPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _fullLengthPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота нижнего скоса зуба упора с противоположной стороны шкафной стенки")>
    <Category("Свойства")>
    <DisplayName("Высота скоса зуба упора")>
    Public Property HeightPl() As Double
        Get
            Return _heightPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота зуба упора с противоположной стороны шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Высота зуба упора")>
    Public Property LengthPl() As Double
        Get
            Return _lengthPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина зубца под переходные плиты, м")>
    <Category("Свойства")>
    <DisplayName("Ширина зубца")>
    Public Property WidthUPl() As Double
        Get
            Return _widthUPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthUPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота зубца под переходные плиты, м")>
    <Category("Свойства")>
    <DisplayName("Высота зубца")>
    Public Property HeightUPl() As Double
        Get
            Return _heightUPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightUPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина площадки зуба упора под переходные плиты, м")>
    <Category("Свойства")>
    <DisplayName("Ширина зуба упора")>
    Public Property WidthPl() As Double
        Get
            Return _widthPl
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthPl = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха зуба упора, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха зуба упора")>
    Public Property ElevationTopPlate() As Double
        Get
            Return _elevationTopPlate
        End Get
        Set(value As Double)
            _elevationTopPlate = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Фиксированная высота")>
    <Category("Свойства")>
    <DisplayName("Фиксированная высота")>
    Public Property FixedHeight() As Boolean
        Get
            Return _fixedHeight
        End Get
        Set(value As Boolean)
            _fixedHeight = value
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
    Public Shared Function createAxisCabinetWall(ByVal idBridge As String) As StructureElement
        Dim elementCabinetWall As StructureElement = New StructureElement()
        elementCabinetWall.Label = "Мосты и путепроводы"
        elementCabinetWall.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementCabinetWall.ClassObject = StructureElement.classStructure.CabinetWallPillar
        elementCabinetWall.Name = StructureElement.typeObject.axisCabinetWall
        elementCabinetWall.Description = "Шкафная стенка (ось)"
        elementCabinetWall.KeyParameter = ""
        elementCabinetWall.IdElement = Guid.NewGuid.ToString
        elementCabinetWall.IdStructure = idBridge
        elementCabinetWall.Note = ""
        elementCabinetWall.DWGEntity = New DwgLine()
        Return elementCabinetWall
    End Function
    'ищет шкафную стенку
    Public Shared Function getAxis(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim dataCabinetWall As New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisCabinetWall) = True Then
            Dim listAxisCabinetWall = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisCabinetWall)
            If IsNothing(listAxisCabinetWall) = False Then
                If listAxisCabinetWall.Count > 0 Then
                    For k As Integer = 0 To listAxisCabinetWall.Count - 1
                        Dim tempData As StructureElement = listAxisCabinetWall.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisCabinetWall As CabinetWallPillar = tempData.getCabinetWallPillar
                            If IsNothing(userAxisCabinetWall) = False Then
                                If numberPillar = userAxisCabinetWall.NumberPillar Then
                                    numberSubPillar = userAxisCabinetWall.Number
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisCabinetWall.Number
                                    If numberSubPillar = userAxisCabinetWall.Number Then
                                        dataCabinetWall.Add(tempData)
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataCabinetWall
    End Function
    'чтение данных в датогрид
    Public Shared Function readPropertiesCabinetWall(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_CabinetWall As DataGridView, Optional ByVal fixedHeight As Boolean = False) As CabinetWallPillar
        Dim result As CabinetWallPillar = New CabinetWallPillar
        result.NumberPillar = numbPillar
        result.Number = numbSubPillar
        If DGV_CabinetWall.RowCount > 1 Then
            For i As Integer = 0 To DGV_CabinetWall.RowCount - 1
                Dim tag As String = DGV_CabinetWall.Rows(i).Tag
                Dim value As String = DGV_CabinetWall.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_cabwall_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "offsetProjectSurface" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.ElevationOffsetProjectSurface = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "bridge_cabwall_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.CenterHeight = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "leftHeight" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LeftHeight = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "rightHeight" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.RightHeight = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "heightTopPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightTopPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "fullLenghtPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.FullLengthPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "lenghtPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LengthPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "widthPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.WidthPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "widthUPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.WidthUPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "heightUPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightUPl = Math.Round(Val(value), 3)
                        End If
                    ElseIf tag Like "heightPl" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightPl = Math.Round(Val(value), 3)
                        End If
                    End If
                End If
            Next i
        End If
        If fixedHeight = True Then
            result.FixedHeight = True
        End If
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesCabinetWall(ByRef DGV_CabinetWall As DataGridView) As Boolean
        If IsNothing(DGV_CabinetWall) = True Then Return False
        If DGV_CabinetWall.RowCount > 1 Then
            For j As Integer = 0 To DGV_CabinetWall.RowCount - 1
                Dim tag As String = DGV_CabinetWall.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_cabwall_width" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "offsetProjectSurface" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = ElevationOffsetProjectSurface
                    ElseIf tag Like "bridge_cabwall_height" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = CenterHeight
                    ElseIf tag Like "leftHeight" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = LeftHeight
                    ElseIf tag Like "rightHeight" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = RightHeight
                    ElseIf tag Like "heightTopPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = HeightTopPl
                    ElseIf tag Like "fullLenghtPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = FullLengthPl
                    ElseIf tag Like "lenghtPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = LengthPl
                    ElseIf tag Like "widthPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = WidthPl
                    ElseIf tag Like "widthUPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = WidthUPl
                    ElseIf tag Like "heightPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = HeightPl
                    ElseIf tag Like "heightUPl" Then
                        DGV_CabinetWall.Rows(j).Cells(1).Value = HeightUPl
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_CabinetWall As DataGridView) As Boolean
        If IsNothing(DGV_CabinetWall) = True Then Return False
        If DGV_CabinetWall.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolLenght As Boolean = False
                Dim boolLeftHeight As Boolean = False
                Dim boolRightHeight As Boolean = False
                Dim boolCenterHeight As Boolean = False
                For i As Integer = 0 To DGV_CabinetWall.RowCount - 1
                    Dim oldTag As String = DGV_CabinetWall.Rows(i).Tag
                    If oldTag Like "calc-ElevationTopPlate" Then
                        DGV_CabinetWall.Rows(i).Cells(1).Value = ElevationTopPlate
                        boolTopElevation = True
                    ElseIf oldTag Like "calc-Lenght" Then
                        DGV_CabinetWall.Rows(i).Cells(1).Value = Lenght
                        boolLenght = True
                    ElseIf oldTag Like "leftHeight" Then
                        DGV_CabinetWall.Rows(i).Cells(1).Value = LeftHeight
                        DGV_CabinetWall.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                        boolLeftHeight = True
                    ElseIf oldTag Like "rightHeight" Then
                        DGV_CabinetWall.Rows(i).Cells(1).Value = RightHeight
                        DGV_CabinetWall.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                        boolRightHeight = True
                    ElseIf oldTag Like "bridge_cabwall_height" Then
                        DGV_CabinetWall.Rows(i).Cells(1).Value = CenterHeight
                        DGV_CabinetWall.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                        boolCenterHeight = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_CabinetWall.RowCount - 1
                    DGV_CabinetWall.Rows.Insert(numberRow)
                    DGV_CabinetWall.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_CabinetWall.Rows(numberRow).Cells(0).Value = "Отметка зуба упора, м"
                    DGV_CabinetWall.Rows(numberRow).Cells(1).Value = ElevationTopPlate
                    DGV_CabinetWall.Rows(numberRow).Tag = "calc-ElevationTopPlate"
                End If
                If boolLenght = False Then
                    Dim numberRow As Integer = DGV_CabinetWall.RowCount - 1
                    DGV_CabinetWall.Rows.Insert(numberRow)
                    DGV_CabinetWall.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_CabinetWall.Rows(numberRow).Cells(0).Value = "Полная длина, м"
                    DGV_CabinetWall.Rows(numberRow).Cells(1).Value = Lenght
                    DGV_CabinetWall.Rows(numberRow).Tag = "calc-Lenght"
                End If
            End If
        End If
        Return True
    End Function
    'функция находит минимальное расстояние от балки до линии шкафной стенки
    Public Shared Function getPositionCabinetWall(ByVal axisLinePillar As DwgLine, ByVal listBeamsPillar As List(Of Dictionary(Of Integer, StructureElement)), ByVal firstPillar As Boolean) As Double
        'находим минимальный зазор между балками и шкафной стенкой и строим линию шкафной стенки и находит точки по контуру насадки
        Dim minClearance As Double = 0
        If IsNothing(axisLinePillar) = True Then Return 0
        If axisLinePillar.Length = 0 Then Return 0
        If IsNothing(listBeamsPillar) = True Then Return 0
        Dim dictBeamsPillar As Dictionary(Of Integer, StructureElement) = Nothing
        If firstPillar = True Then
            dictBeamsPillar = listBeamsPillar.Item(1)
        Else
            dictBeamsPillar = listBeamsPillar.Item(0)
        End If
        If IsNothing(dictBeamsPillar) = False Then
            If dictBeamsPillar.Count > 0 Then
                For i As Integer = 0 To dictBeamsPillar.Count - 1
                    Dim dataBeam As StructureElement = dictBeamsPillar.ElementAt(i).Value
                    If IsNothing(dataBeam) = True Then Continue For
                    Dim axisLineBeam As DwgLine = dataBeam.DWGEntity
                    If IsNothing(axisLineBeam) = True Then Continue For
                    If axisLineBeam.Length = 0 Then
                        Continue For
                    End If
                    Dim userBeam As BeamI = dataBeam.getBeamI()
                    If IsNothing(userBeam) = False Then
                        Dim startPointElements As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
                        Dim endPointElements As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
                        Dim boolRestoreElements As Boolean = CalculationBeams.restoreElementsBeam(axisLineBeam, userBeam, startPointElements, endPointElements, CalculationBeams.rectoreBeam.fullBeam)
                        If startPointElements.Count = 4 And endPointElements.Count = 4 Then
                            'ищем пересечения с осью первой или последней опоры
                            If boolRestoreElements = True Then
                                Dim boolRez1 As Boolean = False
                                Dim ptIntersect1 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(0).Pos, startPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez1)
                                Dim boolRez2 As Boolean = False
                                Dim ptIntersect2 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(1).Pos, startPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez2)
                                Dim boolRez3 As Boolean = False
                                Dim ptIntersect3 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(2).Pos, startPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez3)
                                Dim boolRez4 As Boolean = False
                                Dim ptIntersect4 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(3).Pos, startPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez4)
                                If firstPillar = True Then
                                    Dim distZazor1 As Double = (ptIntersect1 - startPointElements(0).Pos).Length
                                    Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(startPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                                    If distZazor1 > minClearance Then
                                        minClearance = distZazor1
                                    End If
                                    Dim distZazor2 As Double = (ptIntersect2 - startPointElements(1).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                                    If distZazor2 > minClearance Then
                                        minClearance = distZazor2
                                    End If
                                    Dim distZazor3 As Double = (ptIntersect3 - startPointElements(2).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                                    If distZazor3 > minClearance Then
                                        minClearance = distZazor3
                                    End If
                                    Dim distZazor4 As Double = (ptIntersect4 - startPointElements(3).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                                    If distZazor4 > minClearance Then
                                        minClearance = distZazor4
                                    End If
                                Else
                                    Dim distZazor1 As Double = (ptIntersect1 - endPointElements(0).Pos).Length
                                    Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(endPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                                    If distZazor1 > minClearance Then
                                        minClearance = distZazor1
                                    End If
                                    Dim distZazor2 As Double = (ptIntersect2 - endPointElements(1).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                                    If distZazor2 > minClearance Then
                                        minClearance = distZazor2
                                    End If
                                    Dim distZazor3 As Double = (ptIntersect3 - endPointElements(2).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                                    If distZazor3 > minClearance Then
                                        minClearance = distZazor3
                                    End If
                                    Dim distZazor4 As Double = (ptIntersect4 - endPointElements(3).Pos).Length
                                    pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                                    If distZazor4 > minClearance Then
                                        minClearance = distZazor4
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return minClearance
    End Function
    'функции для расчета начального положения правой стороны шкафной стенки
    Public Shared Function calculateLineCabinetWall(ByVal axisLinePillar As DwgLine, ByVal clearanceBeam As Double, ByVal dataLeftBeam As StructureElement, ByVal dataRightBeam As StructureElement, ByVal numberPillar As Integer, ByVal userNozzle As NozzlePillar) As DwgLine
        'восстанавливаем линию шкафной стенки
        If IsNothing(axisLinePillar) = True Then Return Nothing
        If axisLinePillar.Length = 0 Then Return Nothing
        Dim tempLineCabinetWall As DwgLine = New DwgLine()
        If numberPillar = 1 Then
            'делаем смещение вправо против хода пикетажа
            Dim coolEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            axisLinePillar.Offset(coolEnt, clearanceBeam)
            If coolEnt.Count > 0 Then
                tempLineCabinetWall = coolEnt(0)
            Else
                Return Nothing
            End If
        Else
            Dim coolEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            'смещение влево по ходу пикетажа
            axisLinePillar.Offset(coolEnt, -1 * clearanceBeam)
            If coolEnt.Count > 0 Then
                tempLineCabinetWall = coolEnt(0)
            Else
                Return Nothing
            End If
        End If
        'находим пересечение крайних балок и линии шкафной стенки
        If IsNothing(userNozzle) = True Then Return Nothing
        'расчет смещения оси левой балки
        If IsNothing(dataLeftBeam) = True Then Return Nothing
        Dim axisLeftBeam As DwgLine = dataLeftBeam.DWGEntity
        If IsNothing(axisLeftBeam) = True Then Return Nothing
        If axisLeftBeam.Length = 0 Then Return Nothing
        userNozzle.LeftDirection = Math.Round(axisLeftBeam.Rotation, 6)
        Dim userLeftBeam As BeamI = dataLeftBeam.getBeamI()
        If IsNothing(userLeftBeam) = True Then Return Nothing
        Dim outletNozzleLeftBeam As Double = userNozzle.OutletLeftBeam + userLeftBeam.widthTopPlateLeft  'отступ насадки слева за балку
        Dim tempLeftBeam As DwgLine = New DwgLine()
        Dim coolEntBeam As List(Of DwgEntity) = New List(Of DwgEntity)
        'смещение влево оси балки на величину вехней платформы
        axisLeftBeam.Offset(coolEntBeam, -1 * outletNozzleLeftBeam)
        If coolEntBeam.Count > 0 Then
            tempLeftBeam = coolEntBeam(0)
        Else
            Return Nothing
        End If

        'расчет смещения оси правой балки
        If IsNothing(dataRightBeam) = True Then Return Nothing
        Dim axisRightBeam As DwgLine = dataRightBeam.DWGEntity
        If IsNothing(axisRightBeam) = True Then Return Nothing
        If axisRightBeam.Length = 0 Then Return Nothing
        userNozzle.RightDirection = Math.Round(axisRightBeam.Rotation, 6)
        Dim userRightBeam As BeamI = dataRightBeam.getBeamI()
        If IsNothing(userRightBeam) = True Then Return Nothing
        Dim outletNozzleRightBeam As Double = userNozzle.OutletRightBeam + userRightBeam.widthTopPlateRight  'отступ насадки справа за балку
        Dim tempRightBeam As DwgLine = New DwgLine()
        coolEntBeam = New List(Of DwgEntity)
        axisRightBeam.Offset(coolEntBeam, outletNozzleRightBeam)
        If coolEntBeam.Count > 0 Then
            tempRightBeam = coolEntBeam(0)
        Else
            Return Nothing
        End If
        'делаем расчет пересечений
        Dim boolRez As Boolean = False
        Dim ptIntersectLeft As Vector2D = MathFunction.FuncFindLineIntersection(tempLeftBeam.StartPoint.Pos, tempLeftBeam.EndPoint.Pos, tempLineCabinetWall.StartPoint.Pos, tempLineCabinetWall.EndPoint.Pos, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать линию Шкафной стенки.")
            Return Nothing
        End If
        boolRez = False
        Dim ptIntersectRight As Vector2D = MathFunction.FuncFindLineIntersection(tempRightBeam.StartPoint.Pos, tempRightBeam.EndPoint.Pos, tempLineCabinetWall.StartPoint.Pos, tempLineCabinetWall.EndPoint.Pos, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать линию Шкафной стенки.")
            Return Nothing
        End If
        tempLineCabinetWall.StartPoint = New Cad.Foundation.Vector3D(ptIntersectLeft, userNozzle.TopElevation)
        tempLineCabinetWall.EndPoint = New Cad.Foundation.Vector3D(ptIntersectRight, userNozzle.TopElevation)

        Return tempLineCabinetWall
    End Function
    'функция делает расчет шкафной стенки
    Public Function calculateCabinetWall(ByVal userNozzle As NozzlePillar, ByVal userLeftHand As HandPillar, ByVal userRightHand As HandPillar, Optional ByVal projectSurface As Surface = Nothing, Optional ByVal projectAlignment As Alignment = Nothing) As Boolean
        Dim dictionaryPointNozzle As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        If IsNothing(userNozzle) = False Then
            dictionaryPointNozzle = userNozzle._elementBridgePoint.ListPointModel
        End If
        If IsNothing(dictionaryPointNozzle) = True Then Return False
        If dictionaryPointNozzle.Count = 0 Then Return False
        Dim leftLineCabinetWall As DwgLine = New DwgLine
        Dim lineDirLeft As DwgLine = New DwgLine
        Dim lineDirRight As DwgLine = New DwgLine
        For i As Integer = 0 To dictionaryPointNozzle.Count - 1
            Dim tempPoint As PointStructure = dictionaryPointNozzle.ElementAt(i).Value
            If tempPoint.Code Like "middlePt1" Then
                leftLineCabinetWall.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                lineDirLeft.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
            ElseIf tempPoint.Code Like "middlePt2" Then
                leftLineCabinetWall.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                lineDirRight.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
            ElseIf tempPoint.Code Like "leftPt1" Then
                lineDirLeft.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
            ElseIf tempPoint.Code Like "leftPt2" Then
                lineDirRight.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
            End If
        Next
        If leftLineCabinetWall.Length = 0 Then Return False
        If lineDirLeft.Length = 0 Then Return True
        If lineDirRight.Length = 0 Then Return True
        Dim leftDirection As Double = lineDirLeft.Rotation
        Dim rightDirection As Double = lineDirRight.Rotation
        Dim pointModelCabinetWall As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        '=============================================================================================
        'крайняя линия насадки должна быть
        If leftLineCabinetWall.Length > 0 Then
            'толщина открылков
            Dim widthLeftHand As Double = Math.Round(userLeftHand.Width, 3)
            Dim widthRightHand As Double = Math.Round(userRightHand.Width, 3)
            'уменьшаем линию насадки на величину толщины открылков
            Dim boolExt As Boolean = BridgeGeometry.extendLine(leftLineCabinetWall, -1 * widthLeftHand, -1 * widthRightHand)
            Lenght = Math.Round(leftLineCabinetWall.Length, 3)
            'проводим вспомогательные линии справа и слева
            Dim tempIntersectPointLeft As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftLineCabinetWall.StartPoint.Pos, leftDirection, 10)
            Dim tempIntersectPointRight As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftLineCabinetWall.EndPoint.Pos, rightDirection, 10)
            'делаем перенос линии шкафной стенки на ширину шкафной стенки
            Dim offsetPositionCabinetWall As Double = Width
            Dim coolEntCabWal As List(Of DwgEntity) = New List(Of DwgEntity)
            Dim rightLineCabinetWall As DwgLine = New DwgLine
            If NumberPillar = 1 Then
                leftLineCabinetWall.Offset(coolEntCabWal, offsetPositionCabinetWall)
                If coolEntCabWal.Count > 0 Then
                    rightLineCabinetWall = coolEntCabWal(0)
                End If
            Else
                leftLineCabinetWall.Offset(coolEntCabWal, -1 * offsetPositionCabinetWall)
                If coolEntCabWal.Count > 0 Then
                    rightLineCabinetWall = coolEntCabWal(0)
                End If
            End If
            'ищем пересечение линий с временными точками
            Dim boolRez As Boolean = False
            Dim ptIntersectLeft As Vector2D = MathFunction.FuncFindLineIntersection(rightLineCabinetWall.EndPoint.Pos, rightLineCabinetWall.StartPoint.Pos, leftLineCabinetWall.StartPoint.Pos, tempIntersectPointLeft, boolRez)
            If boolRez = False Then Return False
            Dim ptIntersectRight As Vector2D = MathFunction.FuncFindLineIntersection(rightLineCabinetWall.StartPoint.Pos, rightLineCabinetWall.EndPoint.Pos, leftLineCabinetWall.EndPoint.Pos, tempIntersectPointRight, boolRez)
            If boolRez = False Then Return False
            rightLineCabinetWall.StartPoint = New Cad.Foundation.Vector3D(ptIntersectLeft, leftLineCabinetWall.StartPoint.Z)
            rightLineCabinetWall.EndPoint = New Cad.Foundation.Vector3D(ptIntersectRight, leftLineCabinetWall.StartPoint.Z)
            '===================================================================================================================
            'вычисляем высоту начальной точки по проектой поверхности
            Dim startElev As Double = leftLineCabinetWall.StartPoint.Z + CenterHeight
            Dim endElev As Double = leftLineCabinetWall.EndPoint.Z + CenterHeight
            Dim centerLeftElev As Double = 0
            Dim centerRightElev As Double = 0
            Dim centerLeftPoint As Vector3D = New Vector3D(0, 0, 0) 'центр шкафной стенки
            Dim centerRightPoint As Vector3D = New Vector3D(0, 0, 0) 'центр шкафной стенки
            'если задана фиксированная высота вычисляем отметки верха 3 точек шкафной стенки, лево, право и центр
            Dim axisPline3D As Polyline3D = New Polyline3D
            If FixedHeight = True Then
                startElev = leftLineCabinetWall.StartPoint.Z + LeftHeight
                endElev = leftLineCabinetWall.EndPoint.Z + RightHeight
                projectAlignment.Plan.CompoundLine.ToPolyLine(axisPline3D)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, leftLineCabinetWall.StartPoint.Pos, leftLineCabinetWall.EndPoint.Pos)
                Dim pointIntersectRightCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, rightLineCabinetWall.StartPoint.Pos, rightLineCabinetWall.EndPoint.Pos)
                If pointIntersectCollection.Count = 0 Then
                    centerLeftPoint = MathFunction.funcCalcMiddleCoordByToPoints3d(leftLineCabinetWall.StartPoint, leftLineCabinetWall.EndPoint)
                    centerLeftElev = centerLeftPoint.Z + CenterHeight
                    centerRightPoint = MathFunction.funcCalcMiddleCoordByToPoints3d(rightLineCabinetWall.StartPoint, rightLineCabinetWall.EndPoint)
                    centerRightElev = centerRightPoint.Z + CenterHeight
                Else
                    centerLeftPoint = pointIntersectCollection.ElementAt(0)
                    Dim tempElev As Double = MathFunction.FuncCalcElevationByLine(leftLineCabinetWall.StartPoint, leftLineCabinetWall.EndPoint, centerLeftPoint.Pos)
                    centerLeftPoint = New Vector3D(centerLeftPoint.Pos, tempElev + CenterHeight)
                    If pointIntersectRightCollection.Count > 0 Then
                        centerRightPoint = pointIntersectRightCollection.ElementAt(0)
                        Dim tempElev2 As Double = MathFunction.FuncCalcElevationByLine(rightLineCabinetWall.StartPoint, rightLineCabinetWall.EndPoint, centerRightPoint.Pos)
                        centerRightPoint = New Vector3D(centerRightPoint.Pos, tempElev2 + CenterHeight)
                    End If
                End If
            End If
            'делаем вычисление высоты шкафной стенки относительно проектной поверхности
            Dim arrayPointSlope As Double(,) = {}
            Dim countArrayPointSlope As Integer = 0
            If FixedHeight = False Then
                If IsNothing(projectSurface) = True Then
                    MsgBox("Проектная поверхность не задана.")
                    Return False
                End If
                If projectSurface.Points.Count = 0 Then
                    MsgBox("Проектная поверхность нулевой площади.")
                    Return False
                End If
                Try
                    startElev = projectSurface.GetElevation(leftLineCabinetWall.StartPoint.Pos) - ElevationOffsetProjectSurface
                    endElev = projectSurface.GetElevation(leftLineCabinetWall.EndPoint.Pos) - ElevationOffsetProjectSurface
                    LeftHeight = Math.Round(startElev, 3)
                    RightHeight = Math.Round(endElev, 3)
                Catch ex As System.NullReferenceException
                    MsgBox("Не удалось получить отметку поверхности в начале либо в конце шкафной стенки!!! Построение может быть не корректным.")
                End Try
                'записываем первую точку шкафной стенки
                ReDim Preserve arrayPointSlope(4, 0)
                arrayPointSlope(0, 0) = Math.Round(leftLineCabinetWall.StartPoint.X, 3)
                arrayPointSlope(1, 0) = Math.Round(leftLineCabinetWall.StartPoint.Y, 3)
                arrayPointSlope(2, 0) = Math.Round(startElev, 3)
                arrayPointSlope(3, 0) = 0
                countArrayPointSlope += 1
                Dim listPoly As List(Of Vector2D) = New List(Of Vector2D)
                listPoly.Add(leftLineCabinetWall.StartPoint.Pos)
                listPoly.Add(leftLineCabinetWall.EndPoint.Pos)
                'находим сечение по поверхности для левой стороны шкафной стенки
                Dim sectCabVal As Sfc.Sections.Section = projectSurface.CreateSection(listPoly, Sfc.Sections.SectionFlags.FilterRibs)
                If sectCabVal.Count > 1 Then
                    Dim numberElement As Integer = sectCabVal.Count * 2
                    'заполняем словарь пустыми значениями
                    For i As Integer = 0 To numberElement - 1
                        pointModelCabinetWall.Add(i + 1, New PointStructure(0, 0, 0, 0, 0, 0, 0, ""))
                    Next i
                    For i As Integer = 0 To sectCabVal.Count - 1
                        Dim sect As Sfc.Sections.SectionNode = sectCabVal.ElementAt(i)
                        Dim dist As Double = Math.Round(sect.Vertex.X, 3)
                        Dim elev As Double = Math.Round(sect.Vertex.Y, 3)
                        If dist = 0 Then Continue For
                        Dim point As Vector2D = MathFunction.FuncCalcPoint2DInLine(leftLineCabinetWall.StartPoint.Pos, leftLineCabinetWall.EndPoint.Pos, dist)
                        Dim prevElev As Double = arrayPointSlope(2, countArrayPointSlope - 1)
                        Dim iLine As Double = Math.Round((elev - prevElev) / dist, 3)
                        ReDim Preserve arrayPointSlope(4, countArrayPointSlope)
                        arrayPointSlope(0, countArrayPointSlope) = point.X
                        arrayPointSlope(1, countArrayPointSlope) = point.Y
                        arrayPointSlope(2, countArrayPointSlope) = elev
                        arrayPointSlope(3, countArrayPointSlope) = iLine
                        countArrayPointSlope += 1
                        If i = sectCabVal.Count / 2 - 1 Then
                            'высота стенки в середине
                            Dim elevMiddleCabinetVall As Double = MathFunction.FuncCalcElevationByLine(leftLineCabinetWall.StartPoint, leftLineCabinetWall.EndPoint, point)
                            CenterHeight = Math.Round(elev - elevMiddleCabinetVall, 3)
                        End If
                    Next i
                    'записываем последний уклон
                    Dim pointPrev As Vector2D = New Vector2D(arrayPointSlope(0, countArrayPointSlope - 1), arrayPointSlope(1, countArrayPointSlope - 1))
                    Dim distPrev As Double = (pointPrev - leftLineCabinetWall.EndPoint.Pos).Length
                    distPrev = Math.Round(distPrev, 3)
                    If distPrev > 0 Then
                        Dim iLinePrev As Double = Math.Round((endElev - arrayPointSlope(2, countArrayPointSlope - 1)) / distPrev, 3)
                        ReDim Preserve arrayPointSlope(4, countArrayPointSlope)
                        arrayPointSlope(0, countArrayPointSlope) = leftLineCabinetWall.EndPoint.X
                        arrayPointSlope(1, countArrayPointSlope) = leftLineCabinetWall.EndPoint.Y
                        arrayPointSlope(2, countArrayPointSlope) = endElev
                        arrayPointSlope(3, countArrayPointSlope) = iLinePrev
                    Else
                        arrayPointSlope(0, countArrayPointSlope - 1) = leftLineCabinetWall.EndPoint.X
                        arrayPointSlope(1, countArrayPointSlope - 1) = leftLineCabinetWall.EndPoint.Y
                    End If
                End If
            Else
                'заполняем словарь пустыми значениями 6 элемента (2элемента начала, 2 середины и 2 конца)
                For i As Integer = 0 To 5
                    pointModelCabinetWall.Add(i + 1, New PointStructure(0, 0, 0, 0, 0, 0, 0, ""))
                Next i
            End If
            '==============================================================================================================
            'заполняем крайние элементы
            Dim countElements As Integer = pointModelCabinetWall.Count
            'левая начальная крайняя точка
            Dim x As Double = Math.Round(rightLineCabinetWall.StartPoint.X, 3)
            Dim y As Double = Math.Round(rightLineCabinetWall.StartPoint.Y, 3)
            Dim z As Double = Math.Round(startElev, 3)
            Dim h1 As Double = Math.Round(startElev - rightLineCabinetWall.StartPoint.Z, 3)
            Dim code As String = "leftPt1"
            Dim pointBidge As PointStructure = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointModelCabinetWall.Item(1) = pointBidge

            'правая конечная крайняя точка
            x = Math.Round(rightLineCabinetWall.EndPoint.X, 3)
            y = Math.Round(rightLineCabinetWall.EndPoint.Y, 3)
            z = Math.Round(endElev, 3)
            h1 = Math.Round(endElev - rightLineCabinetWall.EndPoint.Z, 3)
            code = "leftPt2"
            pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointModelCabinetWall.Item(countElements / 2) = pointBidge

            'левая конечная крайняя точка
            x = Math.Round(leftLineCabinetWall.EndPoint.X, 3)
            y = Math.Round(leftLineCabinetWall.EndPoint.Y, 3)
            z = Math.Round(endElev, 3)
            h1 = Math.Round(endElev - leftLineCabinetWall.EndPoint.Z, 3)
            code = "rightPt2"
            pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointModelCabinetWall.Item(countElements / 2 + 1) = pointBidge

            x = Math.Round(leftLineCabinetWall.StartPoint.X, 3)
            y = Math.Round(leftLineCabinetWall.StartPoint.Y, 3)
            z = Math.Round(startElev, 3)
            h1 = Math.Round(startElev - leftLineCabinetWall.StartPoint.Z, 3)
            code = "rightPt1"
            pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
            pointModelCabinetWall.Item(countElements) = pointBidge
            'если высота фиксированная, заполняем еще центральную точку
            If FixedHeight = True Then
                '
                x = Math.Round(centerLeftPoint.X, 3)
                y = Math.Round(centerLeftPoint.Y, 3)
                z = Math.Round(centerLeftPoint.Z, 3)
                h1 = Math.Round(CenterHeight, 3)
                code = "rightPt0"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Item(5) = pointBidge

                x = Math.Round(centerRightPoint.X, 3)
                y = Math.Round(centerRightPoint.Y, 3)
                z = Math.Round(centerRightPoint.Z, 3)
                h1 = Math.Round(CenterHeight, 3)
                code = "leftPt0"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Item(2) = pointBidge
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'строим контур шкафной стенки
            If arrayPointSlope.Length > 0 Then
                Dim pointLeftCabWall1 As Vector2D = leftLineCabinetWall.StartPoint.Pos
                Dim pointLeftCabWall2 As Vector2D = rightLineCabinetWall.StartPoint.Pos
                For i As Integer = 1 To arrayPointSlope.GetUpperBound(1) - 1
                    Dim topElevCabinetWall As Double = arrayPointSlope(2, i) - ElevationOffsetProjectSurface
                    Dim startPointCabWall As Vector2D = New Vector2D(arrayPointSlope(0, i), arrayPointSlope(1, i))
                    Dim iTemp1 As Double = arrayPointSlope(3, i)
                    'вычисляем конечную линию шкафной стенки
                    Dim pointCabWal As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(startPointCabWall, leftDirection, Width)
                    Dim ptIntersectR As Vector2D = MathFunction.FuncFindLineIntersection(ptIntersectLeft, ptIntersectRight, startPointCabWall, pointCabWal)
                    'вычисляем высотк шкафной стенки по низу
                    Dim bottomElevCabinetWall As Double = MathFunction.FuncCalcElevationByLine(leftLineCabinetWall.StartPoint, leftLineCabinetWall.EndPoint, ptIntersectR)
                    'правая конечная крайняя точка
                    x = Math.Round(ptIntersectR.X, 3)
                    y = Math.Round(ptIntersectR.Y, 3)
                    z = Math.Round(topElevCabinetWall, 3)
                    h1 = Math.Round(topElevCabinetWall - bottomElevCabinetWall, 3)
                    code = "leftPt0" & i
                    pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                    pointModelCabinetWall.Item(i + 1) = pointBidge

                    'левая конечная крайняя точка
                    x = Math.Round(startPointCabWall.X, 3)
                    y = Math.Round(startPointCabWall.Y, 3)
                    z = Math.Round(topElevCabinetWall, 3)
                    h1 = Math.Round(topElevCabinetWall - bottomElevCabinetWall, 3)
                    code = "rightPt0" & i
                    pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                    pointModelCabinetWall.Item(countElements - i) = pointBidge
                Next i
            End If
            _elementBridgePoint.ListPointModel = pointModelCabinetWall
            '========================================================================================================================
            'рисуем зуб упора
            '========================================================================================================================
            If WidthPl > 0 Then
                'полная ширина зуба упора
                Dim offsetPositionPl As Double = WidthPl
                'распаралеливаем крайнюю линию шкафной стенки
                Dim coolEntPl As List(Of DwgEntity) = New List(Of DwgEntity)
                Dim LinePl As DwgLine = New DwgLine
                If NumberPillar = 1 Then
                    rightLineCabinetWall.Offset(coolEntPl, offsetPositionPl) 'влево
                    LinePl = coolEntPl(0)
                Else
                    rightLineCabinetWall.Offset(coolEntPl, -1 * offsetPositionPl) 'вправо
                    LinePl = coolEntPl(0)
                End If
                'ищем пересечение линий с временными точками
                Dim ptIntersectLeftPl As Vector2D = MathFunction.FuncFindLineIntersection(LinePl.EndPoint.Pos, LinePl.StartPoint.Pos, rightLineCabinetWall.StartPoint.Pos, tempIntersectPointLeft, boolRez)
                If boolRez = False Then Return False
                Dim ptIntersectRightPl As Vector2D = MathFunction.FuncFindLineIntersection(LinePl.StartPoint.Pos, LinePl.EndPoint.Pos, rightLineCabinetWall.EndPoint.Pos, tempIntersectPointRight, boolRez)
                If boolRez = False Then Return False
                'левые точки
                Dim leftPoint1 As Vector2D = rightLineCabinetWall.StartPoint.Pos 'точка у шкафной стенки крайняя
                Dim leftPoint2 As Vector2D = ptIntersectLeftPl 'противоположная точка
                Dim middleleftPoint2 As Vector2D = MathFunction.FuncCalcPoint2DInLine(leftPoint2, leftPoint1, _widthUPl)
                'правые точки
                Dim rightPoint1 As Vector2D = rightLineCabinetWall.EndPoint.Pos
                Dim rightPoint2 As Vector2D = ptIntersectRightPl
                Dim middleRightPoint2 As Vector2D = MathFunction.FuncCalcPoint2DInLine(rightPoint2, rightPoint1, _widthUPl)
                'отметка верха площадки у шкафной стенки слева
                Dim elevationTopLeftPl As Double = Math.Round(leftLineCabinetWall.StartPoint.Z + HeightTopPl, 3)
                'отметка верха площадки у шкафной стенки справа
                Dim elevationTopRightPl As Double = Math.Round(leftLineCabinetWall.EndPoint.Z + HeightTopPl, 3)
                'отметка низа зуба упора точки у шкафной стенки слева
                Dim elevationBottomLeftPl As Double = Math.Round(elevationTopLeftPl - FullLengthPl, 3)
                'отметка низа зуба упора точки у шкафной стенки справа
                Dim elevationBottomRightPl As Double = Math.Round(elevationTopRightPl - FullLengthPl, 3)
                'отметка низа зуба упора самой крайней точки противоположной шкафной стенки слева
                Dim elevationBottomLeftPl2 As Double = Math.Round(elevationBottomLeftPl + _heightPl, 3)
                'отметка низа зуба упора самой крайней точки противоположной шкафной стенки справа
                Dim elevationBottomRightPl2 As Double = Math.Round(elevationBottomRightPl + _heightPl, 3)
                'делаем расчет высоты средней точки по низу зуба упора слева
                Dim tempPtZU1 As Vector3D = New Vector3D(leftPoint2, elevationBottomLeftPl2)
                Dim tempPtZU2 As Vector3D = New Vector3D(leftPoint1, elevationBottomLeftPl)
                Dim elevationBottomMiddlePoint As Double = MathFunction.FuncCalcElevationByLine(tempPtZU1, tempPtZU2, middleleftPoint2)
                'делаем расчет высоты средней точки по низу зуба упора справа
                tempPtZU1 = New Vector3D(rightPoint2, elevationBottomRightPl2)
                tempPtZU2 = New Vector3D(rightPoint1, elevationBottomRightPl)
                Dim elevationBottomMiddleRightPoint As Double = MathFunction.FuncCalcElevationByLine(tempPtZU1, tempPtZU2, middleRightPoint2)
                'отметка верха зуба упора самой крайней точки противоположной шкафной стенки слева
                Dim elevationTopLeftPl2 As Double = Math.Round(elevationBottomLeftPl2 + _lengthPl, 3)
                'отметка верха зуба упора самой крайней точки противоположной шкафной стенки справа
                Dim elevationTopRightPl2 As Double = Math.Round(elevationBottomRightPl2 + _lengthPl, 3)
                'отметка верха зуба упора самой крайней точки противоположной шкафной стенки слева
                Dim elevationTopMiddlePoint As Double = Math.Round(elevationTopLeftPl2 + _heightUPl, 3)
                'отметка верха зуба упора самой крайней точки противоположной шкафной стенки справа
                Dim elevationTopMiddleRightPoint As Double = Math.Round(elevationTopRightPl2 + _heightUPl, 3)

                pointModelCabinetWall = New Dictionary(Of Integer, PointStructure)
                'левая начальная крайняя точка
                x = Math.Round(leftPoint2.X, 3)
                y = Math.Round(leftPoint2.Y, 3)
                z = Math.Round(elevationTopLeftPl2, 3)
                h1 = Math.Round(LengthPl, 3)
                code = "leftPt1"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(1, pointBidge)

                'правая конечная крайняя точка
                x = Math.Round(rightPoint2.X, 3)
                y = Math.Round(rightPoint2.Y, 3)
                z = Math.Round(elevationTopRightPl2, 3)
                h1 = Math.Round(LengthPl, 3)
                code = "leftPt2"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(2, pointBidge)

                'правая промежуточная точка
                x = Math.Round(middleRightPoint2.X, 3)
                y = Math.Round(middleRightPoint2.Y, 3)
                z = Math.Round(elevationTopMiddleRightPoint, 3)
                h1 = Math.Round(elevationTopMiddleRightPoint - elevationBottomMiddleRightPoint, 3)
                code = "middlePt2"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(3, pointBidge)

                'левая конечная крайняя точка
                x = Math.Round(rightPoint1.X, 3)
                y = Math.Round(rightPoint1.Y, 3)
                z = Math.Round(elevationTopRightPl, 3)
                h1 = Math.Round(FullLengthPl, 3)
                code = "rightPt2"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(4, pointBidge)

                'левая начальная крайняя точка
                x = Math.Round(leftPoint1.X, 3)
                y = Math.Round(leftPoint1.Y, 3)
                z = Math.Round(elevationTopLeftPl, 3)
                h1 = Math.Round(FullLengthPl, 3)
                code = "rightPt1"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(5, pointBidge)

                'левая промежуточная точка
                x = Math.Round(middleleftPoint2.X, 3)
                y = Math.Round(middleleftPoint2.Y, 3)
                z = Math.Round(elevationTopMiddlePoint, 3)
                h1 = Math.Round(elevationTopMiddlePoint - elevationBottomMiddlePoint, 3)
                code = "middlePt1"
                pointBidge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
                pointModelCabinetWall.Add(6, pointBidge)

                _elementBridgePoint.ListPointSecondModel = pointModelCabinetWall
                'записываем ось шкафной стенки
                Dim middlePointLeft As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(leftLineCabinetWall.StartPoint, rightLineCabinetWall.StartPoint)
                Dim middlePointRight As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(leftLineCabinetWall.EndPoint, rightLineCabinetWall.EndPoint)
                _elementBridgePoint.StartAxisPoint = middlePointLeft
                _elementBridgePoint.EndAxisPoint = middlePointRight
                Dim centerAxisPoint As Vector2D = New Vector2D(-1, -1)
                If axisPline3D.Length2D > 0 Then
                    Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, middlePointLeft, middlePointRight)
                    If pointIntersectCollection.Count > 0 Then
                        centerAxisPoint = pointIntersectCollection.ElementAt(0)
                    End If
                End If
                If centerAxisPoint.X = -1 And centerAxisPoint.Y = -1 Then
                    centerAxisPoint = MathFunction.funcCalcMiddleCoordByToPoints2d(middlePointLeft, middlePointRight)
                End If
                _elementBridgePoint.CenterTopPoint = New Vector3D(centerAxisPoint, centerLeftElev)
            End If

        Else
            Return False
        End If
        Return True
    End Function
    'рисование оси шкафной стенки
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineCabinetWall As DwgLine = Nothing
        Dim dataStructureCabinetWall As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getAxis(dictionaryBridgeElements, NumberPillar, Number)
        If listAxis.Count = 0 Then
            dataStructureCabinetWall = createAxisCabinetWall(idBridge)
        ElseIf listAxis.Count = 1 Then
            dataStructureCabinetWall = listAxis.Item(0)
        Else
            dataStructureCabinetWall = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataStructureCabinetWall) Then Return dataStructureCabinetWall
        axisLineCabinetWall = dataStructureCabinetWall.DWGEntity
        If IsNothing(axisLineCabinetWall) = True Then axisLineCabinetWall = New DwgLine
        If axisLineCabinetWall.Length = 0 Then
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisCabinetWall As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisCabinetWall.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Шкафная стенка (ось)")
            styleAxisCabinetWall.setObjectStyle(axisLineCabinetWall)
        End If
        'ось насадки
        axisLineCabinetWall.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineCabinetWall.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLineCabinetWall.StartPoint.Pos - axisLineCabinetWall.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось шкафной стенки имеет нулевое значение.")
            Return dataStructureCabinetWall
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineCabinetWall) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineCabinetWall)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureCabinetWall.KeyParameter = strJson
        dataStructureCabinetWall.DWGEntity = axisLineCabinetWall
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineCabinetWall, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureCabinetWall)
        Return dataStructureCabinetWall
    End Function
    'функция получает все точки Шкафной стенки
    Public Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim ListPointModelWall As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
        If IsNothing(ListPointModelWall) = True Then
            Return result
        End If
        If ListPointModelWall.Count = 0 Then
            Return result
        End If
        'записываем результат
        For i As Integer = 0 To ListPointModelWall.Count - 1
            Dim pointStructure As PointStructure = ListPointModelWall.ElementAt(i).Value
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
    'функция получает все точки Шкафной стенки
    Public Function getProjectionPlatePoint(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim ListPointModelPlate As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointSecondModel
        If IsNothing(ListPointModelPlate) = True Then
            Return result
        End If
        If ListPointModelPlate.Count = 0 Then
            Return result
        End If
        'записываем результат
        For i As Integer = 0 To ListPointModelPlate.Count - 1
            Dim pointStructure As PointStructure = ListPointModelPlate.ElementAt(i).Value
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
                Dim ListPointModelWall As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
                If ListPointModelWall.Count > 0 Then
                    For i As Integer = 0 To ListPointModelWall.Count - 1
                        Dim ptStructure As PointStructure = ListPointModelWall.ElementAt(i).Value
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
    'функция возвращает координаты точки по ее коду
    Public Function getPointByCodeSecond(ByVal code As String) As Vector3D
        Dim result As Vector3D = New Vector3D
        If IsNothing(code) = False Then
            If code.Trim.Length > 0 Then
                Dim ListPointModelWall As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointSecondModel
                If ListPointModelWall.Count > 0 Then
                    For i As Integer = 0 To ListPointModelWall.Count - 1
                        Dim ptStructure As PointStructure = ListPointModelWall.ElementAt(i).Value
                        If ptStructure.Code Like code Then
                            result = New Vector3D(ptStructure.X - ptStructure.dx, ptStructure.Y - ptStructure.dy, ptStructure.Z)
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
        Return result
    End Function

End Class
