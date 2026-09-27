Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class GrillagePillar
    ' Приватные поля класса
    Private _numberPillar As Integer                              ' Номер опоры
    Private _number As Integer                                    ' Номер ростверка
    Private _width As Double                                      ' Ширина ростверка
    Private _height As Double                                     ' Высота ростверка
    Private _lenght As Double                                     ' Длина ростверка
    Private _offsetCenter As Double                               ' Смещение ростверка относительно оси опоры
    Private _depthFoundation As Double                            ' Глубина заложения относительно земли
    Private _offsetLeftRack As Double                             ' Отступ ростверка слева от крайней опоры
    Private _offsetRightRack As Double                            ' Отступ ростверка справа от крайней опоры
    Private _offsetEdgeRack As Double                             ' Отступ края ростверка от стойки (вперед/назад)
    Private _heightDrain As Double                                ' Высота слива
    Private _edgeParallel As Boolean                              ' торцы ростверка параллельны направлениям крайних балок
    Private _pileRowsFieldDiagram As String                       ' Схема расстановки свай для ряда
    Private _pileColumnFieldDiagram As String                     ' Схема расстановки свай для столбца
    Private _topElevation As Double                               ' Отметка верха ростверка
    Private _bottomElevation As Double                            ' Отметка низа ростверка
    Private _model As String
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _number = 0
        _width = 0.0
        _height = 0.0
        _lenght = 0
        _offsetCenter = 0.0
        _depthFoundation = 0.0
        _offsetLeftRack = 0.0
        _offsetRightRack = 0.0
        _offsetEdgeRack = 0.0
        _heightDrain = 0
        _edgeParallel = False
        _pileRowsFieldDiagram = String.Empty
        _pileColumnFieldDiagram = String.Empty
        _topElevation = 0.0
        _bottomElevation = 0.0
        _model = ""
        _elementBridgePoint = New PointsCollections
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
    <Description("Ширина ростверка")>
    <Category("Свойства")>
    <DisplayName("Ширина")>
    <[ReadOnly](True)>
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
    <Description("Высота ростверка")>
    <Category("Свойства")>
    <DisplayName("Высота")>
    Public Property Height() As Double
        Get
            Return _height
        End Get
        Set(value As Double)
            If value >= 0 Then
                _height = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина ростверка")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property Lenght() As Double
        Get
            Return _lenght
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lenght = value
            Else
                Throw New ArgumentException("Длина ростверка не может быть отрицательной")
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение ростверка от оси опоры")>
    <Category("Свойства")>
    <DisplayName("Смещение от оси")>
    Public Property OffsetCenter() As Double
        Get
            Return _offsetCenter
        End Get
        Set(value As Double)
            _offsetCenter = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Глубина заложения ростверка относительно земли")>
    <Category("Свойства")>
    <DisplayName("Глубина заложения")>
    Public Property DepthFoundation() As Double
        Get
            Return _depthFoundation
        End Get
        Set(value As Double)
            _depthFoundation = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Свес влево от крайней стойки, м")>
    <Category("Свойства")>
    <DisplayName("Свес влево")>
    Public Property OffsetLeftRack() As Double
        Get
            Return _offsetLeftRack
        End Get
        Set(value As Double)
            If value >= 0 Then
                _offsetLeftRack = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Свес вправо от крайней стойки, м")>
    <Category("Свойства")>
    <DisplayName("Свес вправо")>
    Public Property OffsetRightRack() As Double
        Get
            Return _offsetRightRack
        End Get
        Set(value As Double)
            If value >= 0 Then
                _offsetRightRack = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Свес вперед\назад от крайней стойки, м")>
    <Category("Свойства")>
    <DisplayName("Свес вперед\назад")>
    Public Property OffsetEdgeRack() As Double
        Get
            Return _offsetEdgeRack
        End Get
        Set(value As Double)
            _offsetEdgeRack = value
        End Set
    End Property

    ' Свойство для доступа к ширине ростверка
    Public Property EdgeParallel() As Boolean
        Get
            Return _edgeParallel
        End Get
        Set(value As Boolean)
            _edgeParallel = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота слива у ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Высота слива")>
    Public Property HeightDrain() As Double
        Get
            Return _heightDrain
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightDrain = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Схема расстановки для рядов свай (например 500+10*500+11*600+....")>
    <Category("Свойства")>
    <DisplayName("Схема дли ряда свай")>
    Public Property PileRowsFieldDiagram() As String
        Get
            Return _pileRowsFieldDiagram
        End Get
        Set(value As String)
            _pileRowsFieldDiagram = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Схема расстановки для столбцов свай (например 500+10*500+11*600+....")>
    <Category("Свойства")>
    <DisplayName("Схема дли столбцов свай")>
    Public Property PileColumnFieldDiagram() As String
        Get
            Return _pileColumnFieldDiagram
        End Get
        Set(value As String)
            _pileColumnFieldDiagram = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    Public Property TopElevation() As Double
        Get
            Return _topElevation
        End Get
        Set(value As Double)
            _topElevation = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка низа ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    Public Property BottomElevation() As Double
        Get
            Return _bottomElevation
        End Get
        Set(value As Double)
            _bottomElevation = value
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
    'создать новый ростверк
    Public Shared Function createAxisGrillagePillar(ByVal idBridge As String) As StructureElement
        Dim elementGrillage As StructureElement = New StructureElement()
        elementGrillage.Label = "Мосты и путепроводы"
        elementGrillage.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementGrillage.ClassObject = StructureElement.classStructure.GrillagePillar
        elementGrillage.Name = StructureElement.typeObject.axisGrillage
        elementGrillage.Description = "Ростверк (ось)"
        elementGrillage.KeyParameter = ""
        elementGrillage.IdElement = Guid.NewGuid.ToString
        elementGrillage.IdStructure = idBridge
        elementGrillage.Note = ""
        elementGrillage.DWGEntity = New DwgLine()
        Return elementGrillage
    End Function
    'ищет ростверк
    Public Shared Function getGrillagePillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataGrillage As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisGrillage) = True Then
            Dim listAxisGrillage = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisGrillage)
            If IsNothing(listAxisGrillage) = False Then
                If listAxisGrillage.Count > 0 Then
                    For k As Integer = 0 To listAxisGrillage.Count - 1
                        Dim tempData As StructureElement = listAxisGrillage.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisGrillage As GrillagePillar = tempData.getGrillagePillar
                            If IsNothing(userAxisGrillage) = False Then
                                If numberPillar = userAxisGrillage.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisGrillage.Number
                                    If numberSubPillar = userAxisGrillage.Number Then
                                        dataGrillage = tempData
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataGrillage
    End Function
    'чтение данных о ростверке
    Public Shared Function readPropertiesGrillage(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal columnsDiagram As String, ByVal rowsDiagram As String, ByVal DGV_Grillage As DataGridView, Optional edgeParallel As Boolean = False) As GrillagePillar
        Dim result As GrillagePillar = New GrillagePillar
        result.NumberPillar = numbPillar
        result.Number = numbSubPillar
        If DGV_Grillage.RowCount > 1 Then
            For i As Integer = 0 To DGV_Grillage.RowCount - 1
                Dim tag As String = DGV_Grillage.Rows(i).Tag
                Dim value As String = DGV_Grillage.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_grillage_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Height = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты ростверка.")
                        End If
                    ElseIf tag Like "bridge_grillage_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Val(value)
                        Else
                            MsgBox("Некорректное значение ширины ростверка.")
                        End If
                    ElseIf tag Like "bridge_grillage_laying" Then
                        If IsNumeric(value) = True Then
                            result.DepthFoundation = Val(value)
                        Else
                            MsgBox("Некорректное значение глубины заложения ростверка.")
                        End If
                    ElseIf tag Like "offsetLeftRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetLeftRack = Val(value)
                        Else
                            MsgBox("Некорректное значение выпуска ростверка за левую стойку.")
                        End If
                    ElseIf tag Like "offsetRightRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetRightRack = Val(value)
                        Else
                            MsgBox("Некорректное значение выпуска ростверка за правую стойку.")
                        End If
                    ElseIf tag Like "offsetCenter" Then
                        If IsNumeric(value) = True Then
                            result.OffsetCenter = Val(value)
                        Else
                            MsgBox("Некорректное значение смещения ростверка относительно оси опоры.")
                        End If
                    ElseIf tag Like "offsetEgeRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetEdgeRack = Val(value)
                        Else
                            MsgBox("Некорректное значение выпуска ростверка за стойку.")
                        End If
                    End If
                End If
            Next i
        End If
        result.PileRowsFieldDiagram = rowsDiagram
        result.PileColumnFieldDiagram = columnsDiagram
        result.EdgeParallel = edgeParallel
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesGrillage(ByRef DGV_Grillage As DataGridView) As Boolean
        If IsNothing(DGV_Grillage) = True Then Return False
        If DGV_Grillage.RowCount > 1 Then
            For j As Integer = 0 To DGV_Grillage.RowCount - 1
                Dim tag As String = DGV_Grillage.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_grillage_height" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = Height
                    ElseIf tag Like "bridge_grillage_width" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "bridge_grillage_laying" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = DepthFoundation
                    ElseIf tag Like "offsetLeftRack" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = OffsetLeftRack
                    ElseIf tag Like "offsetRightRack" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = OffsetRightRack
                    ElseIf tag Like "offsetCenter" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = OffsetCenter
                    ElseIf tag Like "offsetEgeRack" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = OffsetEdgeRack
                    ElseIf tag Like "heightDrain" Then
                        DGV_Grillage.Rows(j).Cells(1).Value = HeightDrain
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Grillage As DataGridView) As Boolean
        If IsNothing(DGV_Grillage) = True Then Return False
        If DGV_Grillage.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolBottomElevation As Boolean = False
                Dim boolLenght As Boolean = False
                Dim boolWidth As Boolean = False
                For i As Integer = 0 To DGV_Grillage.RowCount - 1
                    Dim oldTag As String = DGV_Grillage.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        DGV_Grillage.Rows(i).Cells(1).Value = TopElevation
                        boolTopElevation = True
                    ElseIf oldTag Like "calc-BottomElevation" Then
                        DGV_Grillage.Rows(i).Cells(1).Value = BottomElevation
                        boolBottomElevation = True
                    ElseIf oldTag Like "calc-Lenght" Then
                        DGV_Grillage.Rows(i).Cells(1).Value = Lenght
                        boolLenght = True
                    ElseIf oldTag Like "calc-Width" Then
                        DGV_Grillage.Rows(i).Cells(1).Value = Width
                        boolWidth = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_Grillage.RowCount - 1
                    DGV_Grillage.Rows.Insert(numberRow)
                    DGV_Grillage.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Grillage.Rows(numberRow).Cells(0).Value = "Отметка верха, м"
                    DGV_Grillage.Rows(numberRow).Cells(1).Value = TopElevation
                    DGV_Grillage.Rows(numberRow).Tag = "calc-TopElevation"
                End If
                If boolBottomElevation = False Then
                    Dim numberRow As Integer = DGV_Grillage.RowCount - 1
                    DGV_Grillage.Rows.Insert(numberRow)
                    DGV_Grillage.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Grillage.Rows(numberRow).Cells(0).Value = "Отметка низа, м"
                    DGV_Grillage.Rows(numberRow).Cells(1).Value = BottomElevation
                    DGV_Grillage.Rows(numberRow).Tag = "calc-BottomElevation"
                End If
                If boolLenght = False Then
                    Dim numberRow As Integer = DGV_Grillage.RowCount - 1
                    DGV_Grillage.Rows.Insert(numberRow)
                    DGV_Grillage.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Grillage.Rows(numberRow).Cells(0).Value = "Полная длина, м"
                    DGV_Grillage.Rows(numberRow).Cells(1).Value = Lenght
                    DGV_Grillage.Rows(numberRow).Tag = "calc-Lenght"
                End If
                If boolWidth = False Then
                    Dim numberRow As Integer = DGV_Grillage.RowCount - 1
                    DGV_Grillage.Rows.Insert(numberRow)
                    DGV_Grillage.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Grillage.Rows(numberRow).Cells(0).Value = "Ширина, м"
                    DGV_Grillage.Rows(numberRow).Cells(1).Value = Width
                    DGV_Grillage.Rows(numberRow).Tag = "calc-Width"
                End If
            End If
        End If
        Return True
    End Function
    'расчет ростверка
    Public Function calculateGrillage(ByVal axisPillar As DwgLine, ByVal arrayRack As RackPillar(), Optional ByVal axisPlineAlign As Polyline3D = Nothing) As Boolean
        If IsNothing(axisPillar) = True Then Return False
        If axisPillar.Length = 0 Then Return False

        'Для расчета нужна удлиненная вспомогательная линия. DwgLine является
        'ссылочным типом, поэтому простое присваивание изменяло исходную ось
        'опоры, находящуюся в чертеже.
        Dim axisLinePillar As New DwgLine()
        axisLinePillar.StartPoint = axisPillar.StartPoint
        axisLinePillar.EndPoint = axisPillar.EndPoint
        Dim boolFundUserRack As Boolean = False
        If IsNothing(arrayRack) = True Then
            boolFundUserRack = True
        ElseIf arrayRack.Length = 0 Then
            boolFundUserRack = True
        End If
        Dim boolExt As Boolean = BridgeGeometry.extendLine(axisLinePillar, 10, 10)
        Dim dirRPillar As Double = axisLinePillar.Rotation + Math.PI / 2
        If dirRPillar > Math.PI * 2 Then
            dirRPillar -= Math.PI * 2
        End If
        Dim dirRNPillar As Double = dirRPillar + Math.PI
        If dirRNPillar > Math.PI * 2 Then
            dirRNPillar -= Math.PI * 2
        End If
        Dim angleDirectPillar As Double = axisLinePillar.Rotation
        Dim reverseAngleDirectPillar As Double = angleDirectPillar + Math.PI
        If reverseAngleDirectPillar > Math.PI * 2 Then
            reverseAngleDirectPillar -= Math.PI * 2
        End If

        Dim lenghtLeft As Double = 0
        Dim lenghtRight As Double = 0
        Dim lastLeftPoint As Vector2D = New Vector2D(-1, -1)
        Dim lastRightPoint As Vector2D = New Vector2D(-1, -1)
        Dim lastDistLeft As Double = 9999
        Dim lastDistRight As Double = 9999
        Dim elevationGrillage As Double = 0

        For i As Integer = 0 To arrayRack.Length - 1
            Dim userRack As RackPillar = arrayRack(i)
            If IsNothing(userRack) = True Then Continue For
            If userRack._elementBridgePoint.ListPointModel.Count > 3 Then
                For k As Integer = 0 To userRack._elementBridgePoint.ListPointModel.Count - 1
                    Dim ptRack As PointStructure = userRack._elementBridgePoint.ListPointModel.ElementAt(k).Value
                    Dim ptRackGrillage As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(ptRack.X - ptRack.dx, ptRack.Y - ptRack.dy, ptRack.Z + ptRack.dz)
                    Dim boolRez As Boolean = False
                    Dim pointLeftPPillar As Vector2D = MathFunction.funcGetNearestPointOnLine(axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, ptRackGrillage, boolRez)
                    If boolRez = True Then
                        Dim tempLeftDist As Double = (pointLeftPPillar - ptRackGrillage.Pos).Length
                        Dim indLeft As Integer = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptRackGrillage.Pos)
                        If indLeft < 0 Then
                            If tempLeftDist > lenghtLeft Then
                                lenghtLeft = tempLeftDist
                            End If
                        Else
                            If tempLeftDist > lenghtRight Then
                                lenghtRight = tempLeftDist
                            End If
                        End If
                        'высота низа стойкм
                        If ptRackGrillage.Z > elevationGrillage Then
                            elevationGrillage = ptRackGrillage.Z
                            TopElevation = Math.Round(elevationGrillage, 3)
                            BottomElevation = Math.Round(elevationGrillage - Height, 3)
                        End If
                        If i = 0 Then
                            'крайняя левая стойка
                            Dim tempDist As Double = (axisLinePillar.StartPoint.Pos - pointLeftPPillar).Length
                            If tempDist < lastDistLeft Then
                                lastDistLeft = tempDist
                                lastLeftPoint = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeftPPillar, reverseAngleDirectPillar, OffsetLeftRack)
                            End If
                        End If
                        If i = arrayRack.Length - 1 Then
                            'крайняя левая стойка
                            Dim tempDist As Double = (axisLinePillar.EndPoint.Pos - pointLeftPPillar).Length
                            If tempDist < lastDistRight Then
                                lastDistRight = tempDist
                                lastRightPoint = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeftPPillar, angleDirectPillar, OffsetRightRack)
                            End If
                        End If
                    End If
                Next k
            End If
        Next i
        'отметка верха 
        'находим позицию ростверка путем параллельного переноса оси опоры на веричину смещения
        Dim leftLineGrillage As DwgLine = New DwgLine()
        Dim rightLineGrillage As DwgLine = New DwgLine()
        If NumberPillar = 1 Then
            Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            axisLinePillar.Offset(listEnt, -1 * lenghtLeft - OffsetEdgeRack)
            If listEnt.Count > 0 Then
                leftLineGrillage = listEnt(0)
            End If
            listEnt = New List(Of DwgEntity)
            axisLinePillar.Offset(listEnt, lenghtRight + OffsetEdgeRack)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                rightLineGrillage = listEnt(0)
            End If
        Else
            Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            axisLinePillar.Offset(listEnt, -1 * lenghtLeft - OffsetEdgeRack)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                leftLineGrillage = listEnt(0)
            End If
            listEnt = New List(Of DwgEntity)
            axisLinePillar.Offset(listEnt, lenghtRight + OffsetEdgeRack)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                rightLineGrillage = listEnt(0)
            End If
        End If

        Dim pointGrillLeft As Vector2D = New Vector2D(-1, -1)
        Dim pointGrillRight As Vector2D = New Vector2D(-1, -1)
        pointGrillLeft = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lastLeftPoint, dirRPillar, 10)
        pointGrillRight = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lastRightPoint, dirRPillar, 10)
        'ищем пересечение линий с временными точками
        Dim positionGrillagePointLeft1 As Vector2D = MathFunction.FuncFindLineIntersection(leftLineGrillage.StartPoint.Pos, leftLineGrillage.EndPoint.Pos, pointGrillLeft, lastLeftPoint)
        Dim positionGrillagePointRight1 As Vector2D = MathFunction.FuncFindLineIntersection(leftLineGrillage.StartPoint.Pos, leftLineGrillage.EndPoint.Pos, pointGrillRight, lastRightPoint)
        Dim positionGrillagePointLeft2 As Vector2D = MathFunction.FuncFindLineIntersection(rightLineGrillage.StartPoint.Pos, rightLineGrillage.EndPoint.Pos, pointGrillLeft, lastLeftPoint)
        Dim positionGrillagePointRight2 As Vector2D = MathFunction.FuncFindLineIntersection(rightLineGrillage.StartPoint.Pos, rightLineGrillage.EndPoint.Pos, pointGrillRight, lastRightPoint)
        Dim centerLeft As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionGrillagePointLeft1, positionGrillagePointLeft2)
        Dim centerRight As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionGrillagePointRight1, positionGrillagePointRight2)
        Dim middlePoint1 As Vector3D = New Cad.Foundation.Vector3D(centerLeft, elevationGrillage)
        Dim middlePoint2 As Vector3D = New Cad.Foundation.Vector3D(centerRight, elevationGrillage)

        Dim lenghtGrillage As Double = (positionGrillagePointLeft1 - positionGrillagePointRight1).Length
        Dim widthGrillage As Double = (positionGrillagePointLeft1 - positionGrillagePointLeft2).Length
        Lenght = Math.Round(lenghtGrillage, 3)
        Width = Math.Round(widthGrillage, 3)
        Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        Dim countPoint As Integer = 1
        Dim x As Double = Math.Round(positionGrillagePointLeft2.X, 3)
        Dim y As Double = Math.Round(positionGrillagePointLeft2.Y, 3)
        Dim z As Double = Math.Round(elevationGrillage, 3)
        Dim dx As Double = 0
        Dim dy As Double = 0
        Dim dz As Double = Math.Round(Height, 3)
        Dim code As String = "leftPt1"
        Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
        listModelPoint.Add(countPoint, pointModel)
        countPoint += 1

        x = Math.Round(positionGrillagePointRight2.X, 3)
        y = Math.Round(positionGrillagePointRight2.Y, 3)
        z = Math.Round(elevationGrillage, 3)
        dx = 0
        dy = 0
        dz = Math.Round(Height, 3)
        code = "leftPt2"
        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
        listModelPoint.Add(countPoint, pointModel)
        countPoint += 1

        If HeightDrain > 0 Then
            x = Math.Round(centerRight.X, 3)
            y = Math.Round(centerRight.Y, 3)
            z = Math.Round(elevationGrillage + HeightDrain, 3)
            dx = 0
            dy = 0
            dz = Math.Round(Height + HeightDrain, 3)
            code = "middlePt2"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(countPoint, pointModel)
            middlePoint1 = New Cad.Foundation.Vector3D(centerRight, elevationGrillage + HeightDrain)
            countPoint += 1
        End If

        x = Math.Round(positionGrillagePointRight1.X, 3)
        y = Math.Round(positionGrillagePointRight1.Y, 3)
        z = Math.Round(elevationGrillage, 3)
        dx = 0
        dy = 0
        dz = Math.Round(Height, 3)
        code = "rightPt2"
        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
        listModelPoint.Add(countPoint, pointModel)
        countPoint += 1

        x = Math.Round(positionGrillagePointLeft1.X, 3)
        y = Math.Round(positionGrillagePointLeft1.Y, 3)
        z = Math.Round(elevationGrillage, 3)
        dx = 0
        dy = 0
        dz = Math.Round(Height, 3)
        code = "rightPt1"
        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
        listModelPoint.Add(countPoint, pointModel)
        countPoint += 1

        If HeightDrain > 0 Then
            x = Math.Round(centerLeft.X, 3)
            y = Math.Round(centerLeft.Y, 3)
            z = Math.Round(elevationGrillage + HeightDrain, 3)
            dx = 0
            dy = 0
            dz = Math.Round(Height + HeightDrain, 3)
            code = "middlePt1"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(countPoint, pointModel)
            middlePoint2 = New Cad.Foundation.Vector3D(centerLeft, elevationGrillage + HeightDrain)
        End If
        _elementBridgePoint.ListPointModel = listModelPoint

        _elementBridgePoint.StartAxisPoint = middlePoint1
        _elementBridgePoint.EndAxisPoint = middlePoint2
        Dim centerAxisPoint As Vector2D = New Vector2D(-1, -1)
        If IsNothing(axisPlineAlign) = False Then
            If axisPlineAlign.Length2D > 0 Then
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlineAlign, middlePoint1.Pos, middlePoint2)
                If pointIntersectCollection.Count > 0 Then
                    centerAxisPoint = pointIntersectCollection.ElementAt(0)
                End If
            End If
        End If
        If centerAxisPoint.X = -1 And centerAxisPoint.Y = -1 Then
            centerAxisPoint = MathFunction.funcCalcMiddleCoordByToPoints2d(middlePoint1, middlePoint2)
        End If
        _elementBridgePoint.CenterTopPoint = New Vector3D(centerAxisPoint, elevationGrillage + HeightDrain)
        Return True
    End Function
    'рисование оси ростверка
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineGrillage As DwgLine = Nothing
        Dim dataStructureGrillage As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxisGrillage As List(Of StructureElement) = Pillar.getElementPillar(StructureElement.typeObject.axisGrillage, dictionaryBridgeElements, NumberPillar, 0, 0, Number, Pillar.SidePillarElement.None)
        If listAxisGrillage.Count = 0 Then
            dataStructureGrillage = createAxisGrillagePillar(idBridge)
        ElseIf listAxisGrillage.Count = 1 Then
            dataStructureGrillage = listAxisGrillage.Item(0)
        Else
            dataStructureGrillage = StructureElement.isValidateDataStructure(listAxisGrillage)
        End If
        If IsNothing(dataStructureGrillage) Then Return Nothing
        axisLineGrillage = dataStructureGrillage.DWGEntity
        If IsNothing(axisLineGrillage) = True Then Return Nothing
        If axisLineGrillage.Length = 0 Then
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ростверк (ось)")
            styleAxisNozzle.setObjectStyle(axisLineGrillage)
        End If
        'ось насадки
        axisLineGrillage.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineGrillage.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxis As Double = (axisLineGrillage.StartPoint.Pos - axisLineGrillage.EndPoint.Pos).Length
        If lenghtAxis = 0 Then
            MsgBox("Ось ростверка имеет нулевое значение.")
            Return dataStructureGrillage
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineGrillage) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineGrillage)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureGrillage.KeyParameter = strJson
        dataStructureGrillage.DWGEntity = axisLineGrillage
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineGrillage, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureGrillage)
        Return dataStructureGrillage
    End Function
    'функция возвращает все точки ростверка в виде словаря 
    Public Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
        If IsNothing(ListPointModel) = True Then
            Return result
        End If
        If ListPointModel.Count = 0 Then
            Return result
        End If
        'записываем результат
        For i As Integer = 0 To ListPointModel.Count - 1
            Dim pointStructure As PointStructure = ListPointModel.ElementAt(i).Value
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
                Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
                If ListPointModel.Count > 0 Then
                    For i As Integer = 0 To ListPointModel.Count - 1
                        Dim ptStructure As PointStructure = ListPointModel.ElementAt(i).Value
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
