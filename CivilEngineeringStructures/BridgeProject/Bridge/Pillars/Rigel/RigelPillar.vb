Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class RigelPillar
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _number As Integer                        ' Номер ригеля
    Private _width As Double                          ' Ширина ригеля
    Private _height As Double                         ' Высота ригеля
    Private _lenght As Double                         ' Длина ригеля
    Private _outletLeftBeam As Double                 ' Минимальный выпуск ригеля влево за габарит сооружения
    Private _outletRightBeam As Double                ' Выпуск ригеля вправо за габарит сооружения
    Private _heightLeftConsole As Double              ' Высота торца ригеля слева
    Private _heightRightConsole As Double             ' Высота торца ригеля справа
    Private _lengthLeftConsole As Double              ' Длина левой консоли
    Private _lengthRightConsole As Double             ' Длина правой консоли
    Private _minElevationBeams As Double              ' Расстояние от балки до ригеля
    Private _leftDirection As Double                  ' Дирекционное направление левой грани ригеля
    Private _rightDirection As Double                 ' Дирекционное направление правой грани ригеля
    Private _topElevation As Double                   ' Отметка верха
    Private _bottomElevation As Double                ' Отметка низа
    Private _axisOffset As Double                     ' Смещение центра ригеля относительно оси
    Private _offsetLeftRack As Double                 ' Смещение начала левой стойки
    Private _offsetRightRack As Double                ' Смещение начала правой стойки
    Private _countRack As Integer                     ' Количество стоек
    Private _stepRack As Double                       ' Шаг расстановки стоек
    Private _heightDrain As Double                    ' Высота слива
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
        _outletLeftBeam = 0.0
        _outletRightBeam = 0.0
        _heightLeftConsole = 0.0
        _heightRightConsole = 0.0
        _lengthLeftConsole = 0.0
        _lengthRightConsole = 0.0
        _minElevationBeams = 0.0
        _leftDirection = 0.0
        _rightDirection = 0.0
        _topElevation = 0.0
        _bottomElevation = 0.0
        _axisOffset = 0.0
        _offsetLeftRack = 0.0
        _offsetRightRack = 0.0
        _countRack = 0
        _stepRack = 0.0
        _heightDrain = 0.0
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
    <Description("Ширина ригеля, м")>
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
    <Description("Высота ригеля, м")>
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
    <Description("Длина ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    Public Property Lenght() As Double
        Get
            Return _lenght
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lenght = value
            End If
        End Set
    End Property
    <Browsable(True)>
    <Description("Выпуск ригеля за левую стойку, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск влево")>
    Public Property OutletLeftBeam() As Double
        Get
            Return _outletLeftBeam
        End Get
        Set(value As Double)
            If value >= 0 Then
                _outletLeftBeam = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Выпуск ригеля за правую стойку, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск вправо")>
    Public Property OutletRightBeam() As Double
        Get
            Return _outletRightBeam
        End Get
        Set(value As Double)
            If value >= 0 Then
                _outletRightBeam = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота левой консоли ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Высота левой консоли")>
    Public Property HeightLeftConsole() As Double
        Get
            Return _heightLeftConsole
        End Get
        Set(value As Double)
            _heightLeftConsole = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота правой консоли ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Высота правой консоли")>
    Public Property HeightRightConsole() As Double
        Get
            Return _heightRightConsole
        End Get
        Set(value As Double)
            _heightRightConsole = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина левой консоли ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Длина левой консоли")>
    Public Property LengthLeftConsole() As Double
        Get
            Return _lengthLeftConsole
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthLeftConsole = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина правой консоли ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Длина правой консоли")>
    Public Property LengthRightConsole() As Double
        Get
            Return _lengthRightConsole
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthRightConsole = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Минимальное расстояние от точки опирания балки, до верха ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние от балки до ригеля")>
    Public Property MinElevationBeams() As Double
        Get
            Return _minElevationBeams
        End Get
        Set(value As Double)
            If value >= 0 Then
                _minElevationBeams = value
            End If
        End Set
    End Property

    ' Свойство для доступа к дирекционному направлению левой грани
    Public Property LeftDirection() As Double
        Get
            Return _leftDirection
        End Get
        Set(value As Double)
            _leftDirection = value
        End Set
    End Property

    ' Свойство для доступа к дирекционному направлению правой грани
    Public Property RightDirection() As Double
        Get
            Return _rightDirection
        End Get
        Set(value As Double)
            _rightDirection = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха ригеля, м")>
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
    <Description("Отметка низа ригеля, м")>
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
    <Description("Смещение ригеля относительно оси опоры , м")>
    <Category("Свойства")>
    <DisplayName("Смещение ригеля")>
    Public Property AxisOffset() As Double
        Get
            Return _axisOffset
        End Get
        Set(value As Double)
            _axisOffset = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Расстояние от края ригеля до края левой стойки, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние до левой стойки")>
    Public Property OffsetLeftRack() As Double
        Get
            Return _offsetLeftRack
        End Get
        Set(value As Double)
            _offsetLeftRack = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Расстояние от края ригеля до края правой стойки, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние до правой стойки")>
    Public Property OffsetRightRack() As Double
        Get
            Return _offsetRightRack
        End Get
        Set(value As Double)
            _offsetRightRack = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Количество стоек, шт")>
    <Category("Свойства")>
    <DisplayName("Количество стоек")>
    Public Property CountRack() As Integer
        Get
            Return _countRack
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _countRack = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Шаг расстановки стоек, м")>
    <Category("Свойства")>
    <DisplayName("Шаг стоек")>
    Public Property StepRack() As Double
        Get
            Return _stepRack
        End Get
        Set(value As Double)
            If value >= 0 Then
                _stepRack = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота слива, м")>
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
    ' Свойство для доступа к имени модели
    Public Property NameModel() As String
        Get
            Return _model
        End Get
        Set(value As String)
            _model = value
        End Set
    End Property

    'функции для работы с ригелем
    Public Shared Function createAxisRigelPillar(ByVal idBridge As String) As StructureElement
        Dim elementRigel As StructureElement = New StructureElement()
        elementRigel.Label = "Мосты и путепроводы"
        elementRigel.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementRigel.ClassObject = StructureElement.classStructure.RigelPillar
        elementRigel.Name = StructureElement.typeObject.axisRigel
        elementRigel.Description = "Ригель (ось)"
        elementRigel.KeyParameter = ""
        elementRigel.IdElement = Guid.NewGuid.ToString
        elementRigel.IdStructure = idBridge
        elementRigel.Note = ""
        elementRigel.DWGEntity = New DwgLine()
        Return elementRigel
    End Function
    'ищет ригель
    Public Shared Function getAxis(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim dataRigelPillar As List(Of StructureElement) = New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return dataRigelPillar
        If numberPillar < 1 Then Return dataRigelPillar
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisRigel) = True Then
            Dim listAxisRigel = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisRigel)
            If IsNothing(listAxisRigel) = False Then
                If listAxisRigel.Count > 0 Then
                    For k As Integer = 0 To listAxisRigel.Count - 1
                        Dim tempData As StructureElement = listAxisRigel.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisRigel As RigelPillar = tempData.getRigelPillar
                            If IsNothing(userAxisRigel) = False Then
                                If numberPillar = userAxisRigel.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisRigel.Number
                                    If numberSubPillar = userAxisRigel.Number Then
                                        dataRigelPillar.Add(tempData)
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataRigelPillar
    End Function
    'чтение таблиц диалогового окна и запись значений элементов сооружения
    Public Shared Function readPropertiesRigel(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Rigel As DataGridView) As RigelPillar
        Dim result As RigelPillar = New RigelPillar
        result.NumberPillar = numbPillar
        result.Number = numbSubPillar
        If DGV_Rigel.RowCount > 1 Then
            For i As Integer = 0 To DGV_Rigel.RowCount - 1
                Dim tag As String = DGV_Rigel.Rows(i).Tag
                Dim value As String = DGV_Rigel.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_rigel_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Height = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты ригеля.")
                        End If
                    ElseIf tag Like "bridge_rigel_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Val(value)
                        Else
                            MsgBox("Некорректное значение ширины ригеля.")
                        End If
                    ElseIf tag Like "axisOffset" Then
                        If IsNumeric(value) = True Then
                            result.AxisOffset = Val(value)
                        Else
                            MsgBox("Некорректное значение смещения оси ригеля.")
                        End If
                    ElseIf tag Like "minElevationBeams" Then
                        If IsNumeric(value) = True Then
                            result.MinElevationBeams = Val(value)
                        Else
                            MsgBox("Некорректное значение минимальной высоты подферменника..")
                        End If
                    ElseIf tag Like "outletLeftBeam" Then
                        If IsNumeric(value) = True Then
                            result.OutletLeftBeam = Val(value)
                        Else
                            MsgBox("Некорректное значение левого отступа ригеля за балку.")
                        End If
                    ElseIf tag Like "outletRightBeam" Then
                        If IsNumeric(value) = True Then
                            result.OutletRightBeam = Val(value)
                        Else
                            MsgBox("Некорректное значение правого отступа ригеля за балку.")
                        End If
                    ElseIf tag Like "bridge_rigel_face_height" Then
                        If IsNumeric(value) = True Then
                            result.HeightLeftConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты левой консоли.")
                        End If
                    ElseIf tag Like "heightRightConsole" Then
                        If IsNumeric(value) = True Then
                            result.HeightRightConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты правой консоли.")
                        End If
                    ElseIf tag Like "bridge_rigel_bevel_rack_left" Then
                        If IsNumeric(value) = True Then
                            result.LengthLeftConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение длины левой консоли.")
                        End If
                    ElseIf tag Like "bridge_rigel_bevel_rack_right" Then
                        If IsNumeric(value) = True Then
                            result.LengthRightConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение длины правой консоли.")
                        End If
                    ElseIf tag Like "offsetLeftRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetLeftRack = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа края стойки от левой стороны ригеля.")
                        End If
                    ElseIf tag Like "offsetRightRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetRightRack = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа края стойки от правой стороны ригеля.")
                        End If
                    ElseIf tag Like "heightDrain" Then
                        If IsNumeric(value) = True Then
                            result.HeightDrain = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты слива.")
                        End If
                    End If
                End If
            Next i
        End If
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesRigel(ByRef DGV_Rigel As DataGridView) As Boolean
        If IsNothing(DGV_Rigel) = True Then Return False
        If DGV_Rigel.RowCount > 1 Then
            For j As Integer = 0 To DGV_Rigel.RowCount - 1
                Dim tag As String = DGV_Rigel.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_rigel_height" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = Height
                    ElseIf tag Like "bridge_Rigel_width" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "minElevationBeams" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = MinElevationBeams
                    ElseIf tag Like "axisOffset" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = AxisOffset
                    ElseIf tag Like "outletLeftBeam" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = OutletLeftBeam
                    ElseIf tag Like "outletRightBeam" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = OutletRightBeam
                    ElseIf tag Like "offsetLeftRack" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = OffsetLeftRack
                    ElseIf tag Like "offsetRightRack" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = OffsetRightRack
                    ElseIf tag Like "bridge_rigel_face_height" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = HeightLeftConsole
                    ElseIf tag Like "heightRightConsole" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = HeightRightConsole
                    ElseIf tag Like "bridge_rigel_bevel_rack_left" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = LengthLeftConsole
                    ElseIf tag Like "lenghtRightConsole" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = LengthRightConsole
                    ElseIf tag Like "heightDrain" Then
                        DGV_Rigel.Rows(j).Cells(1).Value = HeightDrain
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Rigel As DataGridView) As Boolean
        If IsNothing(DGV_Rigel) = True Then Return False
        If DGV_Rigel.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolBottomElevation As Boolean = False
                Dim boolLenght As Boolean = False
                Dim boolStepRacks As Boolean = False
                For i As Integer = 0 To DGV_Rigel.RowCount - 1
                    Dim oldTag As String = DGV_Rigel.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        DGV_Rigel.Rows(i).Cells(1).Value = TopElevation
                        boolTopElevation = True
                    ElseIf oldTag Like "calc-BottomElevation" Then
                        DGV_Rigel.Rows(i).Cells(1).Value = BottomElevation
                        boolBottomElevation = True
                    ElseIf oldTag Like "calc-Lenght" Then
                        DGV_Rigel.Rows(i).Cells(1).Value = Lenght
                        boolLenght = True
                    ElseIf oldTag Like "StepRacks" Then
                        DGV_Rigel.Rows(i).Cells(1).Value = StepRack
                        boolStepRacks = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_Rigel.RowCount - 1
                    DGV_Rigel.Rows.Insert(numberRow)
                    DGV_Rigel.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rigel.Rows(numberRow).Cells(0).Value = "Отметка верха, м"
                    DGV_Rigel.Rows(numberRow).Cells(1).Value = TopElevation
                    DGV_Rigel.Rows(numberRow).Tag = "calc-TopElevation"
                End If
                If boolBottomElevation = False Then
                    Dim numberRow As Integer = DGV_Rigel.RowCount - 1
                    DGV_Rigel.Rows.Insert(numberRow)
                    DGV_Rigel.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rigel.Rows(numberRow).Cells(0).Value = "Отметка низа, м"
                    DGV_Rigel.Rows(numberRow).Cells(1).Value = BottomElevation
                    DGV_Rigel.Rows(numberRow).Tag = "calc-BottomElevation"
                End If
                If boolLenght = False Then
                    Dim numberRow As Integer = DGV_Rigel.RowCount - 1
                    DGV_Rigel.Rows.Insert(numberRow)
                    DGV_Rigel.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rigel.Rows(numberRow).Cells(0).Value = "Полная длина, м"
                    DGV_Rigel.Rows(numberRow).Cells(1).Value = Lenght
                    DGV_Rigel.Rows(numberRow).Tag = "calc-Lenght"
                End If
                If boolStepRacks = False Then
                    Dim numberRow As Integer = DGV_Rigel.RowCount - 1
                    DGV_Rigel.Rows.Insert(numberRow)
                    DGV_Rigel.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rigel.Rows(numberRow).Cells(0).Value = "Шаг расстановки стоек, м"
                    DGV_Rigel.Rows(numberRow).Cells(1).Value = StepRack
                    DGV_Rigel.Rows(numberRow).Tag = "StepRacks"
                End If
            End If
        End If
        Return True
    End Function

    'функции для расчета положения элементов опоры
    Public Function calculateRigel(ByVal axisLinePillar As DwgLine, ByVal listBeams As List(Of Dictionary(Of Integer, StructureElement)), Optional align As Polyline3D = Nothing, Optional arraySubFerment As SubFermenters() = Nothing) As Boolean
        If IsNothing(axisLinePillar) = True Then
            Dim startPointCenterRigel As Vector3D = getPointByCode("middlePt1", True)
            Dim endPointCenterRigel As Vector3D = getPointByCode("middlePt2", True)
            axisLinePillar = New DwgLine()
            axisLinePillar.StartPoint = startPointCenterRigel
            axisLinePillar.EndPoint = endPointCenterRigel
        End If
        If axisLinePillar.Length = 0 Then
            Return False
        End If
        'находим балку с минимальной высотой
        Dim minElevationBeam As Double = CalculationBeams.getBeamToMinElevation(listBeams, arraySubFerment)
        If minElevationBeam = 999999 Then
            MsgBox("Не удалось найти отметку самой нижней балки!!!")
            Return False
        End If
        'находим отметки насадки
        TopElevation = Math.Round(minElevationBeam - MinElevationBeams, 3)
        BottomElevation = Math.Round(TopElevation - Height, 3)
        'из массива найденных балок находим крайние (справа и с лева
        Dim listExtremeBeams As List(Of StructureElement) = CalculationBeams.getExtrmBeamsToPillar(listBeams)
        'левые балки смежных пролетов
        Dim leftDataBeam1 As StructureElement = Nothing
        Dim leftDataBeam2 As StructureElement = Nothing
        'правые балки смежных пролетов
        Dim rightDataBeam1 As StructureElement = Nothing
        Dim rightDataBeam2 As StructureElement = Nothing
        If IsNothing(listExtremeBeams.Item(0)) = False Then
            leftDataBeam1 = listExtremeBeams.Item(0)
        End If
        If IsNothing(listExtremeBeams.Item(2)) = False Then
            leftDataBeam2 = listExtremeBeams.Item(2)
        End If
        If IsNothing(listExtremeBeams.Item(1)) = False Then
            rightDataBeam1 = listExtremeBeams.Item(1)
        End If
        If IsNothing(listExtremeBeams.Item(3)) = False Then
            rightDataBeam2 = listExtremeBeams.Item(3)
        End If
        Dim leftRotationEgeRigel As Double = 0
        Dim leftAxisBeam1 As DwgLine = leftDataBeam1.DWGEntity
        Dim leftAxisBeam2 As DwgLine = leftDataBeam2.DWGEntity
        If leftAxisBeam1.Length > 0 And leftAxisBeam2.Length > 0 Then
            leftRotationEgeRigel = (leftAxisBeam1.Rotation + leftAxisBeam2.Rotation) / 2
        ElseIf leftAxisBeam1.Length > 0 Then
            leftRotationEgeRigel = leftAxisBeam1.Rotation
        ElseIf leftAxisBeam2.Length > 0 Then
            leftRotationEgeRigel = leftAxisBeam2.Rotation
        Else
            MsgBox("Не удалось рассчитать направление левой грани ригеля.")
            Return False
        End If
        Dim rightRotationEgeRigel As Double = 0
        Dim rightAxisBeam1 As DwgLine = rightDataBeam1.DWGEntity
        Dim rightAxisBeam2 As DwgLine = rightDataBeam2.DWGEntity
        If rightAxisBeam1.Length > 0 And rightAxisBeam2.Length > 0 Then
            rightRotationEgeRigel = (rightAxisBeam1.Rotation + rightAxisBeam2.Rotation) / 2
        ElseIf rightAxisBeam1.Length > 0 Then
            rightRotationEgeRigel = rightAxisBeam1.Rotation
        ElseIf rightAxisBeam2.Length > 0 Then
            rightRotationEgeRigel = rightAxisBeam2.Rotation
        Else
            MsgBox("Не удалось рассчитать направление правой грани ригеля.")
            Return False
        End If
        'находим направления короткой стороны насадки
        LeftDirection = Math.Round(leftRotationEgeRigel, 6)
        RightDirection = Math.Round(rightRotationEgeRigel, 6)
        '===================================================================================================================
        'строим габарит насадки
        Dim offsetLeftDist As Double = Width / 2 + AxisOffset
        Dim offsetRightDist As Double = Width / 2 - AxisOffset
        '4 крайние точки
        Dim positionRigelPointLeft1 As Vector2D = Nothing
        Dim positionRigelPointLeft2 As Vector2D = Nothing
        Dim positionRigelPointRight1 As Vector2D = Nothing
        Dim positionRigelPointRight2 As Vector2D = Nothing
        'находим позицию насадки путем параллельного переноса линиии начала шкафной стенки
        Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLinePillar.Offset(listEnt, offsetRightDist)
        If listEnt.Count > 0 Then
            Dim newLine As DwgLine = listEnt(0)
            positionRigelPointLeft1 = newLine.StartPoint
            positionRigelPointRight1 = newLine.EndPoint
        End If
        listEnt = New List(Of DwgEntity)
        axisLinePillar.Offset(listEnt, -1 * offsetLeftDist)
        If listEnt.Count > 0 Then
            Dim newLine As DwgLine = listEnt(0)
            positionRigelPointLeft2 = newLine.StartPoint
            positionRigelPointRight2 = newLine.EndPoint
        End If

        'находим пересечение крайних балок и оси ригеля
        'левая балка предыдущего пролета
        Dim dataLeftBeam1 As StructureElement = listExtremeBeams(0)
        Dim axisLeftBeam1 As DwgLine = dataLeftBeam1.DWGEntity
        Dim userLeftBeam1 As BeamI = dataLeftBeam1.getBeamI
        'левая балка следующего пролета
        Dim dataLeftBeam2 As StructureElement = listExtremeBeams(2)
        Dim axisLeftBeam2 As DwgLine = dataLeftBeam1.DWGEntity
        Dim userLeftBeam2 As BeamI = dataLeftBeam2.getBeamI
        'распараллеливаем левую балку предыдущего пролета
        Dim tempLeftBeam1 As DwgLine = New DwgLine()
        Dim coolEntBeam As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLeftBeam1.Offset(coolEntBeam, -1 * (userLeftBeam1.widthTopPlateLeft + OutletLeftBeam))
        If coolEntBeam.Count > 0 Then
            tempLeftBeam1 = coolEntBeam(0)
        Else
            Return False
        End If
        Dim boolRez As Boolean = False
        Dim ptIntersectLeft1 As Vector2D = MathFunction.FuncFindLineIntersection(tempLeftBeam1.StartPoint.Pos, tempLeftBeam1.EndPoint.Pos, positionRigelPointLeft1, positionRigelPointRight1, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать левый край ригеля.")
            Return False
        End If
        'распараллеливаем левую балку следующего пролета
        Dim tempLeftBeam2 As DwgLine = New DwgLine()
        coolEntBeam = New List(Of DwgEntity)
        axisLeftBeam2.Offset(coolEntBeam, -1 * (userLeftBeam2.widthTopPlateLeft + OutletLeftBeam))
        If coolEntBeam.Count > 0 Then
            tempLeftBeam2 = coolEntBeam(0)
        Else
            Return False
        End If
        Dim ptIntersectLeft2 As Vector2D = MathFunction.FuncFindLineIntersection(tempLeftBeam2.StartPoint.Pos, tempLeftBeam2.EndPoint.Pos, positionRigelPointLeft1, positionRigelPointRight1, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать левый край ригеля.")
            Return False
        End If
        Dim l1 As Double = (positionRigelPointRight1 - ptIntersectLeft1).Length
        Dim l2 As Double = (positionRigelPointRight1 - ptIntersectLeft2).Length
        Dim ptIntersectLeft As Vector2D = Nothing
        If l1 >= l2 Then
            ptIntersectLeft = ptIntersectLeft1
        Else
            ptIntersectLeft = ptIntersectLeft2
        End If

        'делаем правый край
        'левая балка предыдущего пролета
        Dim dataRightBeam1 As StructureElement = listExtremeBeams(1)
        Dim axisRightBeam1 As DwgLine = dataRightBeam1.DWGEntity
        Dim userRightBeam1 As BeamI = dataRightBeam1.getBeamI
        'левая балка следующего пролета
        Dim dataRightBeam2 As StructureElement = listExtremeBeams(3)
        Dim axisRightBeam2 As DwgLine = dataRightBeam1.DWGEntity
        Dim userRightBeam2 As BeamI = dataRightBeam2.getBeamI

        Dim tempRightBeam1 As DwgLine = New DwgLine()
        coolEntBeam = New List(Of DwgEntity)
        axisRightBeam1.Offset(coolEntBeam, userRightBeam1.widthTopPlateRight + OutletRightBeam)
        If coolEntBeam.Count > 0 Then
            tempRightBeam1 = coolEntBeam(0)
        End If
        Dim tempRightBeam2 As DwgLine = New DwgLine()
        coolEntBeam = New List(Of DwgEntity)
        axisRightBeam2.Offset(coolEntBeam, userRightBeam2.widthTopPlateRight + OutletRightBeam)
        If coolEntBeam.Count > 0 Then
            tempRightBeam2 = coolEntBeam(0)
        End If
        boolRez = False
        Dim ptIntersectRight1 As Vector2D = MathFunction.FuncFindLineIntersection(tempRightBeam1.StartPoint.Pos, tempRightBeam1.EndPoint.Pos, positionRigelPointLeft2, positionRigelPointRight2, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать правый край ригеля.")
            Return False
        End If
        boolRez = False
        Dim ptIntersectRight2 As Vector2D = MathFunction.FuncFindLineIntersection(tempRightBeam2.StartPoint.Pos, tempRightBeam2.EndPoint.Pos, positionRigelPointLeft2, positionRigelPointRight2, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать правый край ригеля.")
            Return False
        End If
        Dim l3 As Double = (positionRigelPointRight1 - ptIntersectLeft1).Length
        Dim l4 As Double = (positionRigelPointRight1 - ptIntersectLeft2).Length
        Dim ptIntersectRight As Vector2D = Nothing
        If l3 >= l4 Then
            ptIntersectRight = ptIntersectRight1
        Else
            ptIntersectRight = ptIntersectRight2
        End If
        Dim middlePoint1 As Vector3D = New Cad.Foundation.Vector3D(ptIntersectLeft, TopElevation)
        Dim middlePoint2 As Vector3D = New Cad.Foundation.Vector3D(ptIntersectRight, TopElevation)
        Dim tempAxisRigel As DwgLine = New DwgLine
        tempAxisRigel.StartPoint = middlePoint1
        tempAxisRigel.EndPoint = middlePoint2
        Lenght = Math.Round((middlePoint1 - middlePoint2).Length, 3)
        Dim directionAngleRigel As Double = tempAxisRigel.Rotation
        Dim reverseDirectionAngleRigel As Double = directionAngleRigel + Math.PI
        If reverseDirectionAngleRigel > Math.PI * 2 Then
            reverseDirectionAngleRigel -= Math.PI * 2
        End If
        Dim leftRigelDir As Double = (axisLeftBeam1.Rotation + axisLeftBeam2.Rotation) / 2
        Dim rightRigelDir As Double = (axisRightBeam1.Rotation + axisRightBeam2.Rotation) / 2
        LeftDirection = Math.Round(leftRigelDir, 6)
        RightDirection = Math.Round(rightRigelDir, 6)

        'крайняя точка начала насадки в районе шкафной стенки
        Dim leftPlateBeam As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(tempAxisRigel.StartPoint.Pos, LeftDirection, 10)
        Dim rightPlateBeam As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(tempAxisRigel.EndPoint.Pos, RightDirection, 10)
        positionRigelPointLeft1 = MathFunction.FuncFindLineIntersection(positionRigelPointLeft1, positionRigelPointRight1, leftPlateBeam, tempAxisRigel.StartPoint.Pos)
        positionRigelPointRight1 = MathFunction.FuncFindLineIntersection(positionRigelPointLeft1, positionRigelPointRight1, rightPlateBeam, tempAxisRigel.EndPoint)
        positionRigelPointRight2 = MathFunction.FuncFindLineIntersection(positionRigelPointLeft2, positionRigelPointRight2, rightPlateBeam, tempAxisRigel.EndPoint)
        positionRigelPointLeft2 = MathFunction.FuncFindLineIntersection(positionRigelPointLeft2, positionRigelPointRight2, leftPlateBeam, tempAxisRigel.StartPoint)
        '===============================================================================================================================================================
        'записываем данные
        'строим точки
        '=================================================================================================
        'первая точка
        Dim ListPointModel As New Dictionary(Of Integer, PointStructure)
        Dim numberPoint As Integer = 1
        Dim x As Double = Math.Round(positionRigelPointLeft1.X, 3)
        Dim y As Double = Math.Round(positionRigelPointLeft1.Y, 3)
        Dim z As Double = Math.Round(tempAxisRigel.StartPoint.Z, 3)
        Dim h1 As Double = Math.Round(Height, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(HeightLeftConsole, 3)
        End If
        Dim h2 As Double = 0
        Dim code As String = "leftPt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'консоли если есть
        If LengthLeftConsole > 0 Then
            Dim posLeftConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionRigelPointLeft1, positionRigelPointRight1, LengthLeftConsole)
            x = Math.Round(posLeftConsole1.X, 3)
            y = Math.Round(posLeftConsole1.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(tempAxisRigel.StartPoint, tempAxisRigel.EndPoint, posLeftConsole1), 3)
            h1 = Math.Round(Height, 3)
            h2 = 0
            code = "leftConsol1"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        If LengthRightConsole > 0 Then
            Dim posRightConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionRigelPointRight1, positionRigelPointLeft1, LengthRightConsole)
            x = Math.Round(posRightConsole1.X, 3)
            y = Math.Round(posRightConsole1.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(tempAxisRigel.StartPoint, tempAxisRigel.EndPoint, posRightConsole1), 3)
            h1 = Math.Round(Height, 3)
            h2 = 0
            code = "leftConsol2"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        'крайняя точка конца насадки в районе шкафной стенки
        x = Math.Round(positionRigelPointRight1.X, 3)
        y = Math.Round(positionRigelPointRight1.Y, 3)
        z = Math.Round(tempAxisRigel.EndPoint.Z, 3)
        h1 = Math.Round(Height, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(_heightRightConsole, 3)
        End If
        h2 = 0
        code = "leftPt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'точка средняя конец ригеля
        Dim middleEndPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionRigelPointRight1, positionRigelPointRight2)
        x = Math.Round(middleEndPoint.X, 3)
        y = Math.Round(middleEndPoint.Y, 3)
        z = Math.Round(middlePoint2.Z + HeightDrain, 3)
        h1 = Math.Round(Height + HeightDrain, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(HeightRightConsole + HeightDrain, 3)
        End If
        h2 = 0
        code = "middlePt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'крайняя точка конца насадки в районе опирания балок
        x = Math.Round(positionRigelPointRight2.X, 3)
        y = Math.Round(positionRigelPointRight2.Y, 3)
        z = Math.Round(tempAxisRigel.EndPoint.Z, 3)
        h1 = Math.Round(Height, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(_heightRightConsole, 3)
        End If
        h2 = 0
        code = "rightPt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        If LengthRightConsole > 0 Then
            Dim posRightConsole2 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionRigelPointRight2, positionRigelPointLeft2, LengthRightConsole)
            x = Math.Round(posRightConsole2.X, 3)
            y = Math.Round(posRightConsole2.Y, 3)
            z = Math.Round(tempAxisRigel.EndPoint.Z, 3)
            h1 = Math.Round(Height, 3)
            h2 = 0
            code = "rightConsol2"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        If LengthLeftConsole > 0 Then
            Dim posRightConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionRigelPointLeft2, positionRigelPointRight2, LengthLeftConsole)
            x = Math.Round(posRightConsole1.X, 3)
            y = Math.Round(posRightConsole1.Y, 3)
            z = Math.Round(tempAxisRigel.EndPoint.Z, 3)
            h1 = Math.Round(Height, 3)
            h2 = 0
            code = "rightConsol1"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        'крайняя точка конца насадки в районе шкафной стенки
        x = Math.Round(positionRigelPointLeft2.X, 3)
        y = Math.Round(positionRigelPointLeft2.Y, 3)
        z = Math.Round(tempAxisRigel.StartPoint.Z, 3)
        h1 = Math.Round(Height, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(_heightLeftConsole, 3)
        End If
        h2 = 0
        code = "rightPt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'точка средняя начало шкафной стенки
        Dim middleStartPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionRigelPointLeft1, positionRigelPointLeft2)
        x = Math.Round(middleStartPoint.X, 3)
        y = Math.Round(middleStartPoint.Y, 3)
        z = Math.Round(middlePoint1.Z + HeightDrain, 3)
        h1 = Math.Round(Height + HeightDrain, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(HeightLeftConsole + HeightDrain, 3)
        End If
        h2 = 0
        code = "middlePt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        _elementBridgePoint.ListPointModel = ListPointModel
        '===============================================================================================================
        'ось ригеля
        Dim middleLeftPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionRigelPointLeft1, positionRigelPointLeft2)
        Dim middleRightPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionRigelPointRight1, positionRigelPointRight2)
        _elementBridgePoint.StartAxisPoint = New Vector3D(middleLeftPoint, TopElevation + HeightDrain)
        _elementBridgePoint.EndAxisPoint = New Vector3D(middleRightPoint, TopElevation + HeightDrain)
        Dim centerAxisPoint As Vector2D = New Vector2D(-1, -1)
        If IsNothing(align) = False Then
            Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(align, middleLeftPoint, middleRightPoint)
            If pointIntersectCollection.Count > 0 Then
                centerAxisPoint = pointIntersectCollection.ElementAt(0)
            End If
        End If
        If centerAxisPoint.X = -1 And centerAxisPoint.Y = -1 Then
            centerAxisPoint = MathFunction.funcCalcMiddleCoordByToPoints2d(middleLeftPoint, middleRightPoint)
        End If
        _elementBridgePoint.CenterTopPoint = New Vector3D(centerAxisPoint, TopElevation)
        Return True
    End Function
    'рисование оси ригеля
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineRigel As DwgLine = Nothing
        Dim dataStructureRigel As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getAxis(dictionaryBridgeElements, NumberPillar, Number)
        If listAxis.Count = 0 Then
            dataStructureRigel = createAxisRigelPillar(idBridge)
        ElseIf listAxis.Count = 1 Then
            dataStructureRigel = listAxis.Item(0)
        Else
            dataStructureRigel = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataStructureRigel) Then Return dataStructureRigel
        axisLineRigel = dataStructureRigel.DWGEntity
        If IsNothing(axisLineRigel) = True Then axisLineRigel = New DwgLine
        If axisLineRigel.Length = 0 Then
            Dim layerAxisRigel As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisRigel As CadColor = New CadColor(7)
            Dim nameTypeLineAxisRigel As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisRigel As Integer = 1
            Dim widthTypeLineAxisRigel As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisRigel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisRigel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ригель (ось)")
            styleAxisRigel.setObjectStyle(axisLineRigel)
        End If
        'ось насадки
        axisLineRigel.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineRigel.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisRigel As Double = (axisLineRigel.StartPoint.Pos - axisLineRigel.EndPoint.Pos).Length
        If lenghtAxisRigel = 0 Then
            MsgBox("Ось ригеля имеет нулевое значение.")
            Return dataStructureRigel
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineRigel) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineRigel)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureRigel.KeyParameter = strJson
        dataStructureRigel.DWGEntity = axisLineRigel
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineRigel, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureRigel)
        Return dataStructureRigel
    End Function
    'функция возвращает все точки насадки в виде словаря 
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
    Public Function getPointByCode(ByVal code As String, Optional topElevation As Boolean = False) As Vector3D
        Dim result As Vector3D = New Vector3D
        If IsNothing(code) = False Then
            If code.Trim.Length > 0 Then
                Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
                If ListPointModel.Count > 0 Then
                    For i As Integer = 0 To ListPointModel.Count - 1
                        Dim ptStructure As PointStructure = ListPointModel.ElementAt(i).Value
                        If ptStructure.Code Like code Then
                            If topElevation = False Then
                                result = New Vector3D(ptStructure.X - ptStructure.dx, ptStructure.Y - ptStructure.dy, ptStructure.Z + ptStructure.dz)
                            Else
                                result = New Vector3D(ptStructure.X, ptStructure.Y, ptStructure.Z)
                            End If
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
        Return result
    End Function



End Class
