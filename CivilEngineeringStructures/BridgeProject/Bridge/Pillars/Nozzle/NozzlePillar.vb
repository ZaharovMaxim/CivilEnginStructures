Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports CivilEnginStructures.BridgeGeometry
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg.Bridges
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
'насадка
Public Class NozzlePillar
    Private _numberPillar As Integer 'номер опоры
    Private _number As Integer 'номер насадки в опоре
    Private _width As Double 'ширина насадки
    Private _lenght As Double 'длина насадки
    Private _firstHeight As Double 'высота насадки в начале в зоне расположения балок
    Private _secondHeight As Double 'высота насадки в месте расположения шкафной стенки
    Private _outletLeftBeam As Double 'выпуск насадки влево за левую балку
    Private _outletRightBeam As Double 'выпуск насадки вправо за правую балку
    Private _lengthLeftConsole As Double 'длина левой консоли
    Private _lengthRightConsole As Double 'длина правой консоли
    Private _heightLeftConsole As Double 'высота торца ригеля слева
    Private _heightRightConsole As Double 'высота торца ригеля слева
    Private _minElevationBeams As Double 'наименьшее расстояние от балки до насадки
    Private _widthPlateCabinetWall As Double 'ширина горизонтальной площадки под шкафную стенку
    Private _topElevation As Double 'отметка верха насадки
    Private _bottomElevation As Double 'отметка низа насадки
    Private _leftDirection As Double 'левый угол скоса короткой стороны
    Private _rightDirection As Double 'правый угол скоса короткой стороны
    Private _offsetLeftRack As Double 'смещение начало левой стойки
    Private _offsetRightRack As Double 'смещение начало правой стойки
    Private _countRacks As Integer 'количество стоек
    Private _stepRacks As Double 'шаг расстановки стоек
    Private _pileRowsDiagram As String 'схема расстановки свай ряда
    Private _pileColumnDiagram As String 'схема расстановки свай столбца
    Private _nameModel As String 'имя модели tlc
    ' Коллекции вынесены в отдельный объект
    Public _elementBridgePoint As PointsCollections
    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _number = 0
        _width = 0.0
        _lenght = 0
        _firstHeight = 0.0
        _secondHeight = 0.0
        _outletLeftBeam = 0.0
        _outletRightBeam = 0.0
        _lengthLeftConsole = 0.0
        _lengthRightConsole = 0.0
        _heightLeftConsole = 0.0
        _heightRightConsole = 0.0
        _minElevationBeams = 0.0
        _widthPlateCabinetWall = 0.0
        _topElevation = 0.0
        _bottomElevation = 0.0
        _leftDirection = 0.0
        _rightDirection = 0.0
        _offsetLeftRack = 0.0
        _offsetRightRack = 0.0
        _countRacks = 0
        _stepRacks = 0.0
        _pileRowsDiagram = String.Empty
        _pileColumnDiagram = String.Empty
        _nameModel = ""
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
    <Description("Ширина насадки, м")>
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
    <Description("Длина насадки")>
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
    <Description("Высота насадки по фасаду, м")>
    <Category("Свойства")>
    <DisplayName("Высота по фасаду")>
    Public Property FirstHeight() As Double
        Get
            Return _firstHeight
        End Get
        Set(value As Double)
            _firstHeight = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота насадки в районе шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Высота")>
    Public Property SecondHeight() As Double
        Get
            Return _secondHeight
        End Get
        Set(value As Double)
            _secondHeight = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Выпуск насадки влево за крайнюю балку, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск влево")>
    Public Property OutletLeftBeam() As Double
        Get
            Return _outletLeftBeam
        End Get
        Set(value As Double)
            _outletLeftBeam = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Выпуск насадки вправо за крайнюю балку, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск вправо")>
    Public Property OutletRightBeam() As Double
        Get
            Return _outletRightBeam
        End Get
        Set(value As Double)
            _outletRightBeam = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина левой консоли, м")>
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
    <Description("Длина правой консоли, м")>
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
    <Description("Высота левой консоли, м")>
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
    <Description("Высота правой консоли, м")>
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
    <Description("Минимальное расстояние от точки опирания балки, до верха насадки, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние от балки до насадки")>
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

    <Browsable(True)>
    <Description("Ширина площадки насадки под шкафную стенку, м")>
    <Category("Свойства")>
    <DisplayName("Ширина шк. стенки")>
    Public Property WidthPlateCabinetWall() As Double
        Get
            Return _widthPlateCabinetWall
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthPlateCabinetWall = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха насадки в районе шкафной стенки, м")>
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
    <Description("Отметка низа насадки в районе шкафной стенки, м")>
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
    <Description("Расстояние от края насадки до края левой стойки, м")>
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
    <Description("Расстояние от края насадки до края правой стойки, м")>
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
    Public Property CountRacks() As Integer
        Get
            Return _countRacks
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _countRacks = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Шаг расстановки стоек, м")>
    <Category("Свойства")>
    <DisplayName("Шаг стоек")>
    Public Property StepRacks() As Double
        Get
            Return _stepRacks
        End Get
        Set(value As Double)
            If value >= 0 Then
                _stepRacks = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Схема расстановки для рядов свай (например 500+10*500+11*600+....")>
    <Category("Свойства")>
    <DisplayName("Схема дли ряда свай")>
    Public Property PileRowsDiagram() As String
        Get
            Return _pileRowsDiagram
        End Get
        Set(value As String)
            _pileRowsDiagram = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Схема расстановки для столбцов свай (например 500+10*500+11*600+....")>
    <Category("Свойства")>
    <DisplayName("Схема дли столбцов свай")>
    Public Property PileColumnDiagram() As String
        Get
            Return _pileColumnDiagram
        End Get
        Set(value As String)
            _pileColumnDiagram = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Левый угол скоса короткой стороны")>
    <Category("Свойства")>
    <DisplayName("Левый угол скоса")>
    Public Property LeftDirection() As Double
        Get
            Return _leftDirection
        End Get
        Set(value As Double)
            _leftDirection = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Правый угол скоса короткой стороны")>
    <Category("Свойства")>
    <DisplayName("Правый угол скоса")>
    Public Property RightDirection() As Double
        Get
            Return _rightDirection
        End Get
        Set(value As Double)
            _rightDirection = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Имя модели")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    Public Property NameModel() As String
        Get
            Return _nameModel
        End Get
        Set(value As String)
            _nameModel = value
        End Set
    End Property
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую пустую насадку
    Public Shared Function createAxisNozzlePillar(ByVal idBridge As String) As StructureElement
        Dim elementNozzlePillar As StructureElement = New StructureElement()
        elementNozzlePillar.Label = "Мосты и путепроводы"
        elementNozzlePillar.ClassObject = StructureElement.classStructure.NozzlePillar
        elementNozzlePillar.Name = StructureElement.typeObject.axisNozzle
        elementNozzlePillar.Description = "Насадка (ось)"
        elementNozzlePillar.KeyParameter = ""
        elementNozzlePillar.IdElement = Guid.NewGuid.ToString
        elementNozzlePillar.IdStructure = idBridge
        elementNozzlePillar.Note = ""
        elementNozzlePillar.DWGEntity = New DwgLine()
        Return elementNozzlePillar
    End Function
    'функция ищет ось существующующей насадки
    Public Shared Function getAxis(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim dataNozzlePillar As List(Of StructureElement) = New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return dataNozzlePillar
        If numberPillar < 1 Then Return dataNozzlePillar
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisNozzle) = True Then
            Dim listAxisNozzle = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisNozzle)
            If IsNothing(listAxisNozzle) = False Then
                If listAxisNozzle.Count > 0 Then
                    For k As Integer = 0 To listAxisNozzle.Count - 1
                        Dim tempData As StructureElement = listAxisNozzle.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisNozzle As NozzlePillar = tempData.getNozzlePillar
                            If IsNothing(userAxisNozzle) = False Then
                                If numberPillar = userAxisNozzle.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisNozzle.Number
                                    If numberSubPillar = userAxisNozzle.Number Then
                                        dataNozzlePillar.Add(tempData)
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataNozzlePillar
    End Function
    'чтение данных из датогрид
    Public Shared Function readPropertiesNozzle(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal columnsDiagram As String, ByVal rowsDiagram As String, ByVal DGV_Nozzle As DataGridView) As NozzlePillar
        Dim result As NozzlePillar = New NozzlePillar
        result.NumberPillar = numbPillar
        result.Number = numbSubPillar
        If DGV_Nozzle.RowCount > 1 Then
            For i As Integer = 0 To DGV_Nozzle.RowCount - 1
                Dim tag As String = DGV_Nozzle.Rows(i).Tag
                Dim value As String = DGV_Nozzle.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_nozzle_height" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.SecondHeight = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты насадки по линии шкафной стенки.")
                        End If
                    ElseIf tag Like "firstHeight" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.FirstHeight = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты насадки по фасаду.")
                        End If
                    ElseIf tag Like "bridge_nozzle_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Val(value)
                        Else
                            MsgBox("Некорректное значение ширины насадки.")
                        End If
                    ElseIf tag Like "bridge_nozzle_cabwidth" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.WidthPlateCabinetWall = Val(value)
                        Else
                            MsgBox("Некорректное значение ширины насадки под площадкой шкафной стенки.")
                        End If
                    ElseIf tag Like "minElevationBeams" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.MinElevationBeams = Val(value)
                        Else
                            MsgBox("Некорректное значение минимальной высоты подферменника.")
                        End If
                    ElseIf tag Like "outletLeftBeam" Then
                        If IsNumeric(value) = True Then
                            result.OutletLeftBeam = Val(value)
                        Else
                            MsgBox("Некорректное значение выпуска насадки за габарит левой балки.")
                        End If
                    ElseIf tag Like "outletRightBeam" Then
                        If IsNumeric(value) = True Then
                            result.OutletRightBeam = Val(value)
                        Else
                            MsgBox("Некорректное значение выпуска насадки за габарит правой балки.")
                        End If
                    ElseIf tag Like "heightLeftConsole" Then
                        If IsNumeric(value) = True Then
                            result.HeightLeftConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты стороны левой консоли.")
                        End If
                    ElseIf tag Like "heightRightConsole" Then
                        If IsNumeric(value) = True Then
                            result.HeightRightConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты стороны правой консоли.")
                        End If
                    ElseIf tag Like "lenghtLeftConsole" Then
                        If IsNumeric(value) = True Then
                            result.LengthLeftConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение длины стороны левой консоли.")
                        End If
                    ElseIf tag Like "lenghtRightConsole" Then
                        If IsNumeric(value) = True Then
                            result.LengthRightConsole = Val(value)
                        Else
                            MsgBox("Некорректное значение длины стороны правой консоли.")
                        End If
                    ElseIf tag Like "offsetLeftRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetLeftRack = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа края стойки от левого края насадки.")
                        End If
                    ElseIf tag Like "offsetRightRack" Then
                        If IsNumeric(value) = True Then
                            result.OffsetRightRack = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа края стойки от правого края насадки.")
                        End If
                    End If
                End If
            Next i
        End If
        result.PileRowsDiagram = rowsDiagram
        result.PileColumnDiagram = columnsDiagram
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesNozzle(ByRef DGV_Nozzle As DataGridView) As Boolean
        If IsNothing(DGV_Nozzle) = True Then Return False
        If DGV_Nozzle.RowCount > 1 Then
            For j As Integer = 0 To DGV_Nozzle.RowCount - 1
                Dim tag As String = DGV_Nozzle.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_nozzle_height" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = SecondHeight
                    ElseIf tag Like "firstHeight" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = FirstHeight
                    ElseIf tag Like "bridge_nozzle_width" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "bridge_nozzle_cabwidth" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = WidthPlateCabinetWall
                    ElseIf tag Like "minElevationBeams" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = MinElevationBeams
                    ElseIf tag Like "outletLeftBeam" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = OutletLeftBeam
                    ElseIf tag Like "outletRightBeam" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = OutletRightBeam
                    ElseIf tag Like "offsetLeftRack" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = OffsetLeftRack
                    ElseIf tag Like "offsetRightRack" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = OffsetRightRack
                    ElseIf tag Like "heightLeftConsole" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = HeightLeftConsole
                    ElseIf tag Like "heightRightConsole" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = HeightRightConsole
                    ElseIf tag Like "lenghtLeftConsole" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = LengthLeftConsole
                    ElseIf tag Like "lenghtRightConsole" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = LengthRightConsole
                    ElseIf tag Like "countRacks" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = CountRacks
                    ElseIf tag Like "stepRacks" Then
                        DGV_Nozzle.Rows(j).Cells(1).Value = StepRacks
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Nozzle As DataGridView) As Boolean
        If IsNothing(DGV_Nozzle) = True Then Return False
        If DGV_Nozzle.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolBottomElevation As Boolean = False
                Dim boolLenght As Boolean = False
                Dim boolStepRacks As Boolean = False
                For i As Integer = 0 To DGV_Nozzle.RowCount - 1
                    Dim oldTag As String = DGV_Nozzle.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        DGV_Nozzle.Rows(i).Cells(1).Value = TopElevation
                        boolTopElevation = True
                    ElseIf oldTag Like "calc-BottomElevation" Then
                        DGV_Nozzle.Rows(i).Cells(1).Value = BottomElevation
                        boolBottomElevation = True
                    ElseIf oldTag Like "calc-Lenght" Then
                        DGV_Nozzle.Rows(i).Cells(1).Value = Lenght
                        boolLenght = True
                    ElseIf oldTag Like "calc-StepRacks" Then
                        DGV_Nozzle.Rows(i).Cells(1).Value = StepRacks
                        boolStepRacks = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_Nozzle.RowCount - 1
                    DGV_Nozzle.Rows.Insert(numberRow)
                    DGV_Nozzle.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Nozzle.Rows(numberRow).Cells(0).Value = "Отметка верха, м"
                    DGV_Nozzle.Rows(numberRow).Cells(1).Value = TopElevation
                    DGV_Nozzle.Rows(numberRow).Tag = "calc-TopElevation"
                End If
                If boolBottomElevation = False Then
                    Dim numberRow As Integer = DGV_Nozzle.RowCount - 1
                    DGV_Nozzle.Rows.Insert(numberRow)
                    DGV_Nozzle.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Nozzle.Rows(numberRow).Cells(0).Value = "Отметка низа, м"
                    DGV_Nozzle.Rows(numberRow).Cells(1).Value = BottomElevation
                    DGV_Nozzle.Rows(numberRow).Tag = "calc-BottomElevation"
                End If
                If boolLenght = False Then
                    Dim numberRow As Integer = DGV_Nozzle.RowCount - 1
                    DGV_Nozzle.Rows.Insert(numberRow)
                    DGV_Nozzle.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Nozzle.Rows(numberRow).Cells(0).Value = "Полная длина, м"
                    DGV_Nozzle.Rows(numberRow).Cells(1).Value = Lenght
                    DGV_Nozzle.Rows(numberRow).Tag = "calc-Lenght"
                End If
                If boolStepRacks = False Then
                    Dim numberRow As Integer = DGV_Nozzle.RowCount - 1
                    DGV_Nozzle.Rows.Insert(numberRow)
                    DGV_Nozzle.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Nozzle.Rows(numberRow).Cells(0).Value = "Шаг расстановки стоек, м"
                    DGV_Nozzle.Rows(numberRow).Cells(1).Value = StepRacks
                    DGV_Nozzle.Rows(numberRow).Tag = "calc-StepRacks"
                End If
            End If
        End If
        Return True
    End Function
    'предварительный расчет насадки
    Public Function calculateNozzle(ByVal lineCabinetWall As DwgLine, Optional align As Polyline3D = Nothing) As Boolean
        If IsNothing(lineCabinetWall) = True Then
            Dim startPointCW As Vector3D = getPointByCode("middlePt1", True)
            Dim endPointCW As Vector3D = getPointByCode("middlePt2", True)
            lineCabinetWall = New DwgLine()
            lineCabinetWall.StartPoint = startPointCW
            lineCabinetWall.EndPoint = endPointCW
        End If
        If lineCabinetWall.Length = 0 Then Return False
        '===================================================================================================================
        'строим габарит насадки
        Dim dist1 As Double = WidthPlateCabinetWall 'расстояние смещения линии шкафной стенки до края надки влево
        Dim dist2 As Double = Width - dist1
        '4 крайние точки
        Dim positionNozzlePointLeft1 As Vector2D = Nothing
        Dim positionNozzlePointLeft2 As Vector2D = Nothing
        Dim positionNozzlePointRight1 As Vector2D = Nothing
        Dim positionNozzlePointRight2 As Vector2D = Nothing
        'находим позицию насадки путем параллельного переноса линиии начала шкафной стенки
        If NumberPillar = 1 Then
            Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            lineCabinetWall.Offset(listEnt, dist1)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                positionNozzlePointLeft1 = newLine.StartPoint
                positionNozzlePointRight1 = newLine.EndPoint
            Else
                Return False
            End If
            listEnt = New List(Of DwgEntity)
            lineCabinetWall.Offset(listEnt, -1 * dist2)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                positionNozzlePointLeft2 = newLine.StartPoint
                positionNozzlePointRight2 = newLine.EndPoint
            Else
                Return False
            End If
        Else
            Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            lineCabinetWall.Offset(listEnt, -1 * dist1)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                positionNozzlePointLeft1 = newLine.StartPoint
                positionNozzlePointRight1 = newLine.EndPoint
            Else
                Return False
            End If
            listEnt = New List(Of DwgEntity)
            lineCabinetWall.Offset(listEnt, dist2)
            If listEnt.Count > 0 Then
                Dim newLine As DwgLine = listEnt(0)
                positionNozzlePointLeft2 = newLine.StartPoint
                positionNozzlePointRight2 = newLine.EndPoint
            Else
                Return False
            End If
        End If
        'крайняя точка начала насадки в районе шкафной стенки
        Dim leftPlateBeam As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineCabinetWall.StartPoint.Pos, LeftDirection, 10)
        Dim rightPlateBeam As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineCabinetWall.EndPoint.Pos, RightDirection, 10)
        Dim boolrez As Boolean = False
        positionNozzlePointLeft1 = MathFunction.FuncFindLineIntersection(positionNozzlePointLeft1, positionNozzlePointRight1, leftPlateBeam, lineCabinetWall.StartPoint.Pos, boolrez)
        If boolrez = False Then Return False
        positionNozzlePointRight1 = MathFunction.FuncFindLineIntersection(positionNozzlePointLeft1, positionNozzlePointRight1, rightPlateBeam, lineCabinetWall.EndPoint, boolrez)
        If boolrez = False Then Return False
        positionNozzlePointRight2 = MathFunction.FuncFindLineIntersection(positionNozzlePointLeft2, positionNozzlePointRight2, rightPlateBeam, lineCabinetWall.EndPoint, boolrez)
        If boolrez = False Then Return False
        positionNozzlePointLeft2 = MathFunction.FuncFindLineIntersection(positionNozzlePointLeft2, positionNozzlePointRight2, leftPlateBeam, lineCabinetWall.StartPoint, boolrez)
        '===============================================================================================================================================================
        'записываем данные
        'длина
        Dim lenghtNozzle As Double = (positionNozzlePointLeft1 - positionNozzlePointRight1).Length
        Lenght = Math.Round(lenghtNozzle, 3)
        'отметка
        TopElevation = Math.Round(lineCabinetWall.StartPoint.Z, 3)
        BottomElevation = Math.Round(lineCabinetWall.StartPoint.Z - SecondHeight, 3)
        'строим точки
        '=================================================================================================
        Dim numberPoint As Integer = 1
        Dim ListPointModel As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        'первая точка
        Dim dhFace As Double = Math.Round(SecondHeight - FirstHeight, 3)
        Dim x As Double = Math.Round(positionNozzlePointLeft1.X, 3)
        Dim y As Double = Math.Round(positionNozzlePointLeft1.Y, 3)
        Dim z As Double = Math.Round(lineCabinetWall.StartPoint.Z, 3)
        Dim h1 As Double = Math.Round(SecondHeight, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(HeightLeftConsole + dhFace, 3)
        End If
        Dim h2 As Double = 0
        Dim code As String = "leftPt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'консоли если есть
        If LengthLeftConsole > 0 Then
            Dim posLeftConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionNozzlePointLeft1, positionNozzlePointRight1, LengthLeftConsole)
            x = Math.Round(posLeftConsole1.X, 3)
            y = Math.Round(posLeftConsole1.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(lineCabinetWall.StartPoint, lineCabinetWall.EndPoint, posLeftConsole1), 3)
            h1 = Math.Round(SecondHeight, 3)
            h2 = 0
            code = "leftConsol1"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        If LengthRightConsole > 0 Then
            Dim posRightConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionNozzlePointRight1, positionNozzlePointLeft1, LengthRightConsole)
            x = Math.Round(posRightConsole1.X, 3)
            y = Math.Round(posRightConsole1.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(lineCabinetWall.StartPoint, lineCabinetWall.EndPoint, posRightConsole1), 3)
            h1 = Math.Round(SecondHeight, 3)
            h2 = 0
            code = "leftConsol2"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        'крайняя точка конца насадки в районе шкафной стенки
        x = Math.Round(positionNozzlePointRight1.X, 3)
        y = Math.Round(positionNozzlePointRight1.Y, 3)
        z = Math.Round(lineCabinetWall.EndPoint.Z, 3)
        h1 = Math.Round(SecondHeight, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(HeightRightConsole + dhFace, 3)
        End If
        h2 = 0
        code = "leftPt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'точка средняя конец шкафной стенки
        x = Math.Round(lineCabinetWall.EndPoint.X, 3)
        y = Math.Round(lineCabinetWall.EndPoint.Y, 3)
        z = Math.Round(lineCabinetWall.EndPoint.Z, 3)
        h1 = Math.Round(SecondHeight, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(HeightRightConsole + dhFace, 3)
        End If
        h2 = 0
        code = "middlePt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        'крайняя точка конца насадки в районе опирания балок
        Dim deltaHNozzle As Double = SecondHeight - FirstHeight
        x = Math.Round(positionNozzlePointRight2.X, 3)
        y = Math.Round(positionNozzlePointRight2.Y, 3)
        z = Math.Round(lineCabinetWall.EndPoint.Z - deltaHNozzle, 3)
        h1 = Math.Round(FirstHeight, 3)
        If LengthRightConsole > 0 Then
            h1 = Math.Round(HeightRightConsole, 3)
        End If
        h2 = 0
        code = "rightPt2"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1
        If LengthRightConsole > 0 Then
            Dim posRightConsole2 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionNozzlePointRight2, positionNozzlePointLeft2, LengthRightConsole)
            x = Math.Round(posRightConsole2.X, 3)
            y = Math.Round(posRightConsole2.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(lineCabinetWall.StartPoint, lineCabinetWall.EndPoint, posRightConsole2) - deltaHNozzle, 3)
            h1 = Math.Round(FirstHeight, 3)
            h2 = 0
            code = "rightConsol2"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If
        If LengthLeftConsole > 0 Then
            Dim posRightConsole1 As Vector2D = MathFunction.FuncCalcPoint2DInLine(positionNozzlePointLeft2, positionNozzlePointRight2, LengthLeftConsole)
            x = Math.Round(posRightConsole1.X, 3)
            y = Math.Round(posRightConsole1.Y, 3)
            z = Math.Round(MathFunction.FuncCalcElevationByLine(lineCabinetWall.StartPoint, lineCabinetWall.EndPoint, posRightConsole1) - deltaHNozzle, 3)
            h1 = Math.Round(FirstHeight, 3)
            h2 = 0
            code = "rightConsol1"
            ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            numberPoint += 1
        End If

        x = Math.Round(positionNozzlePointLeft2.X, 3)
        y = Math.Round(positionNozzlePointLeft2.Y, 3)
        z = Math.Round(lineCabinetWall.StartPoint.Z - deltaHNozzle, 3)
        h1 = Math.Round(FirstHeight, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(HeightLeftConsole, 3)
        End If
        h2 = 0
        code = "rightPt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        numberPoint += 1

        x = Math.Round(lineCabinetWall.StartPoint.X, 3)
        y = Math.Round(lineCabinetWall.StartPoint.Y, 3)
        z = Math.Round(lineCabinetWall.StartPoint.Z, 3)
        h1 = Math.Round(SecondHeight, 3)
        If LengthLeftConsole > 0 Then
            h1 = Math.Round(HeightLeftConsole + dhFace, 3)
        End If
        h2 = 0
        code = "middlePt1"
        ListPointModel.Add(numberPoint, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
        _elementBridgePoint.ListPointModel = ListPointModel
        '===============================================================================================================
        'ось насадки
        Dim middleLeftPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionNozzlePointLeft1, positionNozzlePointLeft2)
        Dim middleRightPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(positionNozzlePointRight1, positionNozzlePointRight2)
        _elementBridgePoint.StartAxisPoint = New Vector3D(middleLeftPoint, TopElevation)
        _elementBridgePoint.EndAxisPoint = New Vector3D(middleRightPoint, TopElevation)
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
    'рисование оси насадки
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineNozzle As DwgLine = Nothing
        Dim dataStructureNozzle As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getAxis(dictionaryBridgeElements, NumberPillar, Number)
        If listAxis.Count = 0 Then
            dataStructureNozzle = createAxisNozzlePillar(idBridge)
        ElseIf listAxis.Count = 1 Then
            dataStructureNozzle = listAxis.Item(0)
        Else
            dataStructureNozzle = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataStructureNozzle) Then Return dataStructureNozzle
        axisLineNozzle = dataStructureNozzle.DWGEntity
        If IsNothing(axisLineNozzle) = True Then axisLineNozzle = New DwgLine
        If axisLineNozzle.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Насадка (ось)")
            styleAxisNozzle.setObjectStyle(axisLineNozzle)
        End If
        'ось насадки
        axisLineNozzle.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineNozzle.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLineNozzle.StartPoint.Pos - axisLineNozzle.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось насадки имеет нулевое значение. Насадка не построена.")
            Return dataStructureNozzle
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineNozzle) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineNozzle)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureNozzle.KeyParameter = strJson
        dataStructureNozzle.DWGEntity = axisLineNozzle
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineNozzle, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureNozzle)
        Return dataStructureNozzle
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
