Imports System.ComponentModel
Imports System.Drawing
Imports System.Net
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc

Public Class RackPillar
    Public Enum TypeRack
        <Description("Трапециевидная")> Trapezoidal = 0
        <Description("Круглая")> Circle = 1
        <Description("Восьмигранная")> Octagonal = 2
        <Description("Не определено")> None = 3
    End Enum
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _numberSubPillars As Integer             ' Номер подопоры
    Private _number As Integer                       ' Номер стойки
    Private _diameter As Double                      ' Диаметр или длина стойки вдоль насадки или ригеля
    Private _widthTop As Double                      ' Ширина стойки в уровне насадки или ригеля
    Private _widthBottom As Double                   ' Ширина стойки в уровне ростверка
    Private _height As Double                        ' Высота стойки
    Private _rotation As Double                      ' Угол поворота
    Private _topSeal As Double                       ' Величина заделки в насадку или ригель
    Private _bottomSeal As Double                    ' Величина заделки в ростверк
    Private _offsetAxisX As Double                   ' Смещение верха вдоль элемента относительно расчетного значения
    Private _offsetAxisY As Double                   ' Смещение поперек элемента верха относительно левого края
    Private _topElevation As Double                  ' Отметка верха
    Private _bottomElevation As Double               ' Отметка низа
    Private _edgesParallel As Boolean                ' Грани стоек параллельны граням насадки
    Private _fixedHeight As Boolean                  ' Высота стойки задается пользователем
    Private _type As TypeRack                        ' Тип стойки
    Private _model As String                         ' Имя модели
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _numberSubPillars = 0
        _number = 0
        _diameter = 0.0
        _widthTop = 0.0
        _widthBottom = 0.0
        _height = 0.0
        _rotation = 0.0
        _topSeal = 0.0
        _bottomSeal = 0.0
        _offsetAxisX = 0.0
        _offsetAxisY = 0.0
        _topElevation = 0.0
        _bottomElevation = 0.0
        _edgesParallel = True
        _fixedHeight = False
        _type = RackPillar.TypeRack.None
        _model = String.Empty
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

    <Browsable(False)>
    <Description("Номер элемента")>
    <Category("Свойства")>
    <DisplayName("Номер элемента")>
    <[ReadOnly](True)>
    Public Property NumberSubPillars() As Integer
        Get
            Return _numberSubPillars
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberSubPillars = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Номер стойки")>
    <Category("Свойства")>
    <DisplayName("Номер стойки")>
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
    <Description("Диаметр (длина) стойки, м")>
    <Category("Свойства")>
    <DisplayName("Диаметр (длина) стойки")>
    Public Property Diameter() As Double
        Get
            Return _diameter
        End Get
        Set(value As Double)
            If value >= 0 Then
                _diameter = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина стойки в уровне насадки или ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Ширина (верх)")>
    Public Property WidthTop() As Double
        Get
            Return _widthTop
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthTop = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина стойки в уровне ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Ширина (низ)")>
    Public Property WidthBottom() As Double
        Get
            Return _widthBottom
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthBottom = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота стойки, м")>
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

    <Browsable(False)>
    <Description("Поворот стойки, град")>
    <Category("Свойства")>
    <DisplayName("Поворот")>
    Public Property Rotation() As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            ' Нормализация угла в диапазоне 0-360 градусов
            _rotation = value Mod 360
            If _rotation < 0 Then
                _rotation += 360
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Величина заделки стойки в насадку или ригель, м")>
    <Category("Свойства")>
    <DisplayName("Высота заделки")>
    Public Property TopSeal() As Double
        Get
            Return _topSeal
        End Get
        Set(value As Double)
            If value >= 0 Then
                _topSeal = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение стойки вдоль ригеля или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение X")>
    Public Property OffsetAxisX() As Double
        Get
            Return _offsetAxisX
        End Get
        Set(value As Double)
            _offsetAxisX = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение стойки поперек ригеля или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение Y")>
    Public Property OffsetAxisY() As Double
        Get
            Return _offsetAxisY
        End Get
        Set(value As Double)
            _offsetAxisY = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха стойки, м")>
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
    <Description("Отметка низа стойки, м")>
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
    <Description("Грани стойки паралельны стороне насадки или ригеля")>
    <Category("Свойства")>
    <DisplayName("Паралельность граней")>
    Public Property EdgesParallel() As Boolean
        Get
            Return _edgesParallel
        End Get
        Set(value As Boolean)
            _edgesParallel = value
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
    ' Свойство для доступа к типу стойки
    <Browsable(True)>
    <Description("Тип стойки")>
    <Category("Свойства")>
    <DisplayName("Тип стойки")>
    Public Property RackType() As TypeRack
        Get
            Return _type
        End Get
        Set(value As TypeRack)
            _type = value
        End Set
    End Property

    ' Свойство для доступа к имени модели
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
    'функции для работы со стойками
    Public Shared Function createAxisRackPillar(ByVal idBridge As String) As StructureElement
        Dim elementRack As StructureElement = New StructureElement()
        elementRack.Label = "Мосты и путепроводы"
        elementRack.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementRack.ClassObject = StructureElement.classStructure.RackPillar
        elementRack.Name = StructureElement.typeObject.axisRack
        elementRack.Description = "Стойка (ось)"
        elementRack.KeyParameter = ""
        elementRack.IdElement = Guid.NewGuid.ToString
        elementRack.IdStructure = idBridge
        elementRack.Note = ""
        elementRack.DWGEntity = New DwgLine()
        Return elementRack
    End Function
    'ищет стойки результат словарь
    Public Shared Function getRackPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As Dictionary(Of Integer, StructureElement)
        Dim dictRack As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisRack) = True Then
            Dim listAxisRack = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisRack)
            If IsNothing(listAxisRack) = False Then
                If listAxisRack.Count > 0 Then
                    For k As Integer = 0 To listAxisRack.Count - 1
                        Dim tempData As StructureElement = listAxisRack.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisRack As RackPillar = tempData.getRackPillar
                            If IsNothing(userAxisRack) = False Then
                                If numberPillar = userAxisRack.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisRack.NumberSubPillars
                                    If numberSubPillar = userAxisRack.NumberSubPillars Then
                                        Dim numRow As Integer = userAxisRack.Number
                                        If dictRack.ContainsKey(numRow) = False Then
                                            dictRack.Add(numRow, listAxisRack.Item(k))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        If dictRack.Count > 1 Then
            dictRack = dictRack.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return dictRack
    End Function

    'ищет стойки результат массив
    Public Shared Function getRackPillarToArray(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As RackPillar()
        Dim result As RackPillar() = {}
        Dim count As Integer = 0
        Dim dictRack As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisRack) = True Then
            Dim listAxisRack = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisRack)
            If IsNothing(listAxisRack) = False Then
                If listAxisRack.Count > 0 Then
                    For k As Integer = 0 To listAxisRack.Count - 1
                        Dim tempData As StructureElement = listAxisRack.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisRack As RackPillar = tempData.getRackPillar
                            If IsNothing(userAxisRack) = False Then
                                If numberPillar = userAxisRack.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisRack.NumberSubPillars
                                    If numberSubPillar = userAxisRack.NumberSubPillars Then
                                        Dim numRow As Integer = userAxisRack.Number
                                        If dictRack.ContainsKey(numRow) = False Then
                                            dictRack.Add(numRow, listAxisRack.Item(k))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        If dictRack.Count > 1 Then
            dictRack = dictRack.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            For i As Integer = 0 To dictRack.Count - 1
                Dim dataRack As StructureElement = dictRack.ElementAt(i).Value
                Dim userRack As RackPillar = dataRack.getRackPillar
                ReDim Preserve result(count)
                result(count) = userRack
                count += 1
            Next i
        End If
        Return result
    End Function
    Public Shared Function readPropertiesRack(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Rack As DataGridView, Optional ByVal fixedHeightRack As Boolean = False, Optional ByVal EgeParallel As Boolean = True, Optional ByVal userTypeRack As TypeRack = TypeRack.Circle, Optional ByVal nameModel As String = "") As RackPillar()
        Dim result As RackPillar() = {}
        Dim countArrayRack As Integer = 0
        If DGV_Rack.RowCount > 1 Then
            Dim countRack As Integer = DGV_Rack.ColumnCount - 1
            If countRack < 1 Then Return result
            For j As Integer = 1 To countRack
                Dim tempUserRack As RackPillar = New RackPillar
                tempUserRack.NumberPillar = numbPillar
                tempUserRack.NumberSubPillars = numbSubPillar
                tempUserRack.NameModel = nameModel
                For i As Integer = 0 To DGV_Rack.RowCount - 1
                    Dim tag As String = DGV_Rack.Rows(i).Tag
                    Dim value As String = DGV_Rack.Rows(i).Cells(j).Value
                    If IsNothing(tag) = False And IsNothing(value) = False Then
                        If tag Like "number" Then
                            Dim numberRack As Integer = CInt(value)
                            If numberRack < 1 Then
                                Continue For
                            Else
                                tempUserRack.Number = numberRack
                            End If
                        ElseIf tag Like "bridge_racks_diam" Then
                            If IsNumeric(value) = True And value > 0 Then
                                tempUserRack.Diameter = Val(value)
                            End If
                        ElseIf tag Like "bridge_trapracks_thickness" Then
                            If IsNumeric(value) = True And value > 0 Then
                                tempUserRack.Diameter = Val(value)
                            End If
                        ElseIf tag Like "bridge_racks_height" Then
                            If IsNumeric(value) = True And value > 0 Then
                                tempUserRack.Height = Val(value)
                            End If
                        ElseIf tag Like "topSeal" Then
                            If IsNumeric(value) = True Then
                                tempUserRack.TopSeal = Val(value)
                            End If
                        ElseIf tag Like "bridge_trapracks_wheightnozzle" Then
                            If IsNumeric(value) = True And value > 0 Then
                                tempUserRack.WidthTop = Val(value)
                            End If
                        ElseIf tag Like "bridge_trapracks_wheightgrillage" Then
                            If IsNumeric(value) = True And value > 0 Then
                                tempUserRack.WidthBottom = Val(value)
                            End If
                        ElseIf tag Like "bridge_trapracks_offsetnozzle" Then
                            'tempUserRack.offsetEdgeTop = Math.Round(Val(value), 3)
                        ElseIf tag Like "bridge_trapracks_offsetgrilage" Then
                            'tempUserRack.offsetEdgeBottom = Math.Round(Val(value), 3)
                        ElseIf tag Like "offsetAxisTopX" Then
                            If IsNumeric(value) = True Then
                                tempUserRack.OffsetAxisX = Val(value)
                            End If
                        ElseIf tag Like "offsetAxisTopY" Then
                            If IsNumeric(value) = True Then
                                tempUserRack.OffsetAxisY = Val(value)
                            End If
                        End If
                    End If
                Next i
                'тип стойки
                tempUserRack.RackType = userTypeRack
                'фиксированное положение
                If fixedHeightRack = True Then
                    tempUserRack.FixedHeight = True
                End If
                'параллельность граней
                If EgeParallel = True Then
                    tempUserRack.EdgesParallel = True
                End If
                ReDim Preserve result(countArrayRack)
                result(countArrayRack) = tempUserRack
                countArrayRack += 1
            Next j
        End If
        Return result
    End Function
    Public Function writePropertiesRack(ByRef DGV_Rack As DataGridView) As Boolean
        If IsNothing(DGV_Rack) = True Then Return False
        Dim numberRack As Integer = Number
        If Number > 0 Then
            If DGV_Rack.RowCount > 1 Then
                For j As Integer = 0 To DGV_Rack.RowCount - 1
                    Dim tag As String = DGV_Rack.Rows(j).Tag
                    If IsNothing(tag) = False Then
                        If tag Like "bridge_racks_diam" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = Diameter
                        ElseIf tag Like "bridge_trapracks_thickness" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = Diameter
                        ElseIf tag Like "bridge_racks_height" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = Height
                        ElseIf tag Like "topSeal" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = TopSeal
                        ElseIf tag Like "bridge_trapracks_wheightnozzle" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = WidthTop
                        ElseIf tag Like "bridge_trapracks_wheightgrillage" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = WidthBottom
                        ElseIf tag Like "offsetAxisTopX" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = OffsetAxisX
                        ElseIf tag Like "offsetAxisTopY" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = OffsetAxisY
                        ElseIf tag Like "type" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = StructureElement.GetDescription(RackType)
                        ElseIf tag Like "number" Then
                            DGV_Rack.Rows(j).Cells(numberRack).Value = Number
                        End If
                    End If
                Next j
            Else
                Return False
            End If
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Rack As DataGridView) As Boolean
        If IsNothing(DGV_Rack) = True Then Return False
        If DGV_Rack.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim numberTopElevation As Integer = -1
                Dim numberBottomElevation As Integer = -1
                Dim numberheightRack As Integer = -1
                Dim numberRack As Integer = -1
                For i As Integer = 0 To DGV_Rack.RowCount - 1
                    Dim oldTag As String = DGV_Rack.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        numberTopElevation = i
                    ElseIf oldTag Like "calc-BottomElevation" Then
                        numberBottomElevation = i
                    ElseIf oldTag Like "bridge_racks_height" Then
                        numberheightRack = i
                    ElseIf oldTag Like "number" Then
                        numberRack = i
                    End If
                Next i
                If numberTopElevation = -1 Then
                    numberTopElevation = DGV_Rack.RowCount - 1
                    DGV_Rack.Rows.Insert(numberTopElevation)
                    DGV_Rack.Rows(numberTopElevation).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rack.Rows(numberTopElevation).Tag = "calc-TopElevation"
                    DGV_Rack.Rows(numberTopElevation).Cells(0).Value = "Отметка верха, м"
                End If
                If numberBottomElevation = -1 Then
                    numberBottomElevation = DGV_Rack.RowCount - 1
                    DGV_Rack.Rows.Insert(numberBottomElevation)
                    DGV_Rack.Rows(numberBottomElevation).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Rack.Rows(numberBottomElevation).Tag = "calc-BottomElevation"
                    DGV_Rack.Rows(numberBottomElevation).Cells(0).Value = "Отметка низа, м"
                End If
                For i As Integer = 1 To DGV_Rack.ColumnCount - 1
                    Dim numbeRack As Integer = DGV_Rack.Rows(numberRack).Cells(i).Value
                    If numbeRack = Number Then
                        DGV_Rack.Rows(numberTopElevation).Cells(i).Value = TopElevation
                        DGV_Rack.Rows(numberBottomElevation).Cells(i).Value = BottomElevation
                        DGV_Rack.Rows(numberheightRack).Cells(i).Value = Height
                        DGV_Rack.Rows(numberheightRack).DefaultCellStyle.ForeColor = Color.Red
                        Exit For
                    End If
                Next i
            End If
        End If
        Return True
    End Function
    'предварительный расчет стоек (всех)надо проверить, высота центральной точки по веху 0
    Public Shared Function calculateRacks(ByRef userNozzle As NozzlePillar, ByRef userRigel As RigelPillar, ByRef userGrillage As GrillagePillar, ByVal arrayRack As RackPillar(), ByVal numberPillar As Integer, Optional ByVal elevationLand As Double = 0, Optional ByVal egSurface As Surface = Nothing) As Dictionary(Of Integer, RackPillar)
        Dim result As Dictionary(Of Integer, RackPillar) = New Dictionary(Of Integer, RackPillar)
        Dim listPointBridge As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        If IsNothing(arrayRack) = True Then Return result
        If arrayRack.Length = True Then Return result
        'крайние точки насадки
        Dim leftPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim leftPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim offsetLeftRack As Double = 0
        Dim offsetRightRack As Double = 0
        Dim tempLineRack As DwgLine = New DwgLine()
        Dim bottomElevation As Double = 0
        'если опора первая или последняя, но берем насадку иниче ригель
        If IsNothing(userNozzle) = False Then
            'точки насадки
            listPointBridge = userNozzle._elementBridgePoint.ListPointModel
            'смещение от края насаадки до первой опоры
            offsetLeftRack = userNozzle.OffsetLeftRack
            offsetRightRack = userNozzle.OffsetRightRack
            'ось насадки
            bottomElevation = userNozzle.BottomElevation
            tempLineRack.StartPoint = New Vector3D(userNozzle._elementBridgePoint.StartAxisPoint.Pos, bottomElevation)
            tempLineRack.EndPoint = New Vector3D(userNozzle._elementBridgePoint.EndAxisPoint.Pos, bottomElevation)
        ElseIf IsNothing(userRigel) = False Then
            'точки ригеля
            listPointBridge = userRigel._elementBridgePoint.ListPointModel
            'смещение от края ригеля до первой опоры
            offsetLeftRack = userRigel.OffsetLeftRack
            offsetRightRack = userRigel.OffsetRightRack
            'ось ригеля
            bottomElevation = userRigel.BottomElevation
            tempLineRack.StartPoint = New Vector3D(userRigel._elementBridgePoint.StartAxisPoint.Pos, bottomElevation)
            tempLineRack.EndPoint = New Vector3D(userRigel._elementBridgePoint.EndAxisPoint.Pos, bottomElevation)
        End If
        'находим крайние точки (нужны для определения разворота стойки)
        If listPointBridge.Count > 3 Then
            For i As Integer = 0 To listPointBridge.Count - 1
                Dim tempPoint As PointStructure = listPointBridge.ElementAt(i).Value
                If tempPoint.Code.IndexOf("leftPt1") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If leftPoint1.X = -1 And leftPoint1.Y = -1 And leftPoint1.Z = -1 Then
                        leftPoint1 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("leftPt2") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If leftPoint2.X = -1 And leftPoint2.Y = -1 And leftPoint2.Z = -1 Then
                        leftPoint2 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("rightPt1") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If rightPoint1.X = -1 And rightPoint1.Y = -1 And rightPoint1.Z = -1 Then
                        rightPoint1 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("rightPt2") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If rightPoint2.X = -1 And rightPoint2.Y = -1 And rightPoint2.Z = -1 Then
                        rightPoint2 = tempPt
                    End If
                End If
            Next
        End If
        If leftPoint1.X = -1 And leftPoint1.Y = -1 And leftPoint1.Z = -1 Then
            Return result
        End If
        If leftPoint2.X = -1 And leftPoint2.Y = -1 And leftPoint2.Z = -1 Then
            Return result
        End If
        If rightPoint1.X = -1 And rightPoint1.Y = -1 And rightPoint1.Z = -1 Then
            Return result
        End If
        If rightPoint2.X = -1 And rightPoint2.Y = -1 And rightPoint2.Z = -1 Then
            Return result
        End If
        If tempLineRack.Length = 0 Then Return result
        'направление расстановки стоек в прямом направлении (слева на право)
        Dim anglePillar As Double = tempLineRack.Rotation
        'обратное направление
        Dim reverseAnglePillar As Double = anglePillar + Math.PI
        If reverseAnglePillar > Math.PI * 2 Then
            reverseAnglePillar -= Math.PI * 2
        End If
        'блок для расчета поправок в угол разворота стойки который связан с непараллельностью торцов насадки
        '===================================================================================================
        '1 торец левый
        Dim tempLine1 As DwgLine = New DwgLine()
        tempLine1.StartPoint = leftPoint1
        tempLine1.EndPoint = rightPoint1
        'второй торец правый
        Dim tempLine2 As DwgLine = New DwgLine()
        tempLine2.StartPoint = leftPoint2
        tempLine2.EndPoint = rightPoint2
        'делаем расчет угла поворорота для каждой стойки
        Dim deltaAngle As Double = 0
        Dim angleN As Double = 0 'угол по ходу пикетажа перпендикулярен оси опоры
        Dim angleRN As Double = Math.PI 'угол против хода пикетажа перпендикулярен оси опоры
        'рассчитываем средний угол доворота в каждую стойку, за счет неравенства направлений слева и справа
        If arrayRack.Length > 1 Then
            deltaAngle = (tempLine2.Rotation - tempLine1.Rotation) / (arrayRack.Length - 1)
        Else
            deltaAngle = (tempLine2.Rotation - tempLine1.Rotation) / 2
        End If
        '====================================================================================================
        angleN = tempLine1.Rotation
        angleRN = angleN + Math.PI
        If angleRN > Math.PI * 2 Then
            angleRN -= Math.PI * 2
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'достаем последовательно стойки
        If arrayRack.Length > 0 Then
            Dim userFirstRack As RackPillar = arrayRack(0) 'первая стойка
            Dim userLastRack As RackPillar = arrayRack(arrayRack.Length - 1) 'последняя стойка
            Dim diamertRackLeft As Double = userFirstRack.Diameter 'диаметр первой стойки
            Dim diamertRackRight As Double = userLastRack.Diameter 'диаметр последней стойки
            'добавляем полдиаметра к отступу
            Dim leftOffset As Double = offsetLeftRack + diamertRackLeft / 2
            Dim rightOffset As Double = offsetRightRack + diamertRackRight / 2
            Dim stepRack As Double = 0
            Dim countRack As Integer = arrayRack.Length
            Dim distLeft As Double = 0
            '===================================================================================================
            'расчет высоты стойки
            Dim hEarth As Double = 9999999
            If elevationLand <> 0 Then
                'отметка земли фиксирована
                hEarth = elevationLand
            ElseIf IsNothing(egSurface) = False Then
                'определяем наименьшую отметку земли
                Dim listPoly As List(Of Vector2D) = New List(Of Vector2D)
                listPoly.Add(tempLineRack.StartPoint.Pos)
                listPoly.Add(tempLineRack.EndPoint.Pos)
                Dim sectLinePillar As Sfc.Sections.Section = egSurface.CreateSection(listPoly, Sfc.Sections.SectionFlags.FilterRibs)
                If sectLinePillar.Count > 1 Then
                    For i As Integer = 0 To sectLinePillar.Count - 1
                        Dim sect As Sfc.Sections.SectionNode = sectLinePillar.ElementAt(i)
                        Dim elev As Double = Math.Round(sect.Vertex.Y, 3)
                        If elev < hEarth Then
                            hEarth = elev
                        End If
                    Next
                End If
            Else
                Dim hRack As Double = userFirstRack.Height
                If IsNothing(userNozzle) = False Then
                    hEarth = userNozzle.BottomElevation - hRack
                ElseIf IsNothing(userRigel) = False Then
                    hEarth = userRigel.BottomElevation - hRack
                End If
            End If
            If hEarth = 9999999 Then
                MsgBox("Не удалось определить отметку земли при расчете позиции стоек. Возможно поверхность существующей земли не корректная. Задайте отметку вручную.")
                Return result
            End If
            'начинаем расчет стоек
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If arrayRack.Length > 1 Then 'число стоек более 1
                Dim startPtTopRack As Vector2D = New Vector2D
                Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
                'если это опора не последняя
                If IsNothing(userNozzle) = False And numberPillar > 1 Then
                    tempLine1.Offset(listEnt, -1 * leftOffset)
                    If listEnt.Count > 0 Then
                        Dim newLine As DwgLine = listEnt(0)
                        startPtTopRack = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempLineRack.StartPoint.Pos, tempLineRack.EndPoint.Pos)
                    End If
                    Dim endPtTopRack As Vector2D = New Vector2D
                    listEnt = New List(Of DwgEntity)
                    tempLine2.Offset(listEnt, rightOffset)
                    If listEnt.Count > 0 Then
                        Dim newLine As DwgLine = listEnt(0)
                        endPtTopRack = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempLineRack.StartPoint.Pos, tempLineRack.EndPoint.Pos)
                    End If
                    Dim elevStart As Double = MathFunction.FuncCalcElevationByLine(leftPoint1, rightPoint1, startPtTopRack)
                    Dim elevEnd As Double = MathFunction.FuncCalcElevationByLine(leftPoint1, rightPoint1, endPtTopRack)
                    tempLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(startPtTopRack, bottomElevation)
                    tempLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtTopRack, bottomElevation)
                    Dim lenRack As Double = (endPtTopRack - startPtTopRack).Length
                    stepRack = lenRack / (countRack - 1)
                Else
                    tempLine1.Offset(listEnt, leftOffset)
                    If listEnt.Count > 0 Then
                        Dim newLine As DwgLine = listEnt(0)
                        startPtTopRack = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempLineRack.StartPoint.Pos, tempLineRack.EndPoint.Pos)
                    End If
                    Dim endPtTopRack As Vector2D = New Vector2D
                    listEnt = New List(Of DwgEntity)
                    tempLine2.Offset(listEnt, -1 * rightOffset)
                    If listEnt.Count > 0 Then
                        Dim newLine As DwgLine = listEnt(0)
                        endPtTopRack = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempLineRack.StartPoint.Pos, tempLineRack.EndPoint.Pos)
                    End If
                    Dim elevStart As Double = MathFunction.FuncCalcElevationByLine(leftPoint1, rightPoint1, startPtTopRack)
                    Dim elevEnd As Double = MathFunction.FuncCalcElevationByLine(leftPoint1, rightPoint1, endPtTopRack)
                    tempLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(startPtTopRack, bottomElevation)
                    tempLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtTopRack, bottomElevation)
                    Dim lenRack As Double = (endPtTopRack - startPtTopRack).Length
                    stepRack = lenRack / (countRack - 1)
                End If
            Else
                stepRack = tempLineRack.Length / 2
                distLeft = stepRack
            End If
            'записываем расчетные значения в память
            If IsNothing(userNozzle) = False Then
                userNozzle.StepRacks = Math.Round(stepRack, 3)
                userNozzle.CountRacks = countRack
            End If
            If IsNothing(userRigel) = False Then
                userRigel.StepRack = Math.Round(stepRack, 3)
                userRigel.CountRack = countRack
            End If

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'начинаем расчет позиции каждой стойки
            For i As Integer = 0 To arrayRack.Length - 1
                Dim userRack As RackPillar = arrayRack(i)
                Dim numberRack As Integer = userRack.Number
                'корректируем угол стойки (начиная от левой опоры путем добавления поправки в угол выходим на правую опору
                Dim angleN_v As Double = angleN + i * deltaAngle
                Dim angleRN_v As Double = angleRN + i * deltaAngle
                'начасльная позиция на оси
                Dim offsetStartDistAxisRack As Double = distLeft + stepRack * i + userRack.OffsetAxisX
                'смотрим если есть индивидуальное смещение, то применяем его
                'делаем параллельный перенос левого края насадки или ригеля (dX)
                Dim insPtTopRack As Vector2D = MathFunction.FuncCalcPoint2DInLine(tempLineRack.StartPoint.Pos, tempLineRack.EndPoint.Pos, offsetStartDistAxisRack)
                'рассчитываем смещение по оси Y
                If userRack.OffsetAxisY <> 0 Then
                    insPtTopRack = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleN_v, userRack.OffsetAxisY) 'по ходу пикетажа
                End If
                'рассчитываем высоту стойки
                '===================================================================================================================================
                'отметка в уровне ригеля или насадки
                Dim topElevationRack As Double = MathFunction.FuncCalcElevationByLine(tempLineRack.StartPoint, tempLineRack.EndPoint, insPtTopRack)
                'добавляем заделку
                topElevationRack += userRack.TopSeal
                'точка верха стойки с учетом заглубления в насадку или ригель
                userRack.TopElevation = topElevationRack
                'расчитываем низ стойки
                Dim bottomHRack As Double = 0
                If userRack.FixedHeight = True Then
                    'фиксированная высота
                    bottomHRack = topElevationRack - userRack.Height
                Else
                    'расчет от поверхности земли
                    bottomHRack = hEarth - userGrillage.DepthFoundation
                End If
                userRack.BottomElevation = Math.Round(bottomHRack, 3)
                'расчет высоты стойки от ригеля до насадки
                userRack.Height = Math.Round(topElevationRack - bottomHRack, 3)
                '=================================================================================================================================
                'записываем угол
                If numberPillar = 1 Then
                    userRack.Rotation = angleN_v
                Else
                    userRack.Rotation = angleRN_v
                End If
                userRack.Rotation = Math.Round(angleN_v, 6)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ось стойки
                Dim x As Double = Math.Round(insPtTopRack.X, 3)
                Dim y As Double = Math.Round(insPtTopRack.Y, 3)
                Dim z As Double = Math.Round(topElevationRack, 3)
                Dim dx As Double = 0
                Dim dy As Double = 0
                Dim dz As Double = userRack.Height
                Dim code As String = "center"
                userRack._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                userRack._elementBridgePoint.CenterBottomPoint = New Vector3D(x, y, z - userRack.Height)
                userRack._elementBridgePoint.EndAxisPoint = New Vector3D(x, y, z)
                userRack._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z - userRack.Height)
                If Not (userRack.RackType = TypeRack.Trapezoidal) Then
                    'делаем расчет 4 точек по окружности
                    Dim leftPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, reverseAnglePillar, userRack.Diameter / 2)
                    Dim rightPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, anglePillar, userRack.Diameter / 2)
                    Dim angleFront As Double = anglePillar - Math.PI / 2
                    If angleFront < 0 Then
                        angleFront += Math.PI * 2
                    End If
                    Dim frontPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleFront, userRack.Diameter / 2)
                    Dim angleBack As Double = anglePillar + Math.PI / 2
                    If angleBack > Math.PI * 2 Then
                        angleBack -= Math.PI * 2
                    End If
                    Dim backPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleBack, userRack.Diameter / 2)
                    Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
                    Dim listPointRectangle As List(Of Vector2D) = BridgeGeometry.calculatePointBox(leftPointCircle, rightPointCircle, frontPointCircle, backPointCircle)
                    leftPointCircle = listPointRectangle(1)
                    x = Math.Round(leftPointCircle.X, 3)
                    y = Math.Round(leftPointCircle.Y, 3)
                    z = Math.Round(topElevationRack, 3)
                    dx = 0
                    dy = 0
                    dz = Math.Round(userRack.Height, 3)
                    code = "leftPt1"
                    Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                    listModelPoint.Add(1, pointModel)
                    rightPointCircle = listPointRectangle(2)
                    x = Math.Round(rightPointCircle.X, 3)
                    y = Math.Round(rightPointCircle.Y, 3)
                    z = Math.Round(topElevationRack, 3)
                    dx = 0
                    dy = 0
                    dz = Math.Round(userRack.Height, 3)
                    code = "leftPt2"
                    pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                    listModelPoint.Add(2, pointModel)
                    frontPointCircle = listPointRectangle(3)
                    x = Math.Round(frontPointCircle.X, 3)
                    y = Math.Round(frontPointCircle.Y, 3)
                    z = Math.Round(topElevationRack, 3)
                    dx = 0
                    dy = 0
                    dz = Math.Round(userRack.Height, 3)
                    code = "rightPt2"
                    pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                    listModelPoint.Add(3, pointModel)
                    backPointCircle = listPointRectangle(0)
                    x = Math.Round(backPointCircle.X, 3)
                    y = Math.Round(backPointCircle.Y, 3)
                    z = Math.Round(topElevationRack, 3)
                    dx = 0
                    dy = 0
                    dz = Math.Round(userRack.Height, 3)
                    code = "rightPt1"
                    pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                    listModelPoint.Add(4, pointModel)
                    userRack._elementBridgePoint.ListPointModel = listModelPoint
                    'ElseIf userRack.RackType = TypeRack.Octagonal Then
                    '    Dim listVertexPolygon As List(Of Vector3D) = BridgeGeometry.calculateVertexPolygon(userRack._elementBridgePoint.CenterTopPoint, userRack.Diameter / 2, 8, userRack.Rotation)
                    '    If listVertexPolygon.Count > 2 Then
                    '        For j As Integer = 0 To listVertexPolygon.Count - 1

                    '        Next j
                    '    End If
                Else
                    'стойка трапециевидная
                    If userRack.EdgesParallel = False Then 'грани не параллельны оси ригеля или насадки
                        Dim ofsetTopCenter As Double = userRack.WidthTop / 2
                        Dim backPointTr As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleRN_v, userRack.WidthTop / 2)
                        Dim frontPointTr As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleN_v, userRack.WidthTop / 2)
                        Dim angleFront As Double = angleN_v - Math.PI / 2
                        If angleFront < 0 Then
                            angleFront += Math.PI * 2
                        End If
                        Dim angleBack As Double = angleN_v + Math.PI / 2
                        If angleBack > Math.PI * 2 Then
                            angleBack -= Math.PI * 2
                        End If
                        Dim pointTopLeft1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(backPointTr, angleBack, userRack.Diameter / 2)
                        Dim pointTopLeft2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(backPointTr, angleFront, userRack.Diameter / 2)
                        Dim pointTopRight1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr, angleBack, userRack.Diameter / 2)
                        Dim pointTopRight2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr, angleFront, userRack.Diameter / 2)
                        'по низу
                        Dim frontPointTr2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleN_v, userRack.WidthBottom - userRack.WidthTop / 2)
                        Dim pointBottomRight1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr2, angleBack, userRack.Diameter / 2)
                        Dim pointBottomRight2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr2, angleFront, userRack.Diameter / 2)

                        Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
                        x = Math.Round(pointTopLeft1Circle.X, 3)
                        y = Math.Round(pointTopLeft1Circle.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = 0
                        dy = 0
                        dz = Math.Round(userRack.Height, 3)
                        code = "leftPt1"
                        Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(3, pointModel)

                        x = Math.Round(pointTopLeft2Circle.X, 3)
                        y = Math.Round(pointTopLeft2Circle.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = 0
                        dy = 0
                        dz = Math.Round(userRack.Height, 3)
                        code = "leftPt2"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(2, pointModel)

                        x = Math.Round(pointTopRight2Circle.X, 3)
                        y = Math.Round(pointTopRight2Circle.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = Math.Round(pointTopRight2Circle.X - pointBottomRight2Circle.X, 3)
                        dy = Math.Round(pointTopRight2Circle.Y - pointBottomRight2Circle.Y, 3)
                        dz = Math.Round(userRack.Height, 3)
                        code = "rightPt2"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(1, pointModel)

                        x = Math.Round(pointTopRight1Circle.X, 3)
                        y = Math.Round(pointTopRight1Circle.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = Math.Round(pointTopRight1Circle.X - pointBottomRight1Circle.X, 3)
                        dy = Math.Round(pointTopRight1Circle.Y - pointBottomRight1Circle.Y, 3)
                        dz = Math.Round(userRack.Height, 3)
                        code = "rightPt1"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(4, pointModel)
                        userRack._elementBridgePoint.ListPointModel = listModelPoint
                    Else
                        Dim insPtBottomRack1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleRN_v, userRack.WidthTop / 2)
                        Dim insPtBottomRack2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtBottomRack1, angleN_v, userRack.WidthBottom)
                        Dim centerGrillage As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(insPtBottomRack1, insPtBottomRack2)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'делаем расчет 4 точек
                        'находим расстояние от края насадки до края стойки слева
                        Dim offsetTop As Double = userRack.WidthTop / 2
                        Dim insPoint1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleN_v, offsetTop)
                        Dim insPoint2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtTopRack, angleRN_v, offsetTop)
                        Dim tempAxisLineRack As DwgLine = New DwgLine
                        tempAxisLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(insPoint2, 0)
                        tempAxisLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(insPoint1, 0)

                        Dim leftPointTop1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint1, anglePillar, userRack.Diameter / 2)
                        Dim rightPointTop1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint1, reverseAnglePillar, userRack.Diameter / 2)
                        Dim tempBottomLineRack As DwgLine = New DwgLine
                        tempBottomLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(leftPointTop1, 0)
                        tempBottomLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(rightPointTop1, 0)

                        Dim leftPointTop2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint2, anglePillar, userRack.Diameter / 2)
                        Dim rightPointTop2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint2, reverseAnglePillar, userRack.Diameter / 2)
                        Dim tempTopLineRack As DwgLine = New DwgLine
                        tempTopLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(leftPointTop2, 0)
                        tempTopLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(rightPointTop2, 0)
                        'делаем параллельный перенос оси стойки на половину диаметра
                        Dim positionRackTopPointLeft1 As Vector2D = New Vector2D()
                        Dim positionRackTopPointLeft2 As Vector2D = New Vector2D()

                        Dim positionRackTopPointRight1 As Vector2D = New Vector2D()
                        Dim positionRackTopPointRight2 As Vector2D = New Vector2D()

                        Dim tempD As Double = userRack.Diameter / 2
                        Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
                        tempAxisLineRack.Offset(listEnt, tempD)
                        If listEnt.Count > 0 Then
                            Dim newLine As DwgLine = listEnt(0)
                            positionRackTopPointRight2 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempBottomLineRack.StartPoint.Pos, tempBottomLineRack.EndPoint.Pos)
                            positionRackTopPointLeft2 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempTopLineRack.StartPoint.Pos, tempTopLineRack.EndPoint.Pos)
                        End If
                        listEnt = New List(Of DwgEntity)
                        tempAxisLineRack.Offset(listEnt, -1 * tempD)
                        If listEnt.Count > 0 Then
                            Dim newLine As DwgLine = listEnt(0)
                            positionRackTopPointRight1 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempBottomLineRack.StartPoint.Pos, tempBottomLineRack.EndPoint.Pos)
                            positionRackTopPointLeft1 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempTopLineRack.StartPoint.Pos, tempTopLineRack.EndPoint.Pos)
                        End If

                        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        'расчитываем точки по ростверку
                        Dim rightPointBottom1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(positionRackTopPointLeft1, angleN_v, userRack.WidthBottom)
                        Dim rightPointBottom2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(positionRackTopPointLeft2, angleN_v, userRack.WidthBottom)

                        Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
                        x = Math.Round(positionRackTopPointLeft1.X, 3)
                        y = Math.Round(positionRackTopPointLeft1.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = 0
                        dy = 0
                        dz = Math.Round(userRack.Height, 3)
                        code = "leftPt1"
                        Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(3, pointModel)

                        x = Math.Round(positionRackTopPointLeft2.X, 3)
                        y = Math.Round(positionRackTopPointLeft2.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = 0
                        dy = 0
                        dz = Math.Round(userRack.Height, 3)
                        code = "leftPt2"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(2, pointModel)

                        x = Math.Round(positionRackTopPointRight2.X, 3)
                        y = Math.Round(positionRackTopPointRight2.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = Math.Round(positionRackTopPointRight2.X - rightPointBottom2.X, 3)
                        dy = Math.Round(positionRackTopPointRight2.Y - rightPointBottom2.Y, 3)
                        dz = Math.Round(userRack.Height, 3)
                        code = "rightPt2"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(1, pointModel)

                        x = Math.Round(positionRackTopPointRight1.X, 3)
                        y = Math.Round(positionRackTopPointRight1.Y, 3)
                        z = Math.Round(topElevationRack, 3)
                        dx = Math.Round(positionRackTopPointRight1.X - rightPointBottom1.X, 3)
                        dy = Math.Round(positionRackTopPointRight1.Y - rightPointBottom1.Y, 3)
                        dz = Math.Round(userRack.Height, 3)
                        code = "rightPt1"
                        pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        listModelPoint.Add(4, pointModel)
                        userRack._elementBridgePoint.ListPointModel = listModelPoint
                    End If
                End If
                If result.ContainsKey(numberRack) = False Then
                    result.Add(numberRack, userRack)
                End If
            Next i
        End If
        Return result
    End Function
    'предварительный расчет отделной стойки
    Public Function calculateRack(ByVal userNozzle As NozzlePillar, ByVal userRigel As RigelPillar, ByVal userGrillage As GrillagePillar) As Boolean
        Dim result As Boolean = False
        Dim listPointBridge As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        'крайние точки насадки
        Dim leftPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim leftPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, -1)
        Dim offsetLeftRack As Double = 0
        Dim offsetRightRack As Double = 0
        Dim centerLineRack As DwgLine = New DwgLine()
        Dim bottomElevation As Double = 0
        'если опора первая или последняя, но берем насадку иниче ригель
        If IsNothing(userNozzle) = False Then
            'точки насадки
            listPointBridge = userNozzle._elementBridgePoint.ListPointModel
            'смещение от края насаадки до первой опоры
            offsetLeftRack = userNozzle.OffsetLeftRack
            offsetRightRack = userNozzle.OffsetRightRack
            'ось насадки
            centerLineRack.StartPoint = New Vector3D(userNozzle._elementBridgePoint.StartAxisPoint.Pos, userNozzle.BottomElevation)
            centerLineRack.EndPoint = New Vector3D(userNozzle._elementBridgePoint.EndAxisPoint.Pos, userNozzle.BottomElevation)
            bottomElevation = userNozzle.BottomElevation
        ElseIf IsNothing(userRigel) = False Then
            'точки ригеля
            listPointBridge = userRigel._elementBridgePoint.ListPointModel
            'смещение от края ригеля до первой опоры
            offsetLeftRack = userRigel.OffsetLeftRack
            offsetRightRack = userRigel.OffsetRightRack
            'ось ригеля
            centerLineRack.StartPoint = New Vector3D(userRigel._elementBridgePoint.StartAxisPoint.Pos, userRigel.BottomElevation)
            centerLineRack.EndPoint = New Vector3D(userRigel._elementBridgePoint.EndAxisPoint.Pos, userRigel.BottomElevation)
            bottomElevation = userRigel.BottomElevation
        End If
        'находим крайние точки (нужны для определения разворота стойки)
        If listPointBridge.Count > 3 Then
            For i As Integer = 0 To listPointBridge.Count - 1
                Dim tempPoint As PointStructure = listPointBridge.ElementAt(i).Value
                If tempPoint.Code.IndexOf("leftPt1") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If leftPoint1.X = -1 And leftPoint1.Y = -1 And leftPoint1.Z = -1 Then
                        leftPoint1 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("leftPt2") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If leftPoint2.X = -1 And leftPoint2.Y = -1 And leftPoint2.Z = -1 Then
                        leftPoint2 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("rightPt1") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If rightPoint1.X = -1 And rightPoint1.Y = -1 And rightPoint1.Z = -1 Then
                        rightPoint1 = tempPt
                    End If
                ElseIf tempPoint.Code.IndexOf("rightPt2") > -1 Then
                    Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X - tempPoint.dx, tempPoint.Y - tempPoint.dy, tempPoint.Z + tempPoint.dz)
                    If rightPoint2.X = -1 And rightPoint2.Y = -1 And rightPoint2.Z = -1 Then
                        rightPoint2 = tempPt
                    End If
                End If
            Next
        End If
        If leftPoint1.X = -1 And leftPoint1.Y = -1 And leftPoint1.Z = -1 Then
            Return result
        End If
        If leftPoint2.X = -1 And leftPoint2.Y = -1 And leftPoint2.Z = -1 Then
            Return result
        End If
        If rightPoint1.X = -1 And rightPoint1.Y = -1 And rightPoint1.Z = -1 Then
            Return result
        End If
        If rightPoint2.X = -1 And rightPoint2.Y = -1 And rightPoint2.Z = -1 Then
            Return result
        End If
        If centerLineRack.Length = 0 Then Return result
        'направление расстановки стоек в прямом направлении (слева на право)
        Dim anglePillar As Double = centerLineRack.Rotation
        'перпендикулярное направление
        Dim normalAnglePillar As Double = anglePillar + Math.PI / 2
        If normalAnglePillar > Math.PI * 2 Then
            normalAnglePillar -= Math.PI * 2
        End If
        'обратное направление
        Dim reverseAnglePillar As Double = anglePillar + Math.PI
        If reverseAnglePillar > Math.PI * 2 Then
            reverseAnglePillar -= Math.PI * 2
        End If
        'блок для расчета поправок в угол разворота стойки который связан с непараллельностью торцов насадки
        '===================================================================================================
        '1 торец левый
        Dim tempLine1 As DwgLine = New DwgLine()
        tempLine1.StartPoint = leftPoint1
        tempLine1.EndPoint = rightPoint1
        'второй торец правый
        Dim tempLine2 As DwgLine = New DwgLine()
        tempLine2.StartPoint = leftPoint2
        tempLine2.EndPoint = rightPoint2
        'делаем расчет угла поворорота для каждой стойки
        Dim angleN As Double = 0 'угол по ходу пикетажа перпендикулярен оси опоры
        Dim angleRN As Double = Math.PI 'угол против хода пикетажа перпендикулярен оси опоры
        '====================================================================================================
        angleN = tempLine1.Rotation
        angleRN = angleN + Math.PI
        If angleRN > Math.PI * 2 Then
            angleRN -= Math.PI * 2
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'верх стойки
        Dim centerRack As Vector3D = _elementBridgePoint.CenterTopPoint
        'расчет высоты стойки
        Dim topElevationRack As Double = topElevationRack
        Dim bottomElevationRack As Double = bottomElevationRack
        If IsNothing(userNozzle) = False Then
            TopElevation = userNozzle.BottomElevation + TopSeal
            topElevationRack = TopElevation
        ElseIf IsNothing(userRigel) = False Then
            TopElevation = userRigel.BottomElevation + TopSeal
            topElevationRack = TopElevation
        End If
        If IsNothing(userGrillage) = False Then
            bottomElevation = userGrillage.TopElevation
            bottomElevationRack = bottomElevation
        Else
            bottomElevation = TopElevation - Height
            bottomElevationRack = bottomElevation
        End If
        'начинаем расчет стоек
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'корректируем угол стойки (начиная от левой опоры путем добавления поправки в угол выходим на правую опору
        Dim angleN_v As Double = Rotation
        Dim angleRN_v As Double = Rotation + Math.PI
        If angleRN_v > Math.PI * 2 Then
            angleRN_v -= Math.PI * 2
        End If
        'делаем параллельный перенос вдоль ригеля или насадки
        Dim centerRack2d As Vector2D = centerRack.Pos
        If OffsetAxisX <> 0 Then
            centerRack2d = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack.Pos, anglePillar, OffsetAxisX)
        End If
        'рассчитываем смещение по оси Y
        If OffsetAxisY <> 0 Then
            centerRack2d = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, normalAnglePillar, OffsetAxisY)
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'рассчитываем смещение стойки по низу
        If Not (RackType = TypeRack.Trapezoidal) Then 'ось стойки круглой или восьмигранной
            'делаем расчет 4 точек по окружности
            Dim leftPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, reverseAnglePillar, Diameter / 2)
            Dim rightPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, anglePillar, Diameter / 2)
            Dim angleFront As Double = anglePillar - Math.PI / 2
            If angleFront < 0 Then
                angleFront += Math.PI * 2
            End If
            Dim frontPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleFront, Diameter / 2)
            Dim angleBack As Double = anglePillar + Math.PI / 2
            If angleBack > Math.PI * 2 Then
                angleBack -= Math.PI * 2
            End If
            Dim backPointCircle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleBack, Diameter / 2)
            Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
            Dim listPointRectangle As List(Of Vector2D) = BridgeGeometry.calculatePointBox(leftPointCircle, rightPointCircle, frontPointCircle, backPointCircle)
            leftPointCircle = listPointRectangle(1)
            Dim x As Double = Math.Round(leftPointCircle.X, 3)
            Dim y As Double = Math.Round(leftPointCircle.Y, 3)
            Dim z As Double = Math.Round(topElevationRack, 3)
            Dim dx As Double = 0
            Dim dy As Double = 0
            Dim dz As Double = Math.Round(Height, 3)
            Dim code As String = "leftPt1"
            Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(1, pointModel)
            rightPointCircle = listPointRectangle(2)
            x = Math.Round(rightPointCircle.X, 3)
            y = Math.Round(rightPointCircle.Y, 3)
            z = Math.Round(topElevationRack, 3)
            dx = 0
            dy = 0
            dz = Math.Round(Height, 3)
            code = "leftPt2"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(2, pointModel)
            frontPointCircle = listPointRectangle(3)
            x = Math.Round(frontPointCircle.X, 3)
            y = Math.Round(frontPointCircle.Y, 3)
            z = Math.Round(topElevationRack, 3)
            dx = 0
            dy = 0
            dz = Math.Round(Height, 3)
            code = "rightPt2"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(3, pointModel)
            backPointCircle = listPointRectangle(0)
            x = Math.Round(backPointCircle.X, 3)
            y = Math.Round(backPointCircle.Y, 3)
            z = Math.Round(topElevationRack, 3)
            dx = 0
            dy = 0
            dz = Math.Round(Height, 3)
            code = "rightPt1"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(4, pointModel)
            _elementBridgePoint.ListPointModel = listModelPoint

            x = Math.Round(centerRack.X, 3)
            y = Math.Round(centerRack.Y, 3)
            z = Math.Round(topElevationRack, 3)
            dx = 0
            dy = 0
            dz = Height
            code = "center"
            _elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
            _elementBridgePoint.CenterBottomPoint = New Vector3D(x, y, z - Height)
        Else
            'стойка трапециевидная
            If EdgesParallel = False Then 'грани не параллельны оси ригеля или насадки
                Dim ofsetTopCenter As Double = WidthTop / 2
                Dim x As Double = Math.Round(centerRack.X, 3)
                Dim y As Double = Math.Round(centerRack.Y, 3)
                Dim z As Double = Math.Round(topElevationRack, 3)
                Dim dx As Double = 0
                Dim dy As Double = 0
                Dim dz As Double = Height
                Dim code As String = "center"
                _elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                _elementBridgePoint.CenterBottomPoint = New Vector3D(x, y, z - Height)
                Dim backPointTr As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack, angleRN_v, WidthTop / 2)
                Dim frontPointTr As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack, angleN_v, WidthTop / 2)
                Dim angleFront As Double = angleN_v - Math.PI / 2
                If angleFront < 0 Then
                    angleFront += Math.PI * 2
                End If
                Dim angleBack As Double = angleN_v + Math.PI / 2
                If angleBack > Math.PI * 2 Then
                    angleBack -= Math.PI * 2
                End If
                Dim pointTopLeft1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(backPointTr, angleBack, Diameter / 2)
                Dim pointTopLeft2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(backPointTr, angleFront, Diameter / 2)
                Dim pointTopRight1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr, angleBack, Diameter / 2)
                Dim pointTopRight2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr, angleFront, Diameter / 2)
                'по низу
                Dim frontPointTr2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleN_v, WidthBottom - WidthTop / 2)
                Dim pointBottomRight1Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr2, angleBack, Diameter / 2)
                Dim pointBottomRight2Circle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(frontPointTr2, angleFront, Diameter / 2)

                Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
                x = Math.Round(pointTopLeft1Circle.X, 3)
                y = Math.Round(pointTopLeft1Circle.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = 0
                dy = 0
                dz = Math.Round(Height, 3)
                code = "leftPt1"
                Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(3, pointModel)

                x = Math.Round(pointTopLeft2Circle.X, 3)
                y = Math.Round(pointTopLeft2Circle.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = 0
                dy = 0
                dz = Math.Round(Height, 3)
                code = "leftPt2"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(2, pointModel)

                x = Math.Round(pointTopRight2Circle.X, 3)
                y = Math.Round(pointTopRight2Circle.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = Math.Round(pointTopRight2Circle.X - pointBottomRight2Circle.X, 3)
                dy = Math.Round(pointTopRight2Circle.Y - pointBottomRight2Circle.Y, 3)
                dz = Math.Round(Height, 3)
                code = "rightPt2"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(1, pointModel)

                x = Math.Round(pointTopRight1Circle.X, 3)
                y = Math.Round(pointTopRight1Circle.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = Math.Round(pointTopRight1Circle.X - pointBottomRight1Circle.X, 3)
                dy = Math.Round(pointTopRight1Circle.Y - pointBottomRight1Circle.Y, 3)
                dz = Math.Round(Height, 3)
                code = "rightPt1"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(4, pointModel)
                _elementBridgePoint.ListPointModel = listModelPoint
            Else
                Dim insPtBottomRack1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleRN_v, WidthTop / 2)
                Dim insPtBottomRack2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPtBottomRack1, angleN_v, WidthBottom)
                Dim centerGrillage As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(insPtBottomRack1, insPtBottomRack2)
                Dim x As Double = Math.Round(centerRack.X, 3)
                Dim y As Double = Math.Round(centerRack.Y, 3)
                Dim z As Double = Math.Round(topElevationRack, 3)
                Dim dx As Double = Math.Round(centerRack.X - centerGrillage.X, 3)
                Dim dy As Double = Math.Round(centerRack.Y - centerGrillage.Y, 3)
                Dim dz As Double = Height
                Dim code As String = "center"
                _elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                _elementBridgePoint.CenterBottomPoint = New Vector3D(x, y, z - Height)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'делаем расчет 4 точек
                'находим расстояние от края насадки до края стойки слева
                Dim offsetTop As Double = WidthTop / 2
                Dim insPoint1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleN_v, offsetTop)
                Dim insPoint2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerRack2d, angleRN_v, offsetTop)
                Dim tempAxisLineRack As DwgLine = New DwgLine
                tempAxisLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(insPoint2, 0)
                tempAxisLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(insPoint1, 0)

                Dim leftPointTop1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint1, anglePillar, Diameter / 2)
                Dim rightPointTop1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint1, reverseAnglePillar, Diameter / 2)
                Dim tempBottomLineRack As DwgLine = New DwgLine
                tempBottomLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(leftPointTop1, 0)
                tempBottomLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(rightPointTop1, 0)

                Dim leftPointTop2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint2, anglePillar, Diameter / 2)
                Dim rightPointTop2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insPoint2, reverseAnglePillar, Diameter / 2)
                Dim tempTopLineRack As DwgLine = New DwgLine
                tempTopLineRack.StartPoint = New Topomatic.Cad.Foundation.Vector3D(leftPointTop2, 0)
                tempTopLineRack.EndPoint = New Topomatic.Cad.Foundation.Vector3D(rightPointTop2, 0)
                'делаем параллельный перенос оси стойки на половину диаметра
                Dim positionRackTopPointLeft1 As Vector2D = New Vector2D()
                Dim positionRackTopPointLeft2 As Vector2D = New Vector2D()

                Dim positionRackTopPointRight1 As Vector2D = New Vector2D()
                Dim positionRackTopPointRight2 As Vector2D = New Vector2D()

                Dim tempD As Double = Diameter / 2
                Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
                tempAxisLineRack.Offset(listEnt, tempD)
                If listEnt.Count > 0 Then
                    Dim newLine As DwgLine = listEnt(0)
                    positionRackTopPointRight2 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempBottomLineRack.StartPoint.Pos, tempBottomLineRack.EndPoint.Pos)
                    positionRackTopPointLeft2 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempTopLineRack.StartPoint.Pos, tempTopLineRack.EndPoint.Pos)
                End If
                listEnt = New List(Of DwgEntity)
                tempAxisLineRack.Offset(listEnt, -1 * tempD)
                If listEnt.Count > 0 Then
                    Dim newLine As DwgLine = listEnt(0)
                    positionRackTopPointRight1 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempBottomLineRack.StartPoint.Pos, tempBottomLineRack.EndPoint.Pos)
                    positionRackTopPointLeft1 = MathFunction.FuncFindLineIntersection(newLine.StartPoint.Pos, newLine.EndPoint.Pos, tempTopLineRack.StartPoint.Pos, tempTopLineRack.EndPoint.Pos)
                End If

                '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                'расчитываем точки по ростверку
                Dim rightPointBottom1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(positionRackTopPointLeft1, angleN_v, WidthBottom)
                Dim rightPointBottom2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(positionRackTopPointLeft2, angleN_v, WidthBottom)

                Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
                x = Math.Round(positionRackTopPointLeft1.X, 3)
                y = Math.Round(positionRackTopPointLeft1.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = 0
                dy = 0
                dz = Math.Round(Height, 3)
                code = "leftPt1"
                Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(3, pointModel)

                x = Math.Round(positionRackTopPointLeft2.X, 3)
                y = Math.Round(positionRackTopPointLeft2.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = 0
                dy = 0
                dz = Math.Round(Height, 3)
                code = "leftPt2"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(2, pointModel)

                x = Math.Round(positionRackTopPointRight2.X, 3)
                y = Math.Round(positionRackTopPointRight2.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = Math.Round(positionRackTopPointRight2.X - rightPointBottom2.X, 3)
                dy = Math.Round(positionRackTopPointRight2.Y - rightPointBottom2.Y, 3)
                dz = Math.Round(Height, 3)
                code = "rightPt2"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(1, pointModel)

                x = Math.Round(positionRackTopPointRight1.X, 3)
                y = Math.Round(positionRackTopPointRight1.Y, 3)
                z = Math.Round(topElevationRack, 3)
                dx = Math.Round(positionRackTopPointRight1.X - rightPointBottom1.X, 3)
                dy = Math.Round(positionRackTopPointRight1.Y - rightPointBottom1.Y, 3)
                dz = Math.Round(Height, 3)
                code = "rightPt1"
                pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                listModelPoint.Add(4, pointModel)
                _elementBridgePoint.ListPointModel = listModelPoint
            End If
        End If
        Return result
    End Function
    'рисование оси стойки
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineRack As DwgLine = Nothing
        Dim dataStructureRack As StructureElement = Nothing
        'ищем существующую ось стойки
        Dim listAxisRack As List(Of StructureElement) = Pillar.getElementPillar(StructureElement.typeObject.axisRack, dictionaryBridgeElements, NumberPillar, 0, Number, NumberSubPillars, Pillar.SidePillarElement.None)
        If listAxisRack.Count = 0 Then
            dataStructureRack = createAxisRackPillar(idBridge)
        ElseIf listAxisRack.Count = 1 Then
            dataStructureRack = listAxisRack.Item(0)
        Else
            dataStructureRack = StructureElement.isValidateDataStructure(listAxisRack)
        End If
        If IsNothing(dataStructureRack) Then Return Nothing
        axisLineRack = dataStructureRack.DWGEntity
        If IsNothing(axisLineRack) = True Then axisLineRack = New DwgLine
        If axisLineRack.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Стойка (ось)")
            styleAxisNozzle.setObjectStyle(axisLineRack)
        End If
        'ось стойки
        Dim lenghtAxisRack As Double = (_elementBridgePoint.CenterTopPoint - _elementBridgePoint.CenterBottomPoint).Length
        If lenghtAxisRack = 0 Then
            MsgBox("Ось стойки имеет нулевое значение.")
            Return dataStructureRack
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructureRack.KeyParameter = strJson
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineRack) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineRack)
        End If
        axisLineRack.StartPoint = _elementBridgePoint.CenterTopPoint
        axisLineRack.EndPoint = _elementBridgePoint.CenterBottomPoint
        dataStructureRack.DWGEntity = axisLineRack
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineRack, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureRack)
        Return dataStructureRack
    End Function

    'функция возвращает все точки стоек в виде словаря 
    Public Shared Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure, ByVal numberPillar As Integer, ByVal arrayRack As RackPillar(), ByRef dictProjectPointRack As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))) As Boolean
        If numberPillar < 1 Then Return False
        If IsNothing(arrayRack) = True Then Return False
        If IsArray(arrayRack) = False Then Return False
        dictProjectPointRack = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
        For i As Integer = 0 To arrayRack.Length - 1
            Dim userRack As RackPillar = arrayRack(i)
            If IsNothing(userRack) = True Then Continue For
            If userRack.NumberPillar < 1 Then Continue For
            Dim listElementRack As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
            'точки по верху подферменника
            Dim transformPointTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
            'точки по низу подферменника
            Dim transformPointBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
            Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userRack._elementBridgePoint.ListPointModel
            If ListPointModel.Count > 3 Then
                'записываем результат
                For j As Integer = 0 To ListPointModel.Count - 1
                    Dim pointStructure As PointStructure = ListPointModel.ElementAt(j).Value
                    Dim code As String = pointStructure.Code
                    Dim pointTop As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointStructure.X, pointStructure.Y, pointStructure.Z)
                    Dim pointBottom As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointStructure.X - pointStructure.dx, pointStructure.Y - pointStructure.dy, pointStructure.Z + pointStructure.dz)
                    Dim transformedBottomPoint As Cad.Foundation.Vector3D = matrixTransform.transformUserPoint(pointBottom)
                    If transformPointBottom.ContainsKey(code) = False Then
                        Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointBottom
                        projectPoint.projectPoint = transformedBottomPoint
                        transformPointBottom.Add(code, projectPoint)
                    End If
                    Dim transformedTopPoint As Cad.Foundation.Vector3D = matrixTransform.transformUserPoint(pointTop)
                    If transformPointTop.ContainsKey(code) = False Then
                        Dim projectPoint As ProjectionPoint = New ProjectionPoint()
                        projectPoint.originPoint = pointTop
                        projectPoint.projectPoint = transformedTopPoint
                        transformPointTop.Add(code, projectPoint)
                    End If
                Next j
            End If
            listElementRack.Add(transformPointTop)
            listElementRack.Add(transformPointBottom)
            dictProjectPointRack.Add(userRack.Number, listElementRack)
        Next i
        If dictProjectPointRack.Count > 1 Then
            dictProjectPointRack = dictProjectPointRack.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return True
    End Function

    'функция возвращает координаты точки по ее коду
    Public Function getPointByCode(ByVal code As String, ByRef point As Vector3D) As Boolean
        getPointByCode = False
        If IsNothing(code) = False Then
            If code.Trim.Length > 0 Then
                Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
                If ListPointModel.Count > 0 Then
                    For i As Integer = 0 To ListPointModel.Count - 1
                        Dim ptStructure As PointStructure = ListPointModel.ElementAt(i).Value
                        If ptStructure.Code Like code Then
                            point = New Vector3D(ptStructure.X, ptStructure.Y, ptStructure.Z)
                            Return True
                        End If
                    Next i
                End If
            End If
        End If
    End Function

    'функция делает сортировку стоек
    Public Shared Function sortListRack(ByVal listRacks As List(Of StructureElement), ByVal numberPillar As Integer) As Dictionary(Of Integer, StructureElement)
        Dim result As New Dictionary(Of Integer, StructureElement)
        If listRacks.Count = 0 Then Return result
        For i As Integer = 0 To listRacks.Count - 1
            Dim dataRack As StructureElement = listRacks.Item(i)
            Dim userRack As RackPillar = dataRack.getRackPillar
            If userRack.NumberPillar = numberPillar Then
                result.Add(userRack.Number, dataRack)
            End If
        Next i
        If result.Count > 1 Then
            result = result.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return result
    End Function
End Class
