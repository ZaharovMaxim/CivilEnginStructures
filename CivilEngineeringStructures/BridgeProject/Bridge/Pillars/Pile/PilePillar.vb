Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports CivilEnginStructures.NozzlePillar
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Algorithm
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class PilePillar
    Public Enum TypePile
        <Description("Призматическая")> Prismatic = 0
        <Description("Буровая")> Drilling = 1
        <Description("Не определено")> None = 2
    End Enum
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _numberSubPillars As Integer              ' Номер подопоры
    Private _numberRow As Integer                     ' Номер ряда
    Private _numberColumn As Integer                  ' Номер столбца
    Private _diameter As Double                       ' Диаметр сваи (для круглых)
    Private _width As Double                          ' Ширина стороны сваи (для прямоугольных)
    Private _height As Double                         ' Высота/длина сваи
    Private _topSeal As Double                        ' Высота заделки в ростверк или насадку
    Private _offsetX As Double                        ' Смещение сваи вдоль ростверка
    Private _offsetY As Double                        ' Смещение сваи поперек ростверка
    Private _offsetBottomX As Double                      ' Смещение низа сваи в ряде относительно центра
    Private _offsetBottomY As Double                   ' Смещение низа сваи в столбце
    Private _angleX As Double                         ' Угол наклона сваи по оси X
    Private _angleY As Double                         ' Угол наклона сваи по оси Y
    Private _rotation As Double                       ' Угол поворота
    Private _topElevation As Double                   ' Отметка верха
    Private _bottomElevation As Double                ' Отметка низа
    Private _expand As Boolean                        ' Наличие уширения
    Private _heightExpand As Double                   ' Высота уширения
    Private _widthExpand As Double                    ' Диаметр/ширина уширения
    Private _heightDownExpand As Double               ' Высота от низа сваи до уширения
    Private _degExpand As Double                      ' Угол уширения
    Private _pileInRack As Boolean                    ' Стойка в свае
    Private _type As TypePile                         ' Тип: призматическая/буровая
    Private _model As String                          ' Имя модели
    Public _elementBridgePoint As PointsCollections
    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _numberSubPillars = 0
        _numberRow = 0
        _numberColumn = 0
        _diameter = 0.0
        _width = 0.0
        _height = 0.0
        _topSeal = 0.0
        _offsetX = 0.0
        _offsetY = 0.0
        _offsetBottomX = 0.0
        _offsetBottomY = 0.0
        _angleX = 90.0
        _angleY = 90.0
        _rotation = 0.0
        _topElevation = 0.0
        _bottomElevation = 0.0
        _expand = False
        _heightExpand = 0.1
        _widthExpand = 1.0
        _heightDownExpand = 0.2
        _degExpand = 10.0
        _type = TypePile.None
        _pileInRack = False
        _model = String.Empty
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
    <Description("Номер ряда")>
    <Category("Свойства")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberRow = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Номер столбца")>
    <Category("Свойства")>
    <DisplayName("Номер столбца")>
    <[ReadOnly](True)>
    Public Property NumberColumn() As Integer
        Get
            Return _numberColumn
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberColumn = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Диаметр сваи, м")>
    <Category("Свойства")>
    <DisplayName("Диаметр")>
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
    <Description("Ширина сваи, м")>
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
    <Description("Высота сваи, м")>
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
    <Description("Высота заделки сваи в ростверк или насадку, м")>
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
    <Description("Смещение верха сваи вдоль ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение верха X")>
    Public Property OffsetX() As Double
        Get
            Return _offsetX
        End Get
        Set(value As Double)
            _offsetX = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение верха сваи поперек ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение верха Y")>
    Public Property OffsetY() As Double
        Get
            Return _offsetY
        End Get
        Set(value As Double)
            _offsetY = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение низа сваи вдоль ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение низа X")>
    Public Property OffsetBottomX() As Double
        Get
            Return _offsetBottomX
        End Get
        Set(value As Double)
            _offsetBottomX = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение низа сваи поперекь ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение низа Y")>
    Public Property OffsetBottomY() As Double
        Get
            Return _offsetBottomY
        End Get
        Set(value As Double)
            _offsetBottomY = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Угол наклона сваи вдоль ростверка или насадки, град")>
    <Category("Свойства")>
    <DisplayName("Угол наклона X")>
    Public Property AngleX() As Double
        Get
            Return _angleX
        End Get
        Set(value As Double)
            _angleX = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Угол наклона сваи поперек ростверка или насадки, град")>
    <Category("Свойства")>
    <DisplayName("Угол наклона Y")>
    Public Property AngleY() As Double
        Get
            Return _angleY
        End Get
        Set(value As Double)
            _angleY = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Угол поворота сваи, град")>
    <Category("Свойства")>
    <DisplayName("Угол поворота")>
    Public Property Rotation() As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха сваи, м")>
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
    <Description("Отметка низа сваи, м")>
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
    <Description("Наличие уширения")>
    <Category("Свойства")>
    <DisplayName("Наличие уширения")>
    Public Property Expand() As Boolean
        Get
            Return _expand
        End Get
        Set(value As Boolean)
            _expand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота уширения сваи, м")>
    <Category("Свойства")>
    <DisplayName("Высота уширения")>
    Public Property HeightExpand() As Double
        Get
            Return _heightExpand
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightExpand = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Величина уширения сваи, м")>
    <Category("Свойства")>
    <DisplayName("Величина уширения")>
    Public Property WidthExpand() As Double
        Get
            Return _widthExpand
        End Get
        Set(value As Double)
            If value >= 0 Then
                _widthExpand = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Расстояние от низа сваи до уширения, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние от низа сваи")>
    Public Property HeightDownExpand() As Double
        Get
            Return _heightDownExpand
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightDownExpand = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Угол ушмрения, град")>
    <Category("Свойства")>
    <DisplayName("Угол ушмрения сваи")>
    Public Property DegExpand() As Double
        Get
            Return _degExpand
        End Get
        Set(value As Double)
            _degExpand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Стойка вставлена в стойку")>
    <Category("Свойства")>
    <DisplayName("Стойка в свае")>
    Public Property PileInRack() As Boolean
        Get
            Return _pileInRack
        End Get
        Set(value As Boolean)
            _pileInRack = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Тип сваи")>
    <Category("Свойства")>
    <DisplayName("Тип сваи")>
    Public Property Type() As TypePile
        Get
            Return _type
        End Get
        Set(value As TypePile)
            _type = value
        End Set
    End Property

    Public Property NameModel() As String
        Get
            Return _model
        End Get
        Set(value As String)
            _model = value
        End Set
    End Property
    'функции для работы со сваей
    Public Shared Function createAxisPile(ByVal idBridge As String) As StructureElement
        Dim elementPile As StructureElement = New StructureElement()
        elementPile.Label = "Мосты и путепроводы"
        elementPile.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPile.ClassObject = StructureElement.classStructure.PilePillar
        elementPile.Name = StructureElement.typeObject.axisPile
        elementPile.Description = "Свая (ось)"
        elementPile.KeyParameter = ""
        elementPile.IdElement = Guid.NewGuid.ToString
        elementPile.IdStructure = idBridge
        elementPile.Note = ""
        elementPile.DWGEntity = New DwgLine()
        Return elementPile
    End Function
    'ищет сваи и возвращает результат в виде словаря
    Public Shared Function getPilePillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        Dim dictPiles As New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPile) = True Then
            Dim listAxisPile = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPile)
            If IsNothing(listAxisPile) = False Then
                If listAxisPile.Count > 0 Then
                    For k As Integer = 0 To listAxisPile.Count - 1
                        Dim tempData As StructureElement = listAxisPile.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisPile As PilePillar = tempData.getPilePillar
                            If IsNothing(userAxisPile) = False Then
                                If numberPillar = userAxisPile.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisPile.NumberSubPillars
                                    If numberSubPillar = userAxisPile.NumberSubPillars Then
                                        Dim numRow As Integer = userAxisPile.NumberRow
                                        Dim numColl As Integer = userAxisPile.NumberColumn
                                        If dictPiles.ContainsKey(numRow) = False Then
                                            Dim dictColl As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
                                            dictColl.Add(numColl, tempData)
                                            dictPiles.Add(numRow, dictColl)
                                        Else
                                            Dim dictColl As Dictionary(Of Integer, StructureElement) = dictPiles.Item(numRow)
                                            If dictColl.ContainsKey(numColl) = False Then
                                                dictColl.Add(numColl, listAxisPile.Item(k))
                                                dictPiles.Item(numRow) = dictColl
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        If dictPiles.Count > 0 Then
            For i As Integer = 0 To dictPiles.Count - 1
                Dim dictColl As Dictionary(Of Integer, StructureElement) = dictPiles.ElementAt(i).Value
                If dictColl.Count > 1 Then
                    dictColl = dictColl.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
                End If
            Next i
            If dictPiles.Count > 1 Then
                dictPiles = dictPiles.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            End If
        End If
        Return dictPiles
    End Function
    'ищет сваи и возвращает результат в виде массива
    Public Shared Function getPilePillarToArray(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As PilePillar()
        Dim result As PilePillar() = {}
        Dim count As Integer = 0
        Dim dictPiles As New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPile) = True Then
            Dim listAxisPile = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPile)
            If IsNothing(listAxisPile) = False Then
                If listAxisPile.Count > 0 Then
                    For k As Integer = 0 To listAxisPile.Count - 1
                        Dim tempData As StructureElement = listAxisPile.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisPile As PilePillar = tempData.getPilePillar
                            If IsNothing(userAxisPile) = False Then
                                If numberPillar = userAxisPile.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisPile.NumberSubPillars
                                    If numberSubPillar = userAxisPile.NumberSubPillars Then
                                        ReDim Preserve result(count)
                                        result(count) = userAxisPile
                                        count += 1
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return result
    End Function
    Public Shared Function readPropertiesPile(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Piles As DataGridView, Optional ByVal InsertPileInRack As Boolean = False, Optional ByVal PileExpand As Boolean = False, Optional typePile As TypePile = TypePile.Prismatic, Optional nameModel As String = "") As PilePillar()
        Dim result As PilePillar() = {}
        Dim countArrayPile As Integer = 0
        If DGV_Piles.ColumnCount > 1 Then
            For i As Integer = 1 To DGV_Piles.ColumnCount - 1
                Dim tempUserPile As PilePillar = New PilePillar
                tempUserPile.NumberPillar = numbPillar
                tempUserPile.NumberSubPillars = numbSubPillar
                tempUserPile.Type = typePile
                tempUserPile.NameModel = nameModel
                If DGV_Piles.RowCount > 1 Then
                    For j As Integer = 0 To DGV_Piles.RowCount - 1
                        Dim tag As String = DGV_Piles.Rows(j).Tag
                        Dim value As String = DGV_Piles.Rows(j).Cells(i).Value
                        If IsNothing(tag) = False And IsNothing(value) = False Then
                            If tag Like "numberRow" Then
                                Dim numberRowPile As Integer = CInt(value)
                                If numberRowPile > 0 Then
                                    tempUserPile.NumberRow = numberRowPile
                                Else
                                    Continue For
                                End If
                            ElseIf tag Like "numberColl" Then
                                Dim numberCollPile As Integer = CInt(value)
                                If numberCollPile > 0 Then
                                    tempUserPile.NumberColumn = numberCollPile
                                Else
                                    Continue For
                                End If
                            ElseIf tag Like "bridge_piles_length" Then
                                If IsNumeric(value) = True And value > 0 Then
                                    tempUserPile.Height = Val(value)
                                End If
                            ElseIf tag Like "bridge_piles_diam" Then
                                If IsNumeric(value) = True And value > 0 Then
                                    tempUserPile.Diameter = Val(value)
                                End If
                            ElseIf tag Like "bridge_piles_height" Then
                                If IsNumeric(value) = True And value > 0 Then
                                    tempUserPile.TopSeal = Val(value)
                                End If
                            ElseIf tag Like "bridge_piles_direction" Then
                                If IsNumeric(value) = True And value > 0 Then
                                    tempUserPile.Width = Math.Round(Val(value) / 1000, 3)
                                End If
                            ElseIf tag Like "bridge_piles_ushir_down" Then
                                If IsNumeric(value) = True Then
                                    tempUserPile.WidthExpand = Math.Round(Val(value), 3)
                                End If
                            ElseIf tag Like "bridge_piles_ushir_diam" Then
                                If IsNumeric(value) = True Then
                                    tempUserPile.HeightExpand = Math.Round(Val(value), 3)
                                End If
                            ElseIf tag Like "bridge_piles_ushir_downtoushir" Then
                                If IsNumeric(value) = True Then
                                    tempUserPile.HeightDownExpand = Math.Round(Val(value), 3)
                                End If
                            ElseIf tag Like "bridge_piles_ushir_grade" Then
                                If IsNumeric(value) = True Then
                                    tempUserPile.DegExpand = Math.Round(Val(value), 3)
                                End If
                            ElseIf tag Like "bridge_piles_angleX" Then
                                tempUserPile.AngleX = Math.Round(Val(value), 2)
                            ElseIf tag Like "bridge_piles_angleY" Then
                                tempUserPile.AngleY = Math.Round(Val(value), 2)
                            End If
                        End If
                    Next j
                End If
                If InsertPileInRack = True Then
                    tempUserPile.PileInRack = True
                Else
                    tempUserPile.PileInRack = False
                End If
                If PileExpand = True Then
                    tempUserPile.Expand = True
                Else
                    tempUserPile.Expand = False
                End If
                ReDim Preserve result(countArrayPile)
                result(countArrayPile) = tempUserPile
                countArrayPile += 1
            Next i
        End If
        Return result
    End Function

    Public Function writePropertiesPile(ByRef DGV_Piles As DataGridView, ByVal indexPile As Integer) As Boolean
        If IsNothing(DGV_Piles) = True Then Return False
        If DGV_Piles.RowCount > 1 Then
            If DGV_Piles.Columns.Count >= indexPile Then
                For j As Integer = 0 To DGV_Piles.RowCount - 1
                    Dim tag As String = DGV_Piles.Rows(j).Tag
                    If IsNothing(tag) = False Then
                        If tag Like "bridge_piles_length" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = Height
                        ElseIf tag Like "bridge_piles_diam" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = Diameter
                        ElseIf tag Like "bridge_piles_height" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = TopSeal
                        ElseIf tag Like "bridge_piles_direction" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = Width * 1000
                        ElseIf tag Like "bridge_piles_angleX" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = AngleX
                        ElseIf tag Like "bridge_piles_angleY" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = AngleY
                        ElseIf tag Like "bridge_piles_ushir_down" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = WidthExpand
                        ElseIf tag Like "bridge_piles_ushir_diam" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = HeightExpand
                        ElseIf tag Like "bridge_piles_ushir_downtoushir" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = HeightDownExpand
                        ElseIf tag Like "bridge_piles_ushir_grade" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = DegExpand
                        ElseIf tag Like "numberRow" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = NumberRow
                        ElseIf tag Like "numberColl" Then
                            DGV_Piles.Rows(j).Cells(indexPile).Value = NumberColumn
                        End If
                    End If
                Next j
            End If
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Pile As DataGridView) As Boolean
        If IsNothing(DGV_Pile) = True Then Return False
        If DGV_Pile.RowCount > 1 Then
            Dim numberTopElevation As Integer = -1
            Dim numberBottomElevation As Integer = -1
            Dim numberPileRow As Integer = -1
            Dim numberPileColl As Integer = -1
            For i As Integer = 0 To DGV_Pile.RowCount - 1
                Dim oldTag As String = DGV_Pile.Rows(i).Tag
                If oldTag Like "calc-TopElevation" Then
                    numberTopElevation = i
                ElseIf oldTag Like "calc-BottomElevation" Then
                    numberBottomElevation = i
                ElseIf oldTag Like "numberColl" Then
                    numberPileColl = i
                ElseIf oldTag Like "numberRow" Then
                    numberPileRow = i
                End If
            Next i
            If numberTopElevation = -1 Then
                numberTopElevation = DGV_Pile.RowCount - 1
                DGV_Pile.Rows.Insert(numberTopElevation)
                DGV_Pile.Rows(numberTopElevation).DefaultCellStyle.ForeColor = Color.Red
                DGV_Pile.Rows(numberTopElevation).Tag = "calc-TopElevation"
                DGV_Pile.Rows(numberTopElevation).Cells(0).Value = "Отметка верха, м"
            End If
            If numberBottomElevation = -1 Then
                numberBottomElevation = DGV_Pile.RowCount - 1
                DGV_Pile.Rows.Insert(numberTopElevation)
                DGV_Pile.Rows(numberTopElevation).DefaultCellStyle.ForeColor = Color.Red
                DGV_Pile.Rows(numberTopElevation).Tag = "calc-BottomElevation"
                DGV_Pile.Rows(numberTopElevation).Cells(0).Value = "Отметка низа, м"
            End If
            For i As Integer = 1 To DGV_Pile.ColumnCount - 1
                Dim tempNumberRow As Integer = DGV_Pile.Rows(numberPileRow).Cells(i).Value
                Dim tempNumberColl As Integer = DGV_Pile.Rows(numberPileColl).Cells(i).Value
                If tempNumberRow = NumberRow Then
                    If tempNumberColl = NumberColumn Then
                        DGV_Pile.Rows(numberTopElevation).Cells(i).Value = TopElevation
                        DGV_Pile.Rows(numberBottomElevation).Cells(i).Value = BottomElevation
                        Exit For
                    End If
                End If
            Next i
        End If
        Return True
    End Function
    'расчет положения свай
    Public Shared Function calculatePile(ByVal userNozzle As NozzlePillar, ByVal userGrillage As GrillagePillar, ByVal arrayPile As PilePillar(), Optional dictRacks As Dictionary(Of Integer, RackPillar) = Nothing, Optional pileInRack As Boolean = False) As Dictionary(Of Integer, Dictionary(Of Integer, PilePillar))
        Dim result As New Dictionary(Of Integer, Dictionary(Of Integer, PilePillar))
        If pileInRack = False Then
            Dim leftPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, 1)
            Dim leftPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, 1)
            Dim rightPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, 1)
            Dim rightPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(-1, -1, 1)
            Dim listPointBridge As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
            Dim pileColumnFieldDiagram As String = ""
            Dim pileRowsFieldDiagram As String = ""
            Dim bottomElevationElement As Double = 0

            If IsNothing(userGrillage) = False Then
                If userGrillage._elementBridgePoint.ListPointModel.Count > 3 Then
                    pileColumnFieldDiagram = userGrillage.PileColumnFieldDiagram
                    pileRowsFieldDiagram = userGrillage.PileRowsFieldDiagram
                    listPointBridge = userGrillage._elementBridgePoint.ListPointModel

                    bottomElevationElement = userGrillage.BottomElevation
                End If
            End If
            If listPointBridge.Count = 0 Then
                If IsNothing(userNozzle) = False Then
                    pileColumnFieldDiagram = userNozzle.PileColumnDiagram
                    pileRowsFieldDiagram = userNozzle.PileRowsDiagram
                    If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
                        listPointBridge = userNozzle._elementBridgePoint.ListPointModel
                    End If
                    bottomElevationElement = userNozzle.BottomElevation
                End If
            End If
            If listPointBridge.Count > 3 Then
                For i As Integer = 0 To listPointBridge.Count - 1
                    Dim tempPoint As PointStructure = listPointBridge.ElementAt(i).Value
                    If tempPoint.Code.IndexOf("leftPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X + tempPoint.dx, tempPoint.Y + tempPoint.dy, tempPoint.Z + tempPoint.dz)
                        leftPoint1 = tempPt
                    ElseIf tempPoint.Code.IndexOf("leftPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X + tempPoint.dx, tempPoint.Y + tempPoint.dy, tempPoint.Z + tempPoint.dz)
                        leftPoint2 = tempPt
                    ElseIf tempPoint.Code.IndexOf("rightPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X + tempPoint.dx, tempPoint.Y + tempPoint.dy, tempPoint.Z + tempPoint.dz)
                        rightPoint1 = tempPt
                    ElseIf tempPoint.Code.IndexOf("rightPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X + tempPoint.dx, tempPoint.Y + tempPoint.dy, tempPoint.Z + tempPoint.dz)
                        rightPoint2 = tempPt
                    End If
                Next

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

                Dim middleLeftPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(leftPoint1, rightPoint1)
                Dim middleRightPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(leftPoint2, rightPoint2)
                Dim axisLongLine As DwgLine = New DwgLine
                axisLongLine.StartPoint = New Cad.Foundation.Vector3D(middleLeftPoint, bottomElevationElement)
                axisLongLine.EndPoint = New Cad.Foundation.Vector3D(middleRightPoint, bottomElevationElement)
                Dim rotationPileLat As Double = axisLongLine.Rotation
                Dim axisShortLine As DwgLine = New DwgLine()
                axisShortLine.StartPoint = leftPoint1
                axisShortLine.EndPoint = rightPoint1
                Dim rotationPileLon As Double = axisShortLine.Rotation
                'делаем разбор строк
                Dim count As Integer = 0
                Dim arrayLatPile As Double() = {} '(0-
                Dim arrayLonPile As Double() = {} '(0-
                'делаем разбор строки 1
                If pileColumnFieldDiagram.Trim.Length > 0 And pileRowsFieldDiagram.Trim.Length Then
                    Dim arrayStr As String() = pileRowsFieldDiagram.Split("+")
                    Dim newlenghtGrillage As Double = 0
                    If IsArray(arrayStr) = True Then
                        For k1 As Integer = 0 To arrayStr.Length - 1
                            If k1 = 0 And IsNumeric(arrayStr(k1)) Then
                                ReDim Preserve arrayLatPile(count)
                                Dim sh As Double = Val(arrayStr(k1))
                                arrayLatPile(count) = sh
                                count += 1
                                newlenghtGrillage = newlenghtGrillage + Val(arrayStr(k1))
                            ElseIf k1 = arrayStr.Length - 1 And IsNumeric(arrayStr(k1)) Then
                                ReDim Preserve arrayLatPile(count)
                                Dim sh As Double = Val(arrayStr(k1))
                                arrayLatPile(count) = sh
                                count += 1
                                newlenghtGrillage = newlenghtGrillage + Val(arrayStr(k1))
                            Else
                                Dim str As String = arrayStr(k1)
                                Dim arrayD As String() = str.Split("*")
                                If arrayD.Length = 2 Then
                                    Dim countPileTemp As Integer = Val(arrayD(0))
                                    Dim shagPileTemp As Double = Val(arrayD(1))
                                    If countPileTemp > 0 And shagPileTemp > 0 Then
                                        For k2 As Integer = 0 To countPileTemp - 1
                                            ReDim Preserve arrayLatPile(count)
                                            arrayLatPile(count) = shagPileTemp
                                            count += 1
                                            newlenghtGrillage = newlenghtGrillage + shagPileTemp
                                        Next k2
                                    End If
                                End If
                            End If
                        Next k1
                    End If
                    Dim heightGrillage As Double = 0
                    count = 0
                    arrayStr = pileColumnFieldDiagram.Split("+")
                    If IsArray(arrayStr) = True Then
                        For k1 As Integer = 0 To arrayStr.Length - 1
                            If k1 = 0 And IsNumeric(arrayStr(k1)) Then
                                ReDim Preserve arrayLonPile(count)
                                arrayLonPile(count) = Val(arrayStr(k1))
                                count += 1
                                heightGrillage = heightGrillage + Val(arrayStr(k1))
                            ElseIf k1 = arrayStr.Length - 1 And IsNumeric(arrayStr(k1)) Then
                                ReDim Preserve arrayLonPile(count)
                                arrayLonPile(count) = Val(arrayStr(k1))
                                count += 1
                                heightGrillage = heightGrillage + Val(arrayStr(k1))
                            Else
                                Dim str As String = arrayStr(k1)
                                Dim arrayD As String() = str.Split("*")
                                If arrayD.Length = 2 Then
                                    Dim countPileTemp As Integer = Val(arrayD(0))
                                    Dim shagPileTemp As Double = Val(arrayD(1))
                                    If countPileTemp > 0 And shagPileTemp > 0 Then
                                        For k2 As Integer = 0 To countPileTemp - 1
                                            ReDim Preserve arrayLonPile(count)
                                            arrayLonPile(count) = shagPileTemp
                                            count += 1
                                            heightGrillage = heightGrillage + shagPileTemp
                                        Next k2
                                    End If
                                End If
                            End If
                        Next k1
                    End If
                End If
                Dim globalLonDist As Double = 0
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'расстановка сваи
                If arrayPile.Length > 0 Then
                    For i As Integer = 0 To arrayLonPile.Length - 1
                        'делаем мсещение сваи вдоль оси Y
                        globalLonDist = globalLonDist + arrayLonPile(i) / 1000
                        'формируем временную ось для расстановки свай в ряду
                        Dim tempLeftPoint As Cad.Foundation.Vector3D = MathFunction.FuncCalcPointInLine(leftPoint1, rightPoint1, globalLonDist)
                        Dim tempRightPoint As Cad.Foundation.Vector3D = MathFunction.FuncCalcPointInLine(leftPoint2, rightPoint2, globalLonDist)
                        Dim globalLatDist As Double = 0
                        For j As Integer = 0 To arrayLatPile.Length - 1
                            Dim numberRow As Integer = i + 1
                            Dim numberColl As Integer = j + 1
                            Dim userPile As PilePillar = Nothing
                            For k As Integer = 0 To arrayPile.Length - 1
                                Dim tempUserPile As PilePillar = arrayPile(k)
                                If tempUserPile.NumberColumn = numberColl And tempUserPile.NumberRow = numberRow Then
                                    userPile = tempUserPile
                                    Exit For
                                End If
                            Next k
                            If IsNothing(userPile) = False Then
                                globalLatDist = globalLatDist + arrayLatPile(j) / 1000
                                Dim posPileColumn As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftPoint1.Pos, rotationPileLon, globalLonDist) 'смещение по столбцу (выходим на нужный ряд
                                Dim insertTopPointPile As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(posPileColumn, rotationPileLat, globalLatDist) 'смещение вдоль ростверка (выходим на нужный столбец)
                                Dim elevPile As Double = MathFunction.FuncCalcElevationByLine(tempLeftPoint, tempRightPoint, insertTopPointPile) 'высота верха сваи без учета заглубления
                                userPile.BottomElevation = elevPile - userPile.Height
                                userPile.TopElevation = Math.Round(elevPile + userPile.TopSeal, 3) 'высота сваи в уровне ростверка
                                userPile.OffsetX = Math.Round(globalLatDist, 3)
                                userPile.OffsetY = Math.Round(globalLonDist, 3)
                                Dim isertBottomPointPileX As Vector2D = insertTopPointPile
                                Dim isertBottomPointPileY As Vector2D = insertTopPointPile
                                'вычисляем смещение
                                Dim angleX As Double = userPile.AngleX
                                Dim lenghtPile As Double = userPile.Height
                                Dim offsetX As Double = 0
                                If angleX <> 90 Then
                                    Dim a As Double = MathFunction.FuncConvertDegtoRad(90 - angleX)
                                    Dim c As Double = lenghtPile / Math.Cos(a)
                                    offsetX = c * Math.Sin(a)
                                End If
                                Dim angleY As Double = userPile.AngleY
                                Dim offsetY As Double = 0
                                If angleY <> 90 Then
                                    Dim a As Double = MathFunction.FuncConvertDegtoRad(90 - angleY)
                                    Dim c As Double = lenghtPile / Math.Cos(a)
                                    offsetY = c * Math.Sin(a)
                                End If
                                'смещение право-лево (вдоль ростверка)
                                If offsetX <> 0 Then
                                    Dim rotLat As Double = rotationPileLat
                                    Dim l1 As Double = (middleLeftPoint - insertTopPointPile).Length
                                    Dim l2 As Double = (middleRightPoint - insertTopPointPile).Length
                                    If l1 < l2 Then
                                        angleX *= -1
                                        rotLat -= Math.PI
                                        If rotLat < 0 Then
                                            rotLat += Math.PI * 2
                                        End If
                                    End If
                                    isertBottomPointPileX = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insertTopPointPile, rotLat, Math.Abs(offsetX))
                                    userPile.AngleX = Math.Round(angleX, 2)
                                    Dim heightPile As Double = userPile.Height
                                    Dim x As Double = Math.Round(insertTopPointPile.X, 3)
                                    Dim y As Double = Math.Round(insertTopPointPile.Y, 3)
                                    Dim z As Double = Math.Round(userPile.TopElevation, 3)
                                    Dim code As String = "center"
                                    Dim dx As Double = Math.Round(insertTopPointPile.X - isertBottomPointPileX.X, 3)
                                    Dim dy As Double = Math.Round(insertTopPointPile.Y - isertBottomPointPileX.Y, 3)
                                    Dim dz As Double = Math.Round(heightPile, 3)
                                    Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                                    userPile._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.EndAxisPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.CenterBottomPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile.Rotation = Math.Round(rotationPileLat, 6)
                                    userPile.OffsetBottomX = Math.Round(offsetX, 3)
                                End If
                                If offsetY <> 0 Then
                                    Dim rotLon As Double = rotationPileLon
                                    Dim n As Integer = MathFunction.funcLeftOrRightPointToLinearObject(axisLongLine, insertTopPointPile)
                                    If n = 1 Then
                                        angleY *= -1
                                        rotLon -= Math.PI
                                        If rotLon < 0 Then
                                            rotLon += Math.PI * 2
                                        End If
                                    End If
                                    isertBottomPointPileY = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insertTopPointPile, rotLon, Math.Abs(offsetY))
                                    userPile.AngleY = Math.Round(angleY, 2)
                                    Dim heightPile As Double = userPile.Height
                                    Dim x As Double = Math.Round(insertTopPointPile.X, 3)
                                    Dim y As Double = Math.Round(insertTopPointPile.Y, 3)
                                    Dim z As Double = Math.Round(userPile.TopElevation, 3)
                                    Dim code As String = "center"
                                    Dim dx As Double = Math.Round(insertTopPointPile.X - isertBottomPointPileY.X, 3)
                                    Dim dy As Double = Math.Round(insertTopPointPile.Y - isertBottomPointPileY.Y, 3)
                                    Dim dz As Double = Math.Round(heightPile, 3)
                                    Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                                    userPile._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.EndAxisPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.CenterBottomPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile.Rotation = Math.Round(rotationPileLat, 6)
                                    userPile.OffsetBottomY = Math.Round(offsetY, 3)
                                End If
                                'наклон в обе стороны
                                If angleX <> 90 And angleY <> 90 Then
                                    Dim midPoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(isertBottomPointPileX, isertBottomPointPileY)
                                    Dim lenPlane As Double = Math.Sqrt((offsetX ^ 2 + offsetY ^ 2))
                                    Dim newLinePl As DwgLine = New DwgLine
                                    newLinePl.StartPoint = New Cad.Foundation.Vector3D(insertTopPointPile, 0)
                                    newLinePl.EndPoint = New Cad.Foundation.Vector3D(midPoint, 0)
                                    isertBottomPointPileX = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insertTopPointPile, newLinePl.Rotation, lenPlane)
                                    'elevPile = elevPile + tempUserPile.topSeal
                                    Dim heightPile As Double = userPile.Height '+ tempUserPile.topSeal
                                    Dim len1 As Double = (leftPoint1.Pos - isertBottomPointPileX).Length
                                    Dim len2 As Double = (leftPoint2.Pos - isertBottomPointPileX).Length
                                    Dim len3 As Double = (rightPoint1.Pos - isertBottomPointPileX).Length
                                    Dim len4 As Double = (rightPoint2.Pos - isertBottomPointPileX).Length
                                    If len1 < len2 And len1 < len3 And len1 < len4 Then
                                        userPile.AngleX = Math.Abs(angleX)
                                        userPile.AngleY = -1 * Math.Abs(angleY)
                                    ElseIf len2 < len1 And len2 < len3 And len2 < len4 Then
                                        userPile.AngleX = -1 * Math.Abs(angleX)
                                        userPile.AngleY = -1 * Math.Abs(angleY)
                                    ElseIf len3 < len1 And len3 < len2 And len3 < len4 Then
                                        userPile.AngleX = -1 * Math.Abs(angleX)
                                        userPile.AngleY = Math.Abs(angleY)
                                    ElseIf len4 < len1 And len4 < len2 And len4 < len3 Then
                                        userPile.AngleX = Math.Abs(angleX)
                                        userPile.AngleY = Math.Abs(angleY)
                                    End If

                                    Dim x As Double = Math.Round(insertTopPointPile.X, 3)
                                    Dim y As Double = Math.Round(insertTopPointPile.Y, 3)
                                    Dim z As Double = Math.Round(userPile.TopElevation, 3)
                                    Dim code As String = "center"
                                    Dim dx As Double = Math.Round(insertTopPointPile.X - isertBottomPointPileX.X, 3)
                                    Dim dy As Double = Math.Round(insertTopPointPile.Y - isertBottomPointPileX.Y, 3)
                                    Dim dz As Double = Math.Round(heightPile, 3)
                                    Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                                    userPile._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.EndAxisPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.CenterBottomPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile.Rotation = Math.Round(rotationPileLat, 6)
                                End If
                                If angleX = 90 And angleY = 90 Then
                                    'elevPile = elevPile + tempUserPile.topSeal
                                    Dim heightPile As Double = userPile.Height '+ tempUserPile.topSeal
                                    Dim x As Double = Math.Round(insertTopPointPile.X, 3)
                                    Dim y As Double = Math.Round(insertTopPointPile.Y, 3)
                                    Dim z As Double = Math.Round(userPile.TopElevation, 3)
                                    Dim code As String = "center"
                                    Dim dx As Double = 0
                                    Dim dy As Double = 0
                                    Dim dz As Double = Math.Round(userPile.Height, 3)
                                    Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                                    userPile._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.EndAxisPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                                    userPile._elementBridgePoint.CenterBottomPoint = New Vector3D(isertBottomPointPileX.X, isertBottomPointPileX.Y, userPile.BottomElevation)
                                    userPile.Rotation = Math.Round(rotationPileLat, 6)
                                End If
                            End If
                            If result.ContainsKey(numberRow) = False Then
                                Dim dictColumn As Dictionary(Of Integer, PilePillar) = New Dictionary(Of Integer, PilePillar)
                                dictColumn.Add(numberColl, userPile)
                                result.Add(numberRow, dictColumn)
                            Else
                                Dim dictColumn As Dictionary(Of Integer, PilePillar) = result.Item(numberRow)
                                If dictColumn.ContainsKey(numberColl) = False Then
                                    dictColumn.Add(numberColl, userPile)
                                    result.Item(numberRow) = dictColumn
                                End If
                            End If
                        Next j
                    Next i
                End If
            End If
        Else
            'вставляем сваю в стойку
            If dictRacks.Count > 0 Then
                Dim dictColumn As Dictionary(Of Integer, PilePillar) = New Dictionary(Of Integer, PilePillar)
                For i As Integer = 0 To dictRacks.Count - 1
                    Dim numberRack As Integer = dictRacks.ElementAt(i).Key
                    Dim tempUserRack As RackPillar = dictRacks.ElementAt(i).Value
                    Dim insBottomRack As Cad.Foundation.Vector3D = tempUserRack._elementBridgePoint.CenterBottomPoint
                    Dim userPile As PilePillar = Nothing
                    For k As Integer = 0 To arrayPile.Length - 1
                        Dim tempUserPile As PilePillar = arrayPile(k)
                        If tempUserPile.NumberColumn = numberRack Then
                            userPile = tempUserPile
                            Exit For
                        End If
                    Next k
                    If IsNothing(userPile) = False Then
                        Dim insPointPile As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(insBottomRack.Pos, insBottomRack.Z + userPile.TopSeal)
                        Dim x As Double = Math.Round(insPointPile.X, 3)
                        Dim y As Double = Math.Round(insPointPile.Y, 3)
                        Dim z As Double = Math.Round(insPointPile.Z, 3)
                        Dim code As String = "center"
                        Dim dx As Double = 0
                        Dim dy As Double = 0
                        Dim dz As Double = Math.Round(userPile.Height, 3)
                        Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
                        userPile._elementBridgePoint.StartAxisPoint = New Vector3D(x, y, z)
                        userPile._elementBridgePoint.EndAxisPoint = New Vector3D(x, y, z - userPile.Height)
                        userPile._elementBridgePoint.CenterTopPoint = New Vector3D(x, y, z)
                        userPile._elementBridgePoint.CenterBottomPoint = New Vector3D(x, y, z - userPile.Height)
                        If dictColumn.ContainsKey(numberRack) = False Then
                            dictColumn.Add(numberRack, userPile)
                        End If
                    End If
                Next i
                result.Item(1) = dictColumn
            End If
        End If
        Return result
    End Function
    'рисование оси сваи
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLinePile As DwgLine = Nothing
        Dim dataStructurePile As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxisPile As List(Of StructureElement) = Pillar.getElementPillar(StructureElement.typeObject.axisPile, dictionaryBridgeElements, NumberPillar, NumberRow, NumberColumn, NumberSubPillars, Pillar.SidePillarElement.None)
        If listAxisPile.Count = 0 Then
            dataStructurePile = createAxisPile(idBridge)
        ElseIf listAxisPile.Count = 1 Then
            dataStructurePile = listAxisPile.Item(0)
        Else
            dataStructurePile = StructureElement.isValidateDataStructure(listAxisPile)
        End If
        If IsNothing(dataStructurePile) Then Return dataStructurePile
        axisLinePile = dataStructurePile.DWGEntity
        If IsNothing(axisLinePile) = True Then Return Nothing
        If axisLinePile.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Свая (ось)")
            styleAxisNozzle.setObjectStyle(axisLinePile)
        End If
        'ось сваи
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructurePile.KeyParameter = strJson
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLinePile) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLinePile)
        End If
        axisLinePile.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLinePile.EndPoint = _elementBridgePoint.EndAxisPoint
        dataStructurePile.DWGEntity = axisLinePile
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLinePile, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructurePile)
        Return dataStructurePile
    End Function
    'функция возвращает все точки свай в виде словаря 
    Public Shared Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure, ByVal arrayPile As PilePillar(), ByRef dictProjectPointPile As Dictionary(Of Integer, Dictionary(Of Integer, ProjectionPoint()))) As Boolean
        If IsNothing(arrayPile) = True Then Return False
        If IsArray(arrayPile) = False Then Return False
        dictProjectPointPile = New Dictionary(Of Integer, Dictionary(Of Integer, ProjectionPoint()))
        For i As Integer = 0 To arrayPile.Length - 1
            Dim userPile As PilePillar = arrayPile(i)
            If IsNothing(userPile) = True Then Continue For
            If userPile.NumberPillar < 1 Then Continue For
            Dim numberRow As Integer = userPile.NumberRow
            Dim numberColumn As Integer = userPile.NumberColumn
            Dim pointTop As Vector3D = userPile._elementBridgePoint.CenterTopPoint
            Dim pointBottom As Vector3D = userPile._elementBridgePoint.CenterBottomPoint

            Dim transformedTopPoint As Cad.Foundation.Vector3D = matrixTransform.transformUserPoint(pointTop)
            Dim projectTopPoint As ProjectionPoint = New ProjectionPoint()
            projectTopPoint.originPoint = pointTop
            projectTopPoint.projectPoint = transformedTopPoint

            Dim transformedBottomPoint As Cad.Foundation.Vector3D = matrixTransform.transformUserPoint(pointBottom)
            Dim projectBottomPoint As ProjectionPoint = New ProjectionPoint()
            projectBottomPoint.originPoint = pointBottom
            projectBottomPoint.projectPoint = transformedBottomPoint
            Dim arrayProjectPoint As ProjectionPoint() = {projectTopPoint, projectBottomPoint}

            If dictProjectPointPile.ContainsKey(numberRow) = False Then
                Dim dictColPile As Dictionary(Of Integer, ProjectionPoint()) = New Dictionary(Of Integer, ProjectionPoint())
                dictColPile.Add(numberColumn, arrayProjectPoint)
                dictProjectPointPile.Add(numberRow, dictColPile)
            Else
                Dim dictColPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectPointPile.Item(numberRow)
                dictColPile.Add(numberColumn, arrayProjectPoint)
                dictProjectPointPile.Item(numberRow) = dictColPile
            End If
        Next i
        If dictProjectPointPile.Count > 1 Then
            dictProjectPointPile = dictProjectPointPile.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        For i As Integer = 0 To dictProjectPointPile.Count - 1
            Dim dictColPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectPointPile.ElementAt(i).Value
            If dictColPile.Count > 1 Then
                dictColPile = dictColPile.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            End If
        Next i
        Return True
    End Function
    'функция ищет в массиве сваю
    Public Shared Function findPileToArray(ByVal numberRow As Integer, ByVal numberColumn As Integer, ByRef arrayPile As PilePillar()) As PilePillar
        Dim result As PilePillar = Nothing
        If IsArray(arrayPile) = True Then
            For i As Integer = 0 To arrayPile.Length - 1
                Dim userPile As PilePillar = arrayPile(i)
                If userPile.NumberRow = numberRow Then
                    If userPile.NumberColumn = numberColumn Then
                        result = userPile
                        Exit For
                    End If
                End If
            Next
        End If
        Return result
    End Function
    'функция рисует сваю в pictureBox
    Public Shared Function drawPilePictureBox(ByRef bmp As Bitmap, ByVal arrayProjectPoint As ProjectionPoint(), ByVal userPile As PilePillar, ByVal k As Double, ByVal offsetX As Single, ByVal offsetY As String, ByVal prjView As ProjectionPoint.projectView, Optional ByVal numberPile As Integer = 0, Optional drawTextOffset As Double = 0) As Boolean
        If IsNothing(bmp) = True Then Return False
        If IsNothing(userPile) = True Then Return False
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            If IsArray(arrayProjectPoint) = True Then
                If arrayProjectPoint.Length > 1 Then
                    'рисование черным цветом
                    Dim colorPen As System.Drawing.Color = System.Drawing.Color.Black
                    'точки по верху и низу
                    Dim centerTop As Vector3D = arrayProjectPoint(0).projectPoint
                    Dim centerBottom As Vector3D = arrayProjectPoint(1).projectPoint
                    'длина по верху
                    Dim startPoint As Vector2D = Nothing
                    Dim endPoint As Vector2D = Nothing
                    Dim widthPile As Double = userPile.Width
                    If widthPile = 0 Then widthPile = userPile.Diameter

                    'по верху сторона (не зависит от типа сваи)
                    If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                        startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                        endPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                    Else
                        startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                        endPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                    End If
                    Dim x1 As Double = startPoint.X * k + offsetX
                    Dim y1 As Double = startPoint.Y * k + offsetY
                    Dim x2 As Double = endPoint.X * k + offsetX
                    Dim y2 As Double = endPoint.Y * k + offsetY
                    Dim startLine As New System.Drawing.Point(x1, y1)
                    Dim endLine As New System.Drawing.Point(x2, y2)
                    Using pen As New System.Drawing.Pen(colorPen, 1)
                        g.DrawLine(pen, startLine, endLine)
                    End Using
                    If userPile.Type = TypePile.Prismatic Then
                        'длина слева
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'длина справа
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'левый наконечник
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X, centerBottom.Z + 0.5)
                        Else
                            startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y, centerBottom.Z + 0.5)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'правый наконечник
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X, centerBottom.Z + 0.5)
                        Else
                            startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y, centerBottom.Z + 0.5)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                    Else
                        'свая круглая рисуем уширение
                        Dim dh As Double = 0
                        If userPile.Expand = True Then
                            Dim deg As Double = (userPile.DegExpand * Math.PI) / 180.0
                            dh = userPile.WidthExpand * Math.Sin(deg) + userPile.HeightDownExpand
                        End If
                        'длина слева
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                        Else
                            startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'длина справа
                        If userPile.Expand = True Then
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                            Else
                                startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                            End If
                        Else
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            End If
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        '================================================================================
                        'рисуем уширение
                        If userPile.Expand = True Then
                            Dim l As Double = userPile.WidthExpand
                            Dim deg As Double = (userPile.DegExpand * Math.PI) / 180.0
                            dh = l * Math.Sin(deg) + userPile.HeightDownExpand
                            '1 длинное уширение вниз
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 возвращаем на диаметр
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 вниз по свае
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '============================================================================
                            'правое направление
                            '1 длинное уширение вниз
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 возвращаем на диаметр
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 вниз по свае
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                        End If
                        'перемычка по низу
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                    End If
                    '==================================================================================================
                    ' Подписываем сваи
                    If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                        x1 = centerTop.X * k + offsetX
                        y1 = centerBottom.Z * k + offsetY
                    Else
                        x1 = centerTop.Y * k + offsetX
                        y1 = centerBottom.Z * k + offsetY
                    End If
                    Using font As New System.Drawing.Font("Arial", 8)
                        Using Brush As New SolidBrush(Color.Blue)

                            If (numberPile Mod 2) = 0 Then
                                g.DrawString(numberPile, font, Brush, x1, y1 + 15)
                            Else
                                g.DrawString(numberPile, font, Brush, x1, y1 + 7)
                            End If
                        End Using
                    End Using
                    If drawTextOffset <> 0 Then
                        If numberPile = 1 Then
                            Using font As New System.Drawing.Font("Arial", 8)
                                Using brush As New SolidBrush(Color.Blue)
                                    g.DrawString(Math.Round(drawTextOffset * 1000, 0), font, brush, x1 - 60, y1 - userPile.Height * k)
                                End Using
                            End Using
                        Else
                            Using font As New System.Drawing.Font("Arial", 8)
                                Using brush As New SolidBrush(Color.Blue)
                                    g.DrawString(Math.Round(drawTextOffset * 1000, 0), font, brush, x1 + 30, y1 - userPile.Height * k)
                                End Using
                            End Using
                        End If
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Using
        Return True
    End Function

    Public Shared Function drawTopPilePictureBox(ByRef bmp As Bitmap, ByVal arrayProjectPoint As ProjectionPoint(), ByVal userPile As PilePillar, ByVal k As Double, ByVal offsetX As Single, ByVal offsetY As String, ByVal prjView As ProjectionPoint.projectView, Optional ByVal numberPile As Integer = 0, Optional drawTextOffset As Double = 0) As Boolean
        If IsNothing(bmp) = True Then Return False
        If IsNothing(userPile) = True Then Return False
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            If IsArray(arrayProjectPoint) = True Then
                If arrayProjectPoint.Length > 1 Then
                    'рисование черным цветом
                    Dim colorPen As System.Drawing.Color = System.Drawing.Color.Black
                    'точки по верху и низу
                    Dim centerTop As Vector3D = arrayProjectPoint(0).projectPoint
                    Dim centerBottom As Vector3D = arrayProjectPoint(1).projectPoint
                    'длина по верху
                    Dim startPoint As Vector2D = Nothing
                    Dim endPoint As Vector2D = Nothing
                    Dim widthPile As Double = userPile.Width
                    If widthPile = 0 Then widthPile = userPile.Diameter

                    'по верху сторона (не зависит от типа сваи)
                    If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                        startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                        endPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                    Else
                        startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                        endPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                    End If
                    Dim x1 As Double = startPoint.X * k + offsetX
                    Dim y1 As Double = startPoint.Y * k + offsetY
                    Dim x2 As Double = endPoint.X * k + offsetX
                    Dim y2 As Double = endPoint.Y * k + offsetY
                    Dim startLine As New System.Drawing.Point(x1, y1)
                    Dim endLine As New System.Drawing.Point(x2, y2)
                    Using pen As New System.Drawing.Pen(colorPen, 1)
                        g.DrawLine(pen, startLine, endLine)
                    End Using
                    If userPile.Type = TypePile.Prismatic Then
                        'длина слева
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'длина справа
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'левый наконечник
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X, centerBottom.Z + 0.5)
                        Else
                            startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y, centerBottom.Z + 0.5)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'правый наконечник
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X, centerBottom.Z + 0.5)
                        Else
                            startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y, centerBottom.Z + 0.5)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                    Else
                        'свая круглая рисуем уширение
                        Dim dh As Double = 0
                        If userPile.Expand = True Then
                            Dim deg As Double = (userPile.DegExpand * Math.PI) / 180.0
                            dh = userPile.WidthExpand * Math.Sin(deg) + userPile.HeightDownExpand
                        End If
                        'длина слева
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerTop.X - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                        Else
                            startPoint = New Vector2D(centerTop.Y - widthPile / 2, centerTop.Z)
                            endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        'длина справа
                        If userPile.Expand = True Then
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                            Else
                                startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                            End If
                        Else
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerTop.X + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerTop.Y + widthPile / 2, centerTop.Z)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            End If
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                        '================================================================================
                        'рисуем уширение
                        If userPile.Expand = True Then
                            Dim l As Double = userPile.WidthExpand
                            Dim deg As Double = (userPile.DegExpand * Math.PI) / 180.0
                            dh = l * Math.Sin(deg) + userPile.HeightDownExpand
                            '1 длинное уширение вниз
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 возвращаем на диаметр
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2 - l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 вниз по свае
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '============================================================================
                            'правое направление
                            '1 длинное уширение вниз
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand - dh - userPile.HeightExpand)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 возвращаем на диаметр
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2 + l, centerBottom.Z - userPile.HeightDownExpand - dh)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                            '1 вниз по свае
                            If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                                startPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                            Else
                                startPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z - userPile.HeightDownExpand)
                                endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                            End If
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                            startLine = New System.Drawing.Point(x1, y1)
                            endLine = New System.Drawing.Point(x2, y2)
                            Using pen As New Pen(colorPen, 1)
                                g.DrawLine(pen, startLine, endLine)
                            End Using
                        End If
                        'перемычка по низу
                        If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                            startPoint = New Vector2D(centerBottom.X - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.X + widthPile / 2, centerBottom.Z)
                        Else
                            startPoint = New Vector2D(centerBottom.Y - widthPile / 2, centerBottom.Z)
                            endPoint = New Vector2D(centerBottom.Y + widthPile / 2, centerBottom.Z)
                        End If
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                        startLine = New System.Drawing.Point(x1, y1)
                        endLine = New System.Drawing.Point(x2, y2)
                        Using pen As New Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                    End If
                    '==================================================================================================
                    ' Подписываем сваи
                    If prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                        x1 = centerTop.X * k + offsetX
                        y1 = centerBottom.Z * k + offsetY
                    Else
                        x1 = centerTop.Y * k + offsetX
                        y1 = centerBottom.Z * k + offsetY
                    End If
                    Using font As New System.Drawing.Font("Arial", 8)
                        Using Brush As New SolidBrush(Color.Blue)

                            If (numberPile Mod 2) = 0 Then
                                g.DrawString(numberPile, font, Brush, x1, y1 + 15)
                            Else
                                g.DrawString(numberPile, font, Brush, x1, y1 + 7)
                            End If
                        End Using
                    End Using
                    If drawTextOffset <> 0 Then
                        If numberPile = 1 Then
                            Using font As New System.Drawing.Font("Arial", 8)
                                Using brush As New SolidBrush(Color.Blue)
                                    g.DrawString(Math.Round(drawTextOffset * 1000, 0), font, brush, x1 - 30, y1 - 5)
                                End Using
                            End Using
                        Else
                            Using font As New System.Drawing.Font("Arial", 8)
                                Using brush As New SolidBrush(Color.Blue)
                                    g.DrawString(Math.Round(drawTextOffset * 1000, 0), font, brush, x1 + 30, y1 - 5)
                                End Using
                            End Using
                        End If
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Using
        Return True
    End Function

    'функция делает сортировку свай
    Public Shared Function sortListPiles(ByVal listPiles As List(Of StructureElement), ByVal numberPillar As Integer) As Dictionary(Of Integer, StructureElement)
        Dim result As New Dictionary(Of Integer, StructureElement)
        If listPiles.Count = 0 Then Return result
        For i As Integer = 0 To listPiles.Count - 1
            Dim dataPile As StructureElement = listPiles.Item(i)
            Dim userPile As PilePillar = dataPile.getPilePillar
            If userPile.NumberPillar = numberPillar Then
                Dim numberColl As Integer = userPile.NumberColumn
                Dim numberRow As Integer = userPile.NumberRow
                Dim ind As Integer = MathFunction.CombineNumbers(numberRow, numberColl)
                If result.ContainsKey(ind) = False Then
                    result.Add(ind, dataPile)
                End If
            End If
        Next i
        If result.Count > 1 Then
            result = result.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return result
    End Function



End Class
