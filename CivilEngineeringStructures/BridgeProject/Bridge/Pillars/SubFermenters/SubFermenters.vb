Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports System.Windows.Media.Media3D
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.GeometriesGraph
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg.Bridges
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D

Public Class SubFermenters
    ' Приватные поля класса
    Private _numberPillar As Integer                                  ' Номер опоры
    Private _numberSubPillar As Integer                               ' Номер подопоры
    Private _numberProlet As Integer                                  ' Номер пролета
    Private _numberRow As Integer                                     ' номер ряда балок
    Private _lenght As Double                                         ' Длина подферменника поперек насадки или ригеля
    Private _width As Double                                          ' Ширина подферменника вдоль насадки или ригеля
    Private _height As Double                                         ' Высота подферменника в районе точки опирания балок 
    Private _deltaHeightBeam As Double                                ' Высота просвета между балкой и верхом подферменнка
    Private _deltaHeight As Double                                    ' Высота уширения (c)
    Private _topWidthU As Double                                      ' Длина уширения по верху (d)
    Private _bottomWidthU As Double                                   ' Длина уширения по низу (g)
    Private _topElevation As Double                                   ' Отметка верха подферменника
    Private _rotation As Double                                       ' Угол поворота подопоры
    Private _offsetX As Double                                        ' смещение вдоль ригеля или насадки
    Private _offsetY As Double                                        ' смещение вдоль оси опоры
    Private _singleSubFarmer As Boolean                               ' единый подферменник
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _numberSubPillar = 0
        _numberProlet = 0
        _numberRow = 0
        _lenght = 0.0
        _width = 0.5
        _height = 0
        _deltaHeightBeam = 0.0
        _deltaHeight = 0.0
        _topWidthU = 0.0
        _bottomWidthU = 0.0
        _topElevation = 0.0
        _rotation = 0.0
        _offsetX = 0
        _offsetY = 0
        _singleSubFarmer = False
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
    <Description("Номер пролета")>
    <Category("Свойства")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property NumberProlet() As Integer
        Get
            Return _numberProlet
        End Get
        Set(value As Integer)
            If value >= 0 Then
                _numberProlet = value
            End If
        End Set
    End Property
    <Browsable(False)>
    <Description("Номер ряда балок")>
    <Category("Свойства")>
    <DisplayName("Номер ряда балок")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Номер ряда балок")>
    <Category("Свойства")>
    <DisplayName("Номер ряда балок")>
    <[ReadOnly](False)>
    Public ReadOnly Property NumberRowStr() As String
        Get
            Return FuncFormatZn.getConditionalRow(NumberRow)
        End Get
    End Property

    <Browsable(True)>
    <Description("Длина подферменника поперек насадки или ригеля, м")>
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
    <Description("Ширина подферменника вдоль насадки или ригеля, м")>
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
    <Description("Высота подферменника в районе точки опирания балок, м")>
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
    <Description("Высота просвета между балкой (точки опирания балки) и верхом подферменнка, м")>
    <Category("Свойства")>
    <DisplayName("Просвет с балкой")>
    Public Property DeltaHeightBeam() As Double
        Get
            Return _deltaHeightBeam
        End Get
        Set(value As Double)
            _deltaHeightBeam = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота уширения (c), м")>
    <Category("Свойства")>
    <DisplayName("Высота уширения")>
    Public Property DeltaHeight() As Double
        Get
            Return _deltaHeight
        End Get
        Set(value As Double)
            _deltaHeight = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина уширения по верху (d), м")>
    <Category("Свойства")>
    <DisplayName("Длина уширения по верху")>
    Public Property TopWidthU() As Double
        Get
            Return _topWidthU
        End Get
        Set(value As Double)
            _topWidthU = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина уширения по низу (g), м")>
    <Category("Свойства")>
    <DisplayName("Длина уширения по низу")>
    Public Property BottomWidthU() As Double
        Get
            Return _bottomWidthU
        End Get
        Set(value As Double)
            _bottomWidthU = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха подферменника, м")>
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
    <Description("Единый подферменник")>
    <Category("Свойства")>
    <DisplayName("Единый подферменник")>
    Public Property SingleSubFarmer() As Boolean
        Get
            Return _singleSubFarmer
        End Get
        Set(value As Boolean)
            _singleSubFarmer = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Угол поворота подферменника")>
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
    <Description("Смещение вдоль ригеля или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение X")>
    Public Property OffsetX() As Double
        Get
            Return _offsetX
        End Get
        Set(value As Double)
            _offsetX = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение вдоль оси балки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение Y")>
    Public Property OffsetY() As Double
        Get
            Return _offsetY
        End Get
        Set(value As Double)
            _offsetY = value
        End Set
    End Property
    'создать новый подферменник
    Public Shared Function createSubFermenterPillar(ByVal idBridge As String) As StructureElement
        Dim elementSubFermenterPillar As StructureElement = New StructureElement()
        elementSubFermenterPillar.Label = "Мосты и путепроводы"
        elementSubFermenterPillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementSubFermenterPillar.ClassObject = StructureElement.classStructure.SubFermenters
        elementSubFermenterPillar.Name = StructureElement.typeObject.axisSubFermenters
        elementSubFermenterPillar.Description = "Подферменник (ось)"
        elementSubFermenterPillar.KeyParameter = ""
        elementSubFermenterPillar.IdElement = Guid.NewGuid.ToString
        elementSubFermenterPillar.IdStructure = idBridge
        elementSubFermenterPillar.Note = ""
        elementSubFermenterPillar.DWGEntity = New DwgLine()
        Return elementSubFermenterPillar
    End Function
    'ищет все подферменники для указанной опоры
    Public Shared Function getSubFermentersPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As SubFermenters()
        Dim result As SubFermenters() = {}
        Dim countElement As Integer = 0
        Dim dictPrevSubFerm As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        Dim dictNextSubFerm As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisSubFermenters) = True Then
            Dim listAxisSubFerm = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisSubFermenters)
            If IsNothing(listAxisSubFerm) = False Then
                If listAxisSubFerm.Count > 0 Then
                    For k As Integer = 0 To listAxisSubFerm.Count - 1
                        Dim tempData As StructureElement = listAxisSubFerm.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisSubFerm As SubFermenters = tempData.getSubFermenters
                            If IsNothing(userAxisSubFerm) = False Then
                                If numberPillar = userAxisSubFerm.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisSubFerm.NumberSubPillar
                                    If numberSubPillar = userAxisSubFerm.NumberSubPillar Then
                                        Dim numRow As Integer = userAxisSubFerm.NumberRow
                                        Dim numProlet As Integer = userAxisSubFerm.NumberProlet
                                        If numProlet < numberPillar Then
                                            If dictPrevSubFerm.ContainsKey(numRow) = False Then
                                                dictPrevSubFerm.Add(numRow, listAxisSubFerm.Item(k))
                                            End If
                                        Else
                                            If dictNextSubFerm.ContainsKey(numRow) = False Then
                                                dictNextSubFerm.Add(numRow, listAxisSubFerm.Item(k))
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
        If dictPrevSubFerm.Count > 1 Then
            dictPrevSubFerm = dictPrevSubFerm.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            For i As Integer = 0 To dictPrevSubFerm.Count - 1
                Dim dataSubFerm As StructureElement = dictPrevSubFerm.ElementAt(i).Value
                Dim userSubFerm As SubFermenters = dataSubFerm.getSubFermenters()
                ReDim Preserve result(countElement)
                result(countElement) = userSubFerm
                countElement += 1
            Next i
        End If
        If dictNextSubFerm.Count > 1 Then
            dictNextSubFerm = dictNextSubFerm.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            For i As Integer = 0 To dictNextSubFerm.Count - 1
                Dim dataSubFerm As StructureElement = dictNextSubFerm.ElementAt(i).Value
                Dim userSubFerm As SubFermenters = dataSubFerm.getSubFermenters()
                ReDim Preserve result(countElement)
                result(countElement) = userSubFerm
                countElement += 1
            Next i
        End If
        Return result
    End Function
    'ищет подферменник
    Public Shared Function getAxis(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberProlet As Integer, ByVal numberRow As Integer, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim result As List(Of StructureElement) = New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisSubFermenters) = True Then
            Dim listAxisSubFerm = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisSubFermenters)
            If IsNothing(listAxisSubFerm) = False Then
                If listAxisSubFerm.Count > 0 Then
                    For k As Integer = 0 To listAxisSubFerm.Count - 1
                        Dim tempData As StructureElement = listAxisSubFerm.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisSubFerm As SubFermenters = tempData.getSubFermenters
                            If IsNothing(userAxisSubFerm) = False Then
                                If numberPillar = userAxisSubFerm.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisSubFerm.NumberSubPillar
                                    If numberSubPillar = userAxisSubFerm.NumberSubPillar Then
                                        If numberProlet = userAxisSubFerm.NumberProlet Then
                                            If numberRow = userAxisSubFerm.NumberRow Then
                                                result.Add(tempData)
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
        Return result
    End Function
    'чтение данных из датагрид
    Public Shared Function readPropertiesSubFermenter(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_SubFermenter As DataGridView, Optional lastPillar As Boolean = False, Optional singleSubFermenter As Boolean = False) As SubFermenters()
        Dim result As SubFermenters() = {}
        Dim countUserSubFermenter As Integer = 0
        If DGV_SubFermenter.ColumnCount > 1 Then
            For i As Integer = 1 To DGV_SubFermenter.ColumnCount - 1
                Dim numberProlet As Integer = 0
                If lastPillar = True Then
                    If numbPillar = 1 Then
                        numberProlet = 1
                    Else
                        numberProlet = numbPillar - 1
                    End If
                End If
                Dim numberRow As Integer = DGV_SubFermenter.Columns.Item(i).Tag
                Dim lenght As Double = 0
                Dim width As Double = 0
                Dim deltaH As Double = 0
                Dim deltaHeight As Double = 0
                Dim topWidthU As Double = 0
                Dim bottomWidthU As Double = 0
                If DGV_SubFermenter.RowCount > 1 Then
                    For j As Integer = 0 To DGV_SubFermenter.RowCount - 1
                        Dim tag As String = DGV_SubFermenter.Rows(j).Tag
                        Dim value As String = DGV_SubFermenter.Rows(j).Cells(i).Value
                        If IsNothing(tag) = False And IsNothing(value) = False Then
                            If tag Like "numberProlet" Or tag Like "numberColl" Then
                                If IsNumeric(value) = True Then
                                    numberProlet = Val(value)
                                End If
                            ElseIf tag Like "lenght" Then
                                If IsNumeric(value) = True Then
                                    lenght = Val(value)
                                End If
                                If singleSubFermenter = True Then
                                    lenght = 0
                                End If
                            ElseIf tag Like "width" Then
                                If IsNumeric(value) = True And value > 0 Then
                                    width = Val(value)
                                End If
                            ElseIf tag Like "deltaHeightBeam" Then
                                If IsNumeric(value) = True Then
                                    deltaH = Val(value)
                                End If
                            ElseIf tag Like "deltaHeight" Then
                                If IsNumeric(value) = True Then
                                    deltaHeight = Val(value)
                                End If
                            ElseIf tag Like "topWidthU" Then
                                If IsNumeric(value) = True Then
                                    topWidthU = Val(value)
                                End If
                            ElseIf tag Like "bottomWidthU" Then
                                If IsNumeric(value) = True Then
                                    bottomWidthU = Val(value)
                                End If
                            End If
                        End If
                    Next j
                End If
                'единый подферменник
                If lastPillar = False Then 'промежуточная опора
                    Dim userSubFerm1 As SubFermenters = New SubFermenters
                    userSubFerm1.NumberPillar = numbPillar
                    userSubFerm1.NumberProlet = numberProlet
                    userSubFerm1.NumberSubPillar = numbSubPillar
                    userSubFerm1.NumberRow = Val(DGV_SubFermenter.Columns.Item(i).Tag) 'номер ряда
                    If lenght = 0 Then
                        userSubFerm1.SingleSubFarmer = True
                        userSubFerm1.Lenght = 0
                    Else
                        userSubFerm1.SingleSubFarmer = False
                        userSubFerm1.Lenght = lenght
                    End If
                    If width = 0 Then
                        MsgBox("Не указана ширина подферменника в ряду №" & userSubFerm1.NumberRow & " для опоры №" & userSubFerm1.NumberPillar, MsgBoxStyle.Critical, "Ошибка")
                    Else
                        userSubFerm1.Width = width
                    End If
                    userSubFerm1.DeltaHeightBeam = deltaH
                    userSubFerm1.DeltaHeight = deltaHeight
                    userSubFerm1.TopWidthU = topWidthU
                    userSubFerm1.BottomWidthU = bottomWidthU
                    ReDim Preserve result(countUserSubFermenter)
                    result(countUserSubFermenter) = userSubFerm1
                    countUserSubFermenter += 1
                Else
                    'первая или последняя опора
                    Dim userSubFerm1 As SubFermenters = New SubFermenters
                    userSubFerm1.NumberPillar = numbPillar
                    If numbPillar = 1 Then
                        userSubFerm1.NumberProlet = numbPillar '1-пролет
                    Else
                        userSubFerm1.NumberProlet = numbPillar - 1 'последний пролет
                    End If
                    userSubFerm1.NumberSubPillar = numbSubPillar
                    userSubFerm1.NumberRow = Val(DGV_SubFermenter.Columns.Item(i).Tag) 'номер ряда
                    userSubFerm1.Lenght = 0
                    userSubFerm1.SingleSubFarmer = True
                    If width = 0 Then
                        MsgBox("Не указана ширина подферменника в ряду №" & userSubFerm1.NumberRow & " для опоры №" & userSubFerm1.NumberPillar, MsgBoxStyle.Critical, "Ошибка")
                    Else
                        userSubFerm1.Width = width
                    End If
                    userSubFerm1.DeltaHeightBeam = deltaH
                    userSubFerm1.DeltaHeight = deltaHeight
                    userSubFerm1.TopWidthU = topWidthU
                    userSubFerm1.BottomWidthU = bottomWidthU
                    ReDim Preserve result(countUserSubFermenter)
                    result(countUserSubFermenter) = userSubFerm1
                    countUserSubFermenter += 1
                End If
            Next i
        End If
        Return result
    End Function
    'запись данных в датогрид
    Public Function writePropertiesSubFermenters(ByRef DGV_SubFermenter As DataGridView, ByVal indexColumn As Integer, Optional singleSubFerm As Boolean = False) As Boolean
        If indexColumn < 1 Then Return False
        If IsNothing(DGV_SubFermenter) = True Then Return False
        If DGV_SubFermenter.ColumnCount >= indexColumn Then
            For i As Integer = 0 To DGV_SubFermenter.RowCount - 1
                Dim tag As String = DGV_SubFermenter.Rows(i).Tag
                If IsNothing(tag) = False Then
                    If tag Like "numberProlet" Or tag Like "numberColl" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = NumberProlet
                    ElseIf tag Like "numberRow" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = FuncFormatZn.getConditionalRow(NumberRow)
                        DGV_SubFermenter.Columns.Item(indexColumn).Tag = NumberRow
                    ElseIf tag Like "lenght" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = Lenght
                    ElseIf tag Like "width" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = Width
                    ElseIf tag Like "deltaHeightBeam" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = DeltaHeightBeam
                    ElseIf tag Like "deltaHeight" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = DeltaHeight
                    ElseIf tag Like "topWidthU" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = TopWidthU
                    ElseIf tag Like "bottomWidthU" Then
                        DGV_SubFermenter.Rows(i).Cells(indexColumn).Value = BottomWidthU
                    End If
                End If
            Next i
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_SubFermenters As DataGridView) As Boolean
        If IsNothing(DGV_SubFermenters) = True Then Return False
        If DGV_SubFermenters.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim numberTopElevation As Integer = -1
                Dim numberProlet As Integer = -1
                For i As Integer = 0 To DGV_SubFermenters.RowCount - 1
                    Dim oldTag As String = DGV_SubFermenters.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        numberTopElevation = i
                    ElseIf oldTag Like "numberProlet" Then
                        numberProlet = i
                    End If
                Next i
                If numberTopElevation = -1 Then
                    numberTopElevation = DGV_SubFermenters.RowCount - 1
                    DGV_SubFermenters.Rows.Insert(numberTopElevation)
                    DGV_SubFermenters.Rows(numberTopElevation).DefaultCellStyle.ForeColor = Color.Red
                    DGV_SubFermenters.Rows(numberTopElevation).Tag = "calc-TopElevation"
                    DGV_SubFermenters.Rows(numberTopElevation).Cells(0).Value = "Отметка верха, м"
                End If
                For i As Integer = 1 To DGV_SubFermenters.ColumnCount - 1
                    Dim numberRowSubFerm As Integer = DGV_SubFermenters.Columns.Item(i).Tag
                    If numberRowSubFerm = NumberRow Then
                        If numberProlet = -1 Then numberProlet = numberProlet
                        If numberProlet = numberProlet Then
                            DGV_SubFermenters.Rows(numberTopElevation).Cells(i).Value = TopElevation
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
        Return True
    End Function
    'предварительный расчет всех подферменников
    Public Shared Function calculateSubFermenters(ByVal userNozzle As NozzlePillar, ByVal userRigel As RigelPillar, ByVal listBeams As List(Of Dictionary(Of Integer, StructureElement)), ByRef arraySubFermenter() As SubFermenters, Optional singleSubFermenter As Boolean = False) As Dictionary(Of Integer, Dictionary(Of Integer, SubFermenters))
        Dim result As Dictionary(Of Integer, Dictionary(Of Integer, SubFermenters)) = New Dictionary(Of Integer, Dictionary(Of Integer, SubFermenters))
        Dim dictPrevSubFerm As Dictionary(Of Integer, SubFermenters) = Nothing
        Dim dictNextSubFerm As Dictionary(Of Integer, SubFermenters) = Nothing
        If IsNothing(listBeams) = True Then Return result
        If listBeams.Count = 0 Then Return result
        If IsArray(arraySubFermenter) = False Then Return result
        Dim numberPillar As Integer = 0
        Dim listPointBridge As Dictionary(Of Integer, PointStructure) = Nothing
        Dim leftPoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim leftPoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rightPoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim middlePoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim middlePoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        If IsNothing(userNozzle) = False Then
            listPointBridge = userNozzle._elementBridgePoint.ListPointModel
            If listPointBridge.Count > 3 Then
                For i As Integer = 0 To listPointBridge.Count - 1
                    Dim tempPoint As PointStructure = listPointBridge.ElementAt(i).Value
                    If tempPoint.Code.IndexOf("leftPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        leftPoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("leftPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        leftPoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("rightPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        rightPoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("rightPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        rightPoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("middlePt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        middlePoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("middlePt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        middlePoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    End If
                Next
            End If
            numberPillar = userNozzle.NumberPillar
        End If
        If IsNothing(userRigel) = False Then
            listPointBridge = userRigel._elementBridgePoint.ListPointModel
            If listPointBridge.Count > 3 Then
                For i As Integer = 0 To listPointBridge.Count - 1
                    Dim tempPoint As PointStructure = listPointBridge.ElementAt(i).Value
                    If tempPoint.Code.IndexOf("leftPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        leftPoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("leftPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        leftPoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("rightPt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        rightPoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("rightPt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        rightPoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("middlePt1") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        middlePoint1 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    ElseIf tempPoint.Code.IndexOf("middlePt2") > -1 Then
                        Dim tempPt As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                        middlePoint2 = New Topomatic.Cad.Foundation.Vector3D(tempPt.X, tempPt.Y, tempPt.Z)
                    End If
                Next
            End If
            numberPillar = userRigel.NumberPillar
        End If
        Dim centerLine As DwgLine = New DwgLine
        centerLine.StartPoint = middlePoint1
        centerLine.EndPoint = middlePoint2
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начинаем расчет
        For i As Integer = 0 To arraySubFermenter.Length - 1
            Dim userSubFerm As SubFermenters = arraySubFermenter(i)
            Dim numberSubFermPillar As Integer = userSubFerm.NumberPillar
            Dim numberSubFermProlet As Integer = userSubFerm.NumberProlet
            Dim numSubFermRow As Integer = userSubFerm.NumberRow
            'ищем балку для данного подферменника
            Dim rotationF As Double = 0
            Dim elevationTop As Double = 0
            Dim centerPointSubFerm As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
            Dim tempListBeams As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
            'последующий пролет
            If numberSubFermProlet = numberPillar Then
                If listBeams.Count > 1 Then
                    tempListBeams = listBeams(1)
                Else
                    tempListBeams = listBeams(0)
                End If
            Else
                If listBeams.Count > 0 Then
                    tempListBeams = listBeams(0)
                End If
            End If
            'находим балку для выбранного подверменника
            If tempListBeams.Count > 0 Then
                For k As Integer = 0 To tempListBeams.Count - 1
                    Dim dataBeam As StructureElement = tempListBeams.ElementAt(k).Value
                    Dim userBeam As BeamI = dataBeam.getBeamI()
                    If numSubFermRow = userBeam.numberRow Then
                        Dim tempAxisBeam As DwgLine = dataBeam.DWGEntity
                        'последующий пролет
                        If numberSubFermProlet = numberPillar Then
                            rotationF = tempAxisBeam.Rotation + Math.PI
                            If rotationF > Math.PI * 2 Then
                                rotationF -= Math.PI * 2
                            End If
                            elevationTop = tempAxisBeam.StartPoint.Z - userSubFerm.DeltaHeightBeam
                            centerPointSubFerm = tempAxisBeam.StartPoint
                        Else
                            'предыдущий пролет
                            rotationF = tempAxisBeam.Rotation
                            elevationTop = tempAxisBeam.EndPoint.Z - userSubFerm.DeltaHeightBeam
                            centerPointSubFerm = tempAxisBeam.EndPoint
                        End If
                        Exit For
                    End If
                Next k
            End If
            'рисуем прямоугольник с заданными параметрами
            Dim lenghtSubFerm As Double = userSubFerm.Lenght
            If lenghtSubFerm = 0 Then
                lenghtSubFerm = userSubFerm.Width
            End If
            Dim listPoint As List(Of Vector2D) = BridgeGeometry.createRotatedRectangle(centerPointSubFerm, lenghtSubFerm, userSubFerm.Width, rotationF)
            If listPoint.Count > 3 Then
                Dim positionPointLeft1 As Vector2D = listPoint(0) 'левая сторона
                Dim positionPointRight1 As Vector2D = listPoint(1)
                Dim positionPointRight2 As Vector2D = listPoint(2) 'правая сторона
                Dim positionPointLeft2 As Vector2D = listPoint(3)
                'слева или справа центр подферменника находится относительно серединной линии
                Dim zn As Integer = MathFunction.funcLeftOrRightPointToLinearObject(centerLine, centerPointSubFerm)
                'делаем пересечение с осевой линией
                Dim intersectCenterPoint1 As Vector2D = MathFunction.FuncFindLineIntersection(middlePoint1.Pos, middlePoint2.Pos, listPoint(2), listPoint(3))
                Dim intersectCenterPoint2 As Vector2D = MathFunction.FuncFindLineIntersection(middlePoint1.Pos, middlePoint2.Pos, listPoint(0), listPoint(1))
                'находим высоты 4 точек по краю ригеля или насадки и по центральной линии
                Dim elevationCenterPoint1 As Double = MathFunction.FuncCalcElevationByLine(middlePoint1, middlePoint2, intersectCenterPoint1)
                Dim elevationCenterPoint2 As Double = MathFunction.FuncCalcElevationByLine(middlePoint1, middlePoint2, intersectCenterPoint2)
                Dim intersectEgePoint1 As Vector2D = New Vector2D()
                Dim intersectEgePoint2 As Vector2D = New Vector2D()
                Dim elevationEgePoint1 As Double = -1
                Dim elevationEgePoint2 As Double = -1
                If zn = 1 Then
                    intersectEgePoint1 = MathFunction.FuncFindLineIntersection(rightPoint1.Pos, rightPoint2.Pos, listPoint(2), listPoint(3))
                    intersectEgePoint2 = MathFunction.FuncFindLineIntersection(rightPoint1.Pos, rightPoint2.Pos, listPoint(0), listPoint(1))
                    elevationEgePoint1 = MathFunction.FuncCalcElevationByLine(rightPoint1, rightPoint2, intersectEgePoint1)
                    elevationEgePoint2 = MathFunction.FuncCalcElevationByLine(rightPoint1, rightPoint2, intersectEgePoint2)
                ElseIf zn = -1 Then
                    intersectEgePoint1 = MathFunction.FuncFindLineIntersection(rightPoint1.Pos, rightPoint2.Pos, listPoint(2), listPoint(3))
                    intersectEgePoint2 = MathFunction.FuncFindLineIntersection(rightPoint1.Pos, rightPoint2.Pos, listPoint(0), listPoint(1))
                    elevationEgePoint1 = MathFunction.FuncCalcElevationByLine(rightPoint1, rightPoint2, intersectEgePoint1)
                    elevationEgePoint2 = MathFunction.FuncCalcElevationByLine(rightPoint1, rightPoint2, intersectEgePoint2)
                End If
                'если это единый подферменник
                If userSubFerm.SingleSubFarmer = True Then
                    listPoint(1) = intersectCenterPoint2
                    listPoint(2) = intersectCenterPoint1
                End If
                '3д точки по краям ребра
                Dim newPoint3dCenterPoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectCenterPoint1, elevationCenterPoint1)
                Dim newPoint3dCenterPoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectCenterPoint2, elevationCenterPoint2)
                Dim newPoint3dEgePoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectEgePoint1, elevationEgePoint1)
                Dim newPoint3dEgePoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectEgePoint2, elevationEgePoint2)
                'находим высоты 4 точек подферменника
                Dim elevationLeftFerm1 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint1, newPoint3dEgePoint1, listPoint(3))
                Dim elevationLeftFerm2 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint1, newPoint3dEgePoint1, listPoint(2))
                Dim elevationRightFerm2 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint2, newPoint3dEgePoint2, listPoint(0))
                Dim elevationRightFerm1 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint2, newPoint3dEgePoint2, listPoint(1))

                userSubFerm.Rotation = Math.Round(rotationF, 6)
                userSubFerm.TopElevation = Math.Round(elevationTop, 3)
                Dim listModelPoint As New Dictionary(Of Integer, PointStructure)
                Dim x As Double = Math.Round(listPoint(1).X, 3)
                Dim y As Double = Math.Round(listPoint(1).Y, 3)
                Dim z As Double = Math.Round(elevationTop, 3)
                Dim h1 As Double = Math.Round(elevationTop - elevationRightFerm1, 3)
                Dim h2 As Double = 0
                Dim code As String = "leftPt1"
                listModelPoint.Add(1, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

                x = Math.Round(listPoint(2).X, 3)
                y = Math.Round(listPoint(2).Y, 3)
                z = Math.Round(elevationTop, 3)
                h1 = Math.Round(elevationTop - elevationLeftFerm2, 3)
                h2 = 0
                code = "leftPt2"
                listModelPoint.Add(2, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

                x = Math.Round(listPoint(3).X, 3)
                y = Math.Round(listPoint(3).Y, 3)
                z = Math.Round(elevationTop, 3)
                h1 = Math.Round(elevationTop - elevationLeftFerm1, 3)
                h2 = 0
                code = "rightPt2"
                listModelPoint.Add(3, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

                x = Math.Round(listPoint(0).X, 3)
                y = Math.Round(listPoint(0).Y, 3)
                z = Math.Round(elevationTop, 3)
                h1 = Math.Round(elevationTop - elevationRightFerm2, 3)
                h2 = 0
                code = "rightPt1"
                listModelPoint.Add(4, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
                userSubFerm._elementBridgePoint.ListPointModel = listModelPoint
                'центральная точка
                userSubFerm._elementBridgePoint.CenterTopPoint = centerPointSubFerm
                'ось подферменника
                Dim axisPointStart1 As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(listPoint(1), listPoint(2))
                Dim axisPointStart2 As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(listPoint(3), listPoint(0))
                userSubFerm._elementBridgePoint.StartAxisPoint = New Vector3D(axisPointStart1, centerPointSubFerm.Z)
                userSubFerm._elementBridgePoint.EndAxisPoint = New Vector3D(axisPointStart2, centerPointSubFerm.Z)
                '==========================================================================================================
                'рисуем уширение
                If userSubFerm.DeltaHeight > 0 Then
                    If userSubFerm.TopWidthU > 0 Then
                        Dim leftPointU1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                        Dim leftPointU2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                        Dim rightPointU1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                        Dim rightPointU2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                        If numSubFermRow < 0 Then
                            leftPointU1 = New Vector3D(listPoint(1), elevationTop)
                            leftPointU2 = New Vector3D(listPoint(2), elevationTop)
                            rightPointU2 = New Vector3D(listPoint(3), elevationTop)
                            rightPointU1 = New Vector3D(listPoint(0), elevationTop)
                        Else
                            leftPointU1 = New Vector3D(listPoint(2), elevationTop)
                            leftPointU2 = New Vector3D(listPoint(1), elevationTop)
                            rightPointU2 = New Vector3D(listPoint(0), elevationTop)
                            rightPointU1 = New Vector3D(listPoint(3), elevationTop)
                        End If
                        Dim pointLeftTop As Vector3D = MathFunction.FuncCalcPointInLine(leftPointU1, leftPointU2, userSubFerm.TopWidthU)
                        Dim pointRightTop As Vector3D = MathFunction.FuncCalcPointInLine(rightPointU1, rightPointU2, userSubFerm.TopWidthU)
                        Dim pointLeftBottom As Vector3D = MathFunction.FuncCalcPointInLine(leftPointU1, leftPointU2, userSubFerm.TopWidthU + userSubFerm.BottomWidthU)
                        Dim pointRightBottom As Vector3D = MathFunction.FuncCalcPointInLine(rightPointU1, rightPointU2, userSubFerm.TopWidthU + userSubFerm.BottomWidthU)

                        Dim listModelSecondPoint As New Dictionary(Of Integer, PointStructure)
                        x = Math.Round(leftPointU1.X, 3)
                        y = Math.Round(leftPointU1.Y, 3)
                        z = Math.Round(leftPointU1.Z + userSubFerm.DeltaHeight, 3)
                        h1 = Math.Round(userSubFerm.DeltaHeight, 3)
                        h2 = 0
                        code = "leftPt1"
                        listModelSecondPoint.Add(1, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

                        x = Math.Round(pointLeftTop.X, 3)
                        y = Math.Round(pointLeftTop.Y, 3)
                        z = Math.Round(leftPointU1.Z + userSubFerm.DeltaHeight, 3)
                        Dim dx As Double = Math.Round(pointLeftBottom.X - pointLeftTop.X, 3)
                        Dim dy As Double = Math.Round(pointLeftBottom.Y - pointLeftTop.Y, 3)
                        h1 = Math.Round(userSubFerm.DeltaHeight, 3)
                        h2 = 0
                        code = "leftPt2"
                        listModelSecondPoint.Add(2, New PointStructure(x, y, z, dx, dy, -1 * h1, h2, code))

                        x = Math.Round(pointRightTop.X, 3)
                        y = Math.Round(pointRightTop.Y, 3)
                        z = Math.Round(rightPointU1.Z + userSubFerm.DeltaHeight, 3)
                        dx = Math.Round(pointRightBottom.X - pointRightTop.X, 3)
                        dy = Math.Round(pointRightBottom.Y - pointRightTop.Y, 3)
                        h1 = Math.Round(userSubFerm.DeltaHeight, 3)
                        h2 = 0
                        code = "rightPt2"
                        listModelSecondPoint.Add(4, New PointStructure(x, y, z, dx, dy, -1 * h1, h2, code))

                        x = Math.Round(rightPointU1.X, 3)
                        y = Math.Round(rightPointU1.Y, 3)
                        z = Math.Round(rightPointU1.Z + userSubFerm.DeltaHeight, 3)
                        h1 = Math.Round(userSubFerm.DeltaHeight, 3)
                        h2 = 0
                        code = "rightPt1"
                        listModelSecondPoint.Add(3, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
                        userSubFerm._elementBridgePoint.ListPointSecondModel = listModelSecondPoint
                    End If
                End If
            End If

            If result.ContainsKey(numberSubFermProlet) = True Then
                Dim dictSubFerms As Dictionary(Of Integer, SubFermenters) = result.Item(numberSubFermProlet)
                If dictSubFerms.ContainsKey(numSubFermRow) = False Then
                    dictSubFerms.Add(numSubFermRow, userSubFerm)
                    result.Item(numberSubFermProlet) = dictSubFerms
                End If
            Else
                Dim dictSubFerms As Dictionary(Of Integer, SubFermenters) = New Dictionary(Of Integer, SubFermenters)
                dictSubFerms.Add(numSubFermRow, userSubFerm)
                result.Item(numberSubFermProlet) = dictSubFerms
            End If
        Next i
        Return result
    End Function
    'предварительный расчет одного подферменника
    Public Shared Function calculateSubFermenter(ByRef userSubFermenter As SubFermenters, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Boolean
        If IsNothing(userSubFermenter) = True Then Return False
        Dim numberPillarSubFerm As Integer = userSubFermenter.NumberPillar
        Dim numberProletSubFerm As Integer = userSubFermenter.NumberProlet
        Dim numberRowSubFerm As Integer = userSubFermenter.NumberRow
        Dim centerPoint As Vector3D = New Vector3D(-1, -1, -1)
        Dim lastPoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim lastPoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim middlePoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim middlePoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Dim rotationF As Double = 0
        Dim elevationTop As Double = 0
        Dim axisBeam As DwgLine = New DwgLine()
        If numberPillarSubFerm > 0 And numberProletSubFerm > 0 Then
            Dim dataAxisBeam As StructureElement = CalculationBeams.getAxisBeam(dictionaryObjectsBridge, numberProletSubFerm, numberRowSubFerm)
            If IsNothing(dataAxisBeam) = False Then
                Dim userBeam As BeamI = dataAxisBeam.getBeamI()
                axisBeam = dataAxisBeam.DWGEntity
                If numberPillarSubFerm = numberProletSubFerm Then
                    centerPoint = userBeam._elementBridgePoint.StartAxisPoint
                    If centerPoint.X = 0 And centerPoint.Y = 0 And centerPoint.Z = 0 Then
                        centerPoint = axisBeam.StartPoint
                    End If
                    rotationF = axisBeam.Rotation + Math.PI
                    If rotationF > Math.PI * 2 Then
                        rotationF -= Math.PI * 2
                    End If
                    elevationTop = axisBeam.StartPoint.Z - userSubFermenter.DeltaHeightBeam
                Else numberPillarSubFerm = numberProletSubFerm + 1
                    centerPoint = userBeam._elementBridgePoint.EndAxisPoint
                    If centerPoint.X = 0 And centerPoint.Y = 0 And centerPoint.Z = 0 Then
                        centerPoint = axisBeam.EndPoint
                    End If
                    rotationF = axisBeam.Rotation
                    elevationTop = axisBeam.EndPoint.Z - userSubFermenter.DeltaHeightBeam
                End If
            End If
            centerPoint = New Vector3D(centerPoint.Pos, elevationTop)
            'находим ригель для подферменника
            Dim listDataRagel As List(Of StructureElement) = RigelPillar.getAxis(dictionaryObjectsBridge, numberPillarSubFerm)
            If listDataRagel.Count > 0 Then
                Dim dataRigel As StructureElement = StructureElement.isValidateDataStructure(listDataRagel)
                Dim userRigel As RigelPillar = dataRigel.getRigelPillar()
                middlePoint1 = userRigel.getPointByCode("middlePt1", True)
                middlePoint2 = userRigel.getPointByCode("middlePt2", True)
                If numberPillarSubFerm = numberProletSubFerm Then
                    lastPoint1 = userRigel.getPointByCode("rightPt1", True)
                    lastPoint2 = userRigel.getPointByCode("rightPt2", True)
                Else numberPillarSubFerm = numberProletSubFerm + 1
                    lastPoint1 = userRigel.getPointByCode("leftPt1", True)
                    lastPoint2 = userRigel.getPointByCode("leftPt2", True)
                End If
            Else
                Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryObjectsBridge, numberPillarSubFerm)
                If listDataNozzle.Count > 0 Then
                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                    Dim userNozzle As NozzlePillar = dataNozzle.getNozzlePillar()
                    middlePoint1 = userNozzle.getPointByCode("middlePt1", True)
                    middlePoint2 = userNozzle.getPointByCode("middlePt2", True)
                    If numberPillarSubFerm = numberProletSubFerm Then
                        lastPoint1 = userNozzle.getPointByCode("rightPt1", True)
                        lastPoint2 = userNozzle.getPointByCode("rightPt2", True)
                    Else numberPillarSubFerm = numberProletSubFerm + 1
                        lastPoint1 = userNozzle.getPointByCode("leftPt1", True)
                        lastPoint2 = userNozzle.getPointByCode("leftPt2", True)
                    End If
                End If
            End If
        End If
        If centerPoint.X = -1 And centerPoint.Y = -1 And centerPoint.Z = -1 Then
            Exit Function
        End If
        If middlePoint1.X = -1 And middlePoint1.Y = -1 And middlePoint1.Z = -1 Then
            Exit Function
        End If
        If middlePoint2.X = -1 And middlePoint2.Y = -1 And middlePoint2.Z = -1 Then
            Exit Function
        End If
        If lastPoint1.X = -1 And lastPoint1.Y = -1 And lastPoint1.Z = -1 Then
            Exit Function
        End If
        If lastPoint2.X = -1 And lastPoint2.Y = -1 And lastPoint2.Z = -1 Then
            Exit Function
        End If
        Dim centerLine As DwgLine = New DwgLine
        centerLine.StartPoint = middlePoint1
        centerLine.EndPoint = middlePoint2

        Dim lastLine As DwgLine = New DwgLine
        lastLine.StartPoint = lastPoint1
        lastLine.EndPoint = lastPoint2
        'делаем смещение вдоль ригеля или насадки
        Dim newCenterPoint As Vector2D = New Vector2D(0, 0)
        If userSubFermenter.OffsetX <> 0 Then
            newCenterPoint = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerPoint.Pos, centerLine.Rotation, userSubFermenter.OffsetX)
        End If
        If userSubFermenter.OffsetY <> 0 Then
            newCenterPoint = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerPoint.Pos, axisBeam.Rotation, userSubFermenter.OffsetY)
        End If
        If newCenterPoint.X <> 0 OrElse newCenterPoint.Y <> 0 Then
            centerPoint = New Vector3D(newCenterPoint, centerPoint.Z)
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начинаем расчет
        'рисуем прямоугольник с заданными параметрами
        Dim lenghtSubFerm As Double = userSubFermenter.Lenght
        If lenghtSubFerm = 0 Then
            lenghtSubFerm = userSubFermenter.Width
        End If
        Dim listPoint As List(Of Vector2D) = BridgeGeometry.createRotatedRectangle(centerPoint, lenghtSubFerm, userSubFermenter.Width, rotationF)
        If listPoint.Count > 3 Then
            Dim positionPointLeft1 As Vector2D = listPoint(0) 'левая сторона
            Dim positionPointRight1 As Vector2D = listPoint(1)
            Dim positionPointRight2 As Vector2D = listPoint(2) 'правая сторона
            Dim positionPointLeft2 As Vector2D = listPoint(3)
            'слева или справа центр подферменника находится относительно серединной линии
            Dim zn As Integer = MathFunction.funcLeftOrRightPointToLinearObject(centerLine, centerPoint)
            'делаем пересечение с осевой линией
            Dim intersectCenterPoint1 As Vector2D = MathFunction.FuncFindLineIntersection(middlePoint1.Pos, middlePoint2.Pos, listPoint(2), listPoint(3))
            Dim intersectCenterPoint2 As Vector2D = MathFunction.FuncFindLineIntersection(middlePoint1.Pos, middlePoint2.Pos, listPoint(0), listPoint(1))
            'находим высоты 4 точек по краю ригеля или насадки и по центральной линии
            Dim elevationCenterPoint1 As Double = MathFunction.FuncCalcElevationByLine(middlePoint1, middlePoint2, intersectCenterPoint1)
            Dim elevationCenterPoint2 As Double = MathFunction.FuncCalcElevationByLine(middlePoint1, middlePoint2, intersectCenterPoint2)
            Dim intersectEgePoint1 As Vector2D = MathFunction.FuncFindLineIntersection(lastPoint1.Pos, lastPoint2.Pos, listPoint(2), listPoint(3))
            Dim intersectEgePoint2 As Vector2D = MathFunction.FuncFindLineIntersection(lastPoint1.Pos, lastPoint2.Pos, listPoint(0), listPoint(1))
            Dim elevationEgePoint1 As Double = MathFunction.FuncCalcElevationByLine(lastPoint1, lastPoint2, intersectEgePoint1)
            Dim elevationEgePoint2 As Double = MathFunction.FuncCalcElevationByLine(lastPoint1, lastPoint2, intersectEgePoint2)
            'если это единый подферменник
            If userSubFermenter.SingleSubFarmer = True Then
                listPoint(1) = intersectCenterPoint2
                listPoint(2) = intersectCenterPoint1
            End If
            '3д точки по краям ребра
            Dim newPoint3dCenterPoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectCenterPoint1, elevationCenterPoint1)
            Dim newPoint3dCenterPoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectCenterPoint2, elevationCenterPoint2)
            Dim newPoint3dEgePoint1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectEgePoint1, elevationEgePoint1)
            Dim newPoint3dEgePoint2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(intersectEgePoint2, elevationEgePoint2)
            'находим высоты 4 точек подферменника
            Dim elevationLeftFerm1 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint1, newPoint3dEgePoint1, listPoint(3))
            Dim elevationLeftFerm2 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint1, newPoint3dEgePoint1, listPoint(2))
            Dim elevationRightFerm2 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint2, newPoint3dEgePoint2, listPoint(0))
            Dim elevationRightFerm1 As Double = MathFunction.FuncCalcElevationByLine(newPoint3dCenterPoint2, newPoint3dEgePoint2, listPoint(1))

            userSubFermenter.Rotation = Math.Round(rotationF, 6)
            userSubFermenter.TopElevation = Math.Round(elevationTop, 3)
            Dim listModelPoint As New Dictionary(Of Integer, PointStructure)
            Dim x As Double = Math.Round(listPoint(1).X, 3)
            Dim y As Double = Math.Round(listPoint(1).Y, 3)
            Dim z As Double = Math.Round(elevationTop, 3)
            Dim h1 As Double = Math.Round(elevationTop - elevationRightFerm1, 3)
            Dim h2 As Double = 0
            Dim code As String = "leftPt1"
            listModelPoint.Add(1, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

            x = Math.Round(listPoint(2).X, 3)
            y = Math.Round(listPoint(2).Y, 3)
            z = Math.Round(elevationTop, 3)
            h1 = Math.Round(elevationTop - elevationLeftFerm2, 3)
            h2 = 0
            code = "leftPt2"
            listModelPoint.Add(2, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

            x = Math.Round(listPoint(3).X, 3)
            y = Math.Round(listPoint(3).Y, 3)
            z = Math.Round(elevationTop, 3)
            h1 = Math.Round(elevationTop - elevationLeftFerm1, 3)
            h2 = 0
            code = "rightPt2"
            listModelPoint.Add(3, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

            x = Math.Round(listPoint(0).X, 3)
            y = Math.Round(listPoint(0).Y, 3)
            z = Math.Round(elevationTop, 3)
            h1 = Math.Round(elevationTop - elevationRightFerm2, 3)
            h2 = 0
            code = "rightPt1"
            listModelPoint.Add(4, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
            userSubFermenter._elementBridgePoint.ListPointModel = listModelPoint
            'центральная точка
            userSubFermenter._elementBridgePoint.CenterTopPoint = centerPoint
            'ось подферменника
            Dim axisPointStart1 As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(listPoint(1), listPoint(2))
            Dim axisPointStart2 As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(listPoint(3), listPoint(0))
            userSubFermenter._elementBridgePoint.StartAxisPoint = New Vector3D(axisPointStart1, elevationTop)
            userSubFermenter._elementBridgePoint.EndAxisPoint = New Vector3D(axisPointStart2, elevationTop)

            '==========================================================================================================
            'рисуем уширение
            If userSubFermenter.DeltaHeight > 0 Then
                If userSubFermenter.TopWidthU > 0 Then
                    Dim leftPointU1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                    Dim leftPointU2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                    Dim rightPointU1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                    Dim rightPointU2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
                    If numberRowSubFerm < 0 Then
                        leftPointU1 = New Vector3D(listPoint(1), elevationTop)
                        leftPointU2 = New Vector3D(listPoint(2), elevationTop)
                        rightPointU2 = New Vector3D(listPoint(3), elevationTop)
                        rightPointU1 = New Vector3D(listPoint(0), elevationTop)
                    Else
                        leftPointU1 = New Vector3D(listPoint(2), elevationTop)
                        leftPointU2 = New Vector3D(listPoint(1), elevationTop)
                        rightPointU2 = New Vector3D(listPoint(0), elevationTop)
                        rightPointU1 = New Vector3D(listPoint(3), elevationTop)
                    End If
                    Dim pointLeftTop As Vector3D = MathFunction.FuncCalcPointInLine(leftPointU1, leftPointU2, userSubFermenter.TopWidthU)
                    Dim pointRightTop As Vector3D = MathFunction.FuncCalcPointInLine(rightPointU1, rightPointU2, userSubFermenter.TopWidthU)
                    Dim pointLeftBottom As Vector3D = MathFunction.FuncCalcPointInLine(leftPointU1, leftPointU2, userSubFermenter.TopWidthU + userSubFermenter.BottomWidthU)
                    Dim pointRightBottom As Vector3D = MathFunction.FuncCalcPointInLine(rightPointU1, rightPointU2, userSubFermenter.TopWidthU + userSubFermenter.BottomWidthU)

                    Dim listModelSecondPoint As New Dictionary(Of Integer, PointStructure)
                    x = Math.Round(leftPointU1.X, 3)
                    y = Math.Round(leftPointU1.Y, 3)
                    z = Math.Round(leftPointU1.Z + userSubFermenter.DeltaHeight, 3)
                    h1 = Math.Round(userSubFermenter.DeltaHeight, 3)
                    h2 = 0
                    code = "leftPt1"
                    listModelSecondPoint.Add(1, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))

                    x = Math.Round(pointLeftTop.X, 3)
                    y = Math.Round(pointLeftTop.Y, 3)
                    z = Math.Round(leftPointU1.Z + userSubFermenter.DeltaHeight, 3)
                    Dim dx As Double = Math.Round(pointLeftBottom.X - pointLeftTop.X, 3)
                    Dim dy As Double = Math.Round(pointLeftBottom.Y - pointLeftTop.Y, 3)
                    h1 = Math.Round(userSubFermenter.DeltaHeight, 3)
                    h2 = 0
                    code = "leftPt2"
                    listModelSecondPoint.Add(2, New PointStructure(x, y, z, dx, dy, -1 * h1, h2, code))

                    x = Math.Round(pointRightTop.X, 3)
                    y = Math.Round(pointRightTop.Y, 3)
                    z = Math.Round(rightPointU1.Z + userSubFermenter.DeltaHeight, 3)
                    dx = Math.Round(pointRightBottom.X - pointRightTop.X, 3)
                    dy = Math.Round(pointRightBottom.Y - pointRightTop.Y, 3)
                    h1 = Math.Round(userSubFermenter.DeltaHeight, 3)
                    h2 = 0
                    code = "rightPt2"
                    listModelSecondPoint.Add(4, New PointStructure(x, y, z, dx, dy, -1 * h1, h2, code))

                    x = Math.Round(rightPointU1.X, 3)
                    y = Math.Round(rightPointU1.Y, 3)
                    z = Math.Round(rightPointU1.Z + userSubFermenter.DeltaHeight, 3)
                    h1 = Math.Round(userSubFermenter.DeltaHeight, 3)
                    h2 = 0
                    code = "rightPt1"
                    listModelSecondPoint.Add(3, New PointStructure(x, y, z, 0, 0, -1 * h1, h2, code))
                    userSubFermenter._elementBridgePoint.ListPointSecondModel = listModelSecondPoint
                End If
            End If
        End If
        Return True
    End Function
    'рисование оси подферменника
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLineSubFermenter As DwgLine = Nothing
        Dim dataSubFermenter As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getAxis(dictionaryBridgeElements, NumberPillar, NumberProlet, NumberRow, NumberSubPillar)
        If listAxis.Count = 0 Then
            dataSubFermenter = createSubFermenterPillar(idBridge)
        ElseIf listAxis.Count = 1 Then
            dataSubFermenter = listAxis.Item(0)
        Else
            dataSubFermenter = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataSubFermenter) Then Return dataSubFermenter
        axisLineSubFermenter = dataSubFermenter.DWGEntity
        If IsNothing(axisLineSubFermenter) = True Then axisLineSubFermenter = New DwgLine
        If axisLineSubFermenter.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Подферменник (ось)")
            styleAxisNozzle.setObjectStyle(axisLineSubFermenter)
        End If
        'ось насадки
        axisLineSubFermenter.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineSubFermenter.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLineSubFermenter.StartPoint.Pos - axisLineSubFermenter.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось подферменника имеет нулевое значение. Ось не построена.")
            Return dataSubFermenter
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineSubFermenter) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineSubFermenter)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataSubFermenter.KeyParameter = strJson
        dataSubFermenter.DWGEntity = axisLineSubFermenter
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineSubFermenter, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataSubFermenter)
        Return dataSubFermenter
    End Function
    'функция возвращает все точки подферменника в виде словаря 
    Public Shared Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure, ByVal numberPillar As Integer, ByVal arraySubFerm As SubFermenters(), ByRef dictPrevProletSubFerm As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint))), ByRef dictNextProletSubFerm As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))) As Boolean
        If numberPillar < 1 Then Return False
        If IsNothing(arraySubFerm) = True Then Return False
        If IsArray(arraySubFerm) = False Then Return False
        dictPrevProletSubFerm = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
        dictNextProletSubFerm = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
        For i As Integer = 0 To arraySubFerm.Length - 1
            Dim userSubFerm As SubFermenters = arraySubFerm(i)
            If IsNothing(userSubFerm) = True Then Continue For
            If userSubFerm.NumberPillar < 1 Then Continue For
            Dim listElementSubFerm As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
            'точки по верху подферменника
            Dim transformPointTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
            'точки по низу подферменника
            Dim transformPointBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
            If userSubFerm._elementBridgePoint.ListPointModel.Count > 3 Then
                'записываем результат
                For j As Integer = 0 To userSubFerm._elementBridgePoint.ListPointModel.Count - 1
                    Dim pointStructure As PointStructure = userSubFerm._elementBridgePoint.ListPointModel.ElementAt(j).Value
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
            listElementSubFerm.Add(transformPointTop)
            listElementSubFerm.Add(transformPointBottom)
            If userSubFerm.NumberProlet < numberPillar Then
                If dictPrevProletSubFerm.ContainsKey(userSubFerm.NumberRow) = False Then
                    dictPrevProletSubFerm.Add(userSubFerm.NumberRow, listElementSubFerm)
                End If
            Else
                If dictNextProletSubFerm.ContainsKey(userSubFerm.NumberRow) = False Then
                    dictNextProletSubFerm.Add(userSubFerm.NumberRow, listElementSubFerm)
                End If
            End If
        Next i
        If dictPrevProletSubFerm.Count > 1 Then
            dictPrevProletSubFerm = dictPrevProletSubFerm.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        If dictNextProletSubFerm.Count > 1 Then
            dictNextProletSubFerm = dictNextProletSubFerm.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return True
    End Function
    'функция делает сортировку подферменников
    Public Shared Function sortListSubFermenter(ByVal listSubFermenter As List(Of StructureElement), ByVal numberPillar As Integer) As Dictionary(Of Integer, StructureElement)
        Dim result As New Dictionary(Of Integer, StructureElement)
        If listSubFermenter.Count = 0 Then Return result
        For i As Integer = 0 To listSubFermenter.Count - 1
            Dim dataSubFerm As StructureElement = listSubFermenter.Item(i)
            Dim userSubFerm As SubFermenters = dataSubFerm.getSubFermenters()
            If userSubFerm.NumberPillar = numberPillar Then
                Dim numberProlet As Integer = userSubFerm.NumberProlet
                Dim numberRow As Integer = userSubFerm.NumberRow + 100
                Dim ind As Integer = numberProlet * numberRow
                result.Add(ind, dataSubFerm)
            End If
        Next i
        If result.Count > 1 Then
            result = result.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        Return result
    End Function
End Class
