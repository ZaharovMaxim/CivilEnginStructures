Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Algorithm
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class PreparationPillar
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _number As Integer                        ' Номер подготовки
    Private _height As Double                         ' Высота подготовки
    Private _offsetLengthTop As Double                ' Смещение сверху ростверка или насадки
    Private _offsetLengthBottom As Double             ' Смещение снизу ростверка или насадки
    Private _material As String                       ' Материал подготовки
    Private _topElevation As Double                   ' Отметка верха подготовки
    Private _bottomElevation As Double                ' Отметка низа подготовки
    Private _topLength As Double                      ' Длина подготовки по верху
    Private _bottomLength As Double                   ' Длина подготовки по низу
    Private _topWidth As Double                       ' Ширина подготовки по верху
    Private _bottomWidth As Double                    ' Ширина подготовки по низу
    Private _model As String
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _number = 0
        _height = 0.0
        _offsetLengthTop = 0.0
        _offsetLengthBottom = 0.0
        _material = String.Empty
        _topElevation = 0.0
        _bottomElevation = 0.0
        _topLength = 0.0
        _bottomLength = 0.0
        _topWidth = 0.0
        _bottomWidth = 0.0
        _elementBridgePoint = New PointsCollections
        _model = ""
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
    <Description("Высота подготовки, м")>
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
    <Description("Отступ верха подготовки от ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Отступ верха")>
    Public Property OffsetLengthTop() As Double
        Get
            Return _offsetLengthTop
        End Get
        Set(value As Double)
            _offsetLengthTop = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отступ низа подготовки от ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Отступ низа")>
    Public Property OffsetLengthBottom() As Double
        Get
            Return _offsetLengthBottom
        End Get
        Set(value As Double)
            _offsetLengthBottom = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Материал подготовки")>
    <Category("Свойства")>
    <DisplayName("Материал подготовки")>
    Public Property Material() As String
        Get
            Return _material
        End Get
        Set(value As String)
            _material = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка верха подготовки, м")>
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
    <Description("Отметка низа подготовки, м")>
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
    <Description("Длина подготовки по верху, м")>
    <Category("Свойства")>
    <DisplayName("Длина по верху")>
    <[ReadOnly](True)>
    Public Property TopLength() As Double
        Get
            Return _topLength
        End Get
        Set(value As Double)
            If value >= 0 Then
                _topLength = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Длина подготовки по низу, м")>
    <Category("Свойства")>
    <DisplayName("Длина по низу")>
    <[ReadOnly](True)>
    Public Property BottomLength() As Double
        Get
            Return _bottomLength
        End Get
        Set(value As Double)
            If value >= 0 Then
                _bottomLength = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина подготовки по верху, м")>
    <Category("Свойства")>
    <DisplayName("Ширина по верху")>
    <[ReadOnly](True)>
    Public Property TopWidth() As Double
        Get
            Return _topWidth
        End Get
        Set(value As Double)
            If value >= 0 Then
                _topWidth = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина подготовки по низу, м")>
    <Category("Свойства")>
    <DisplayName("Ширина по низу")>
    <[ReadOnly](True)>
    Public Property BottomWidth() As Double
        Get
            Return _bottomWidth
        End Get
        Set(value As Double)
            If value >= 0 Then
                _bottomWidth = value
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
    Public Shared Function createAxisPreparationPillar(ByVal idBridge As String) As StructureElement
        Dim elementPreparation As StructureElement = New StructureElement()
        elementPreparation.Label = "Мосты и путепроводы"
        elementPreparation.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPreparation.ClassObject = StructureElement.classStructure.PreparationPillar
        elementPreparation.Name = StructureElement.typeObject.axisPreparation
        elementPreparation.Description = "Подготовка (ось)"
        elementPreparation.KeyParameter = ""
        elementPreparation.IdElement = Guid.NewGuid.ToString
        elementPreparation.IdStructure = idBridge
        elementPreparation.Note = ""
        elementPreparation.DWGEntity = New DwgLine()
        Return elementPreparation
    End Function
    'ищет подготовку
    Public Shared Function getPreparationPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataPreparationPillar As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPreparation) = True Then
            Dim listAxisPreparation = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPreparation)
            If IsNothing(listAxisPreparation) = False Then
                If listAxisPreparation.Count > 0 Then
                    For k As Integer = 0 To listAxisPreparation.Count - 1
                        Dim tempData As StructureElement = listAxisPreparation.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisPreparation As PreparationPillar = tempData.getPreparationPillar
                            If IsNothing(userAxisPreparation) = False Then
                                If numberPillar = userAxisPreparation.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userAxisPreparation.Number
                                    If numberSubPillar = userAxisPreparation.Number Then
                                        dataPreparationPillar = tempData
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataPreparationPillar
    End Function
    Public Shared Function readPropertiesPreparation(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Preparation As DataGridView) As PreparationPillar
        Dim result As PreparationPillar = New PreparationPillar
        result.NumberPillar = numbPillar
        result.Number = numbSubPillar
        If DGV_Preparation.RowCount > 1 Then
            For i As Integer = 0 To DGV_Preparation.RowCount - 1
                Dim tag As String = DGV_Preparation.Rows(i).Tag
                Dim value As String = DGV_Preparation.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_preparation_width" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Height = Val(value)
                        Else
                            MsgBox("Некорректное значение высоты подготовки.")
                        End If
                    ElseIf tag Like "offsetLenghtTop" Then
                        If IsNumeric(value) = True Then
                            result.OffsetLengthTop = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа верха подготовки от края ростверка.")
                        End If
                    ElseIf tag Like "offsetLenghtBottom" Then
                        If IsNumeric(value) = True Then
                            result.OffsetLengthBottom = Val(value)
                        Else
                            MsgBox("Некорректное значение отступа низа подготовки от края ростверка.")
                        End If
                    ElseIf tag Like "material" Then
                        result.Material = value
                    End If
                End If
            Next i
        End If
        Return result
    End Function
    Public Function writePropertiesPreparation(ByRef DGV_Preparation As DataGridView) As Boolean
        If IsNothing(DGV_Preparation) = True Then Return False
        If DGV_Preparation.RowCount > 1 Then
            For j As Integer = 0 To DGV_Preparation.RowCount - 1
                Dim tag As String = DGV_Preparation.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_preparation_width" Then
                        DGV_Preparation.Rows(j).Cells(1).Value = Height
                    ElseIf tag Like "offsetLenghtTop" Then
                        DGV_Preparation.Rows(j).Cells(1).Value = OffsetLengthTop
                    ElseIf tag Like "offsetLenghtBottom" Then
                        DGV_Preparation.Rows(j).Cells(1).Value = OffsetLengthBottom
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Preparation As DataGridView) As Boolean
        If IsNothing(DGV_Preparation) = True Then Return False
        If DGV_Preparation.RowCount > 1 Then
            If _elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolBottomElevation As Boolean = False
                Dim boolTopLenght As Boolean = False
                Dim boolBottomLenght As Boolean = False
                Dim boolTopWidth As Boolean = False
                Dim boolBottomWidth As Boolean = False
                For i As Integer = 0 To DGV_Preparation.RowCount - 1
                    Dim oldTag As String = DGV_Preparation.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = TopElevation
                        boolTopElevation = True
                    ElseIf oldTag Like "calc-BottomElevation" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = BottomElevation
                        boolBottomElevation = True
                    ElseIf oldTag Like "calc-TopLength" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = TopLength
                        boolTopLenght = True
                    ElseIf oldTag Like "calc-BottomLenght" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = BottomLength
                        boolBottomLenght = True
                    ElseIf oldTag Like "calc-TopWidth" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = TopWidth
                        boolTopWidth = True
                    ElseIf oldTag Like "calc-BottomWidth" Then
                        DGV_Preparation.Rows(i).Cells(1).Value = BottomWidth
                        boolBottomWidth = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Отметка верха, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = TopElevation
                    DGV_Preparation.Rows(numberRow).Tag = "calc-TopElevation"
                End If
                If boolBottomElevation = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Отметка низа, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = BottomElevation
                    DGV_Preparation.Rows(numberRow).Tag = "calc-BottomElevation"
                End If
                If boolTopLenght = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Длина по верху, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = TopLength
                    DGV_Preparation.Rows(numberRow).Tag = "calc-TopLength"
                End If
                If boolBottomLenght = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Длина по низу, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = BottomLength
                    DGV_Preparation.Rows(numberRow).Tag = "calc-BottomLenght"
                End If
                If boolTopWidth = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Ширина по верху, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = TopWidth
                    DGV_Preparation.Rows(numberRow).Tag = "calc-TopWidth"
                End If
                If boolBottomWidth = False Then
                    Dim numberRow As Integer = DGV_Preparation.RowCount - 1
                    DGV_Preparation.Rows.Insert(numberRow)
                    DGV_Preparation.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Preparation.Rows(numberRow).Cells(0).Value = "Ширина по низу, м"
                    DGV_Preparation.Rows(numberRow).Cells(1).Value = BottomWidth
                    DGV_Preparation.Rows(numberRow).Tag = "calc-BottomWidth"
                End If
            End If
        End If
        Return True
    End Function
    Public Function calculatePreparation(ByVal userGrillage As GrillagePillar, ByVal userNozzle As NozzlePillar, Optional align As Polyline3D = Nothing) As Boolean
        Dim dictionaryListPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        If IsNothing(userGrillage) = False Then
            If userGrillage._elementBridgePoint.ListPointModel.Count > 3 Then
                dictionaryListPoint = userGrillage._elementBridgePoint.ListPointModel
            End If
        End If
        If dictionaryListPoint.Count < 4 Then
            If IsNothing(userNozzle) = False Then
                If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
                    dictionaryListPoint = userNozzle._elementBridgePoint.ListPointModel
                End If
            End If
        End If
        If dictionaryListPoint.Count > 3 Then
            Dim positionPreparationPoint As Cad.Foundation.Vector3D = Nothing
            Dim polyBottomNozzleOrGrillage As DwgPolyline = New DwgPolyline
            polyBottomNozzleOrGrillage.Add(New BugleVector2D)
            polyBottomNozzleOrGrillage.Add(New BugleVector2D)
            polyBottomNozzleOrGrillage.Add(New BugleVector2D)
            polyBottomNozzleOrGrillage.Add(New BugleVector2D)
            Dim poly3dBottomNozzleOrGrillage As DwgPolyline3D = New DwgPolyline3D
            poly3dBottomNozzleOrGrillage.Add(New Cad.Foundation.Vector3D)
            poly3dBottomNozzleOrGrillage.Add(New Cad.Foundation.Vector3D)
            poly3dBottomNozzleOrGrillage.Add(New Cad.Foundation.Vector3D)
            poly3dBottomNozzleOrGrillage.Add(New Cad.Foundation.Vector3D)
            For i As Integer = 0 To dictionaryListPoint.Count - 1
                Dim pointGrillage As PointStructure = dictionaryListPoint.ElementAt(i).Value
                positionPreparationPoint = New Cad.Foundation.Vector3D(pointGrillage.X - pointGrillage.dx, pointGrillage.Y - pointGrillage.dy, pointGrillage.Z + pointGrillage.dz)
                If pointGrillage.Code Like "leftPt1" Then
                    polyBottomNozzleOrGrillage.Item(0) = New BugleVector2D(positionPreparationPoint.Pos, 0)
                    poly3dBottomNozzleOrGrillage.Item(0) = positionPreparationPoint
                ElseIf pointGrillage.Code Like "leftPt2" Then
                    polyBottomNozzleOrGrillage.Item(1) = New BugleVector2D(positionPreparationPoint.Pos, 0)
                    poly3dBottomNozzleOrGrillage.Item(1) = positionPreparationPoint
                ElseIf pointGrillage.Code Like "rightPt2" Then
                    polyBottomNozzleOrGrillage.Item(2) = New BugleVector2D(positionPreparationPoint.Pos, 0)
                    poly3dBottomNozzleOrGrillage.Item(2) = positionPreparationPoint
                ElseIf pointGrillage.Code Like "rightPt1" Then
                    polyBottomNozzleOrGrillage.Item(3) = New BugleVector2D(positionPreparationPoint.Pos, 0)
                    poly3dBottomNozzleOrGrillage.Item(3) = positionPreparationPoint
                End If
            Next i
            If polyBottomNozzleOrGrillage.Count <> 4 Then Return False
            polyBottomNozzleOrGrillage.Closed = True
            Dim polyTopPreparation As DwgPolyline = New DwgPolyline
            Dim listEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            polyBottomNozzleOrGrillage.Offset(listEnt, OffsetLengthTop)
            If listEnt.Count > 0 Then
                polyTopPreparation = listEnt(0)
                If polyTopPreparation.Area < polyBottomNozzleOrGrillage.Area Then
                    listEnt = New List(Of DwgEntity)
                    polyBottomNozzleOrGrillage.Offset(listEnt, -1 * OffsetLengthTop)
                    polyTopPreparation = listEnt(0)
                End If
            End If
            If polyTopPreparation.Count = 5 Then
                polyTopPreparation.RemoveAt(4)
            End If
            Dim polyBottomPreparation As DwgPolyline = New DwgPolyline
            listEnt = New List(Of DwgEntity)
            Dim deltaL As Double = OffsetLengthBottom - OffsetLengthTop
            polyBottomNozzleOrGrillage.Offset(listEnt, OffsetLengthBottom)
            If listEnt.Count > 0 Then
                polyBottomPreparation = listEnt(0)
                If polyBottomPreparation.Area < polyBottomNozzleOrGrillage.Area Then
                    listEnt = New List(Of DwgEntity)
                    polyBottomNozzleOrGrillage.Offset(listEnt, -1 * OffsetLengthBottom)
                    polyBottomPreparation = listEnt(0)
                End If
            End If
            If polyBottomPreparation.Count = 5 Then
                polyBottomPreparation.RemoveAt(4)
            End If
            If polyTopPreparation.Count <> 4 Then Return False
            If polyBottomPreparation.Count <> 4 Then Return False
            polyTopPreparation.Closed = True
            polyBottomPreparation.Closed = True

            Dim listModelPoint As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
            Dim x As Double = Math.Round(polyTopPreparation.Item(0).Vertex.X, 3)
            Dim y As Double = Math.Round(polyTopPreparation.Item(0).Vertex.Y, 3)
            Dim z As Double = Math.Round(poly3dBottomNozzleOrGrillage.Item(0).Z, 3)
            Dim dx As Double = Math.Round(polyTopPreparation.Item(0).Vertex.X - polyBottomPreparation.Item(0).Vertex.X, 3)
            Dim dy As Double = Math.Round(polyTopPreparation.Item(0).Vertex.Y - polyBottomPreparation.Item(0).Vertex.Y, 3)
            Dim dz As Double = Math.Round(Height, 3)
            Dim code As String = "leftPt1"
            Dim pointModel As PointStructure = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(1, pointModel)

            x = Math.Round(polyTopPreparation.Item(1).Vertex.X, 3)
            y = Math.Round(polyTopPreparation.Item(1).Vertex.Y, 3)
            z = Math.Round(poly3dBottomNozzleOrGrillage.Item(1).Z, 3)
            dx = Math.Round(polyTopPreparation.Item(1).Vertex.X - polyBottomPreparation.Item(1).Vertex.X, 3)
            dy = Math.Round(polyTopPreparation.Item(1).Vertex.Y - polyBottomPreparation.Item(1).Vertex.Y, 3)
            dz = Math.Round(Height, 3)
            code = "leftPt2"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(2, pointModel)

            x = Math.Round(polyTopPreparation.Item(2).Vertex.X, 3)
            y = Math.Round(polyTopPreparation.Item(2).Vertex.Y, 3)
            z = Math.Round(poly3dBottomNozzleOrGrillage.Item(2).Z, 3)
            dx = Math.Round(polyTopPreparation.Item(2).Vertex.X - polyBottomPreparation.Item(2).Vertex.X, 3)
            dy = Math.Round(polyTopPreparation.Item(2).Vertex.Y - polyBottomPreparation.Item(2).Vertex.Y, 3)
            dz = Math.Round(Height, 3)
            code = "rightPt2"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(3, pointModel)

            x = Math.Round(polyTopPreparation.Item(3).Vertex.X, 3)
            y = Math.Round(polyTopPreparation.Item(3).Vertex.Y, 3)
            z = Math.Round(poly3dBottomNozzleOrGrillage.Item(3).Z, 3)
            dx = Math.Round(polyTopPreparation.Item(3).Vertex.X - polyBottomPreparation.Item(3).Vertex.X, 3)
            dy = Math.Round(polyTopPreparation.Item(3).Vertex.Y - polyBottomPreparation.Item(3).Vertex.Y, 3)
            dz = Math.Round(Height, 3)
            code = "rightPt1"
            pointModel = New PointStructure(x, y, z, dx, dy, -1 * dz, 0, code)
            listModelPoint.Add(4, pointModel)
            _elementBridgePoint.ListPointModel = listModelPoint
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'вычисляем длину
            Dim leftPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
            Dim leftPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
            Dim rightPoint1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
            Dim rightPoint2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
            If listModelPoint.Count > 1 Then
                For i As Integer = 0 To listModelPoint.Count - 1
                    Dim tempPoint As PointStructure = listModelPoint.ElementAt(i).Value
                    If tempPoint.Code Like "leftPt1" Then
                        leftPoint1 = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    ElseIf tempPoint.Code Like "leftPt2" Then
                        leftPoint2 = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    ElseIf tempPoint.Code Like "rightPt1" Then
                        rightPoint1 = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    ElseIf tempPoint.Code Like "rightPt2" Then
                        rightPoint2 = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    End If
                Next
            End If
            Dim lenghtPreparation1 As Double = (leftPoint1 - leftPoint2).Length
            Dim lenghtPreparation2 As Double = (leftPoint1 - rightPoint1).Length
            TopLength = Math.Round(lenghtPreparation1, 3)
            BottomLength = Math.Round(lenghtPreparation1 - 2 * OffsetLengthTop + 2 * OffsetLengthBottom, 3)
            TopWidth = Math.Round(lenghtPreparation2, 3)
            BottomWidth = Math.Round(lenghtPreparation2 - 2 * OffsetLengthTop + 2 * OffsetLengthBottom, 3)
            TopElevation = Math.Round(leftPoint1.Z, 3)
            BottomElevation = Math.Round(leftPoint1.Z - Height, 3)
            'вычисляем осевую линию
            Dim middlePt1 As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(leftPoint1, rightPoint1)
            Dim middlePt2 As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(leftPoint2, rightPoint2)
            _elementBridgePoint.StartAxisPoint = middlePt1
            _elementBridgePoint.EndAxisPoint = middlePt2
            Dim centerAxisPoint As Vector2D = New Vector2D(-1, -1)
            If IsNothing(align) = False Then
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(align, middlePt1.Pos, middlePt2.Pos)
                If pointIntersectCollection.Count > 0 Then
                    centerAxisPoint = pointIntersectCollection.ElementAt(0)
                End If
            End If
            If centerAxisPoint.X = -1 And centerAxisPoint.Y = -1 Then
                centerAxisPoint = MathFunction.funcCalcMiddleCoordByToPoints2d(middlePt1.Pos, middlePt2.Pos)
            End If
            _elementBridgePoint.CenterTopPoint = New Vector3D(centerAxisPoint, TopElevation)
        End If
        Return True
    End Function
    'рисование оси подготовки
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLinePreparation As DwgLine = Nothing
        Dim dataStructurePreparation As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxisPreparation As List(Of StructureElement) = Pillar.getElementPillar(StructureElement.typeObject.axisPreparation, dictionaryBridgeElements, NumberPillar, 0, 0, Number)
        If listAxisPreparation.Count = 0 Then
            dataStructurePreparation = createAxisPreparationPillar(idBridge)
        ElseIf listAxisPreparation.Count = 1 Then
            dataStructurePreparation = listAxisPreparation.Item(0)
        Else
            dataStructurePreparation = StructureElement.isValidateDataStructure(listAxisPreparation)
        End If
        If IsNothing(dataStructurePreparation) Then Return Nothing
        axisLinePreparation = dataStructurePreparation.DWGEntity
        If IsNothing(axisLinePreparation) = True Then Return Nothing
        If axisLinePreparation.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Подготовка (ось)")
            styleAxisNozzle.setObjectStyle(axisLinePreparation)
        End If
        'ось насадки
        axisLinePreparation.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLinePreparation.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLinePreparation.StartPoint.Pos - axisLinePreparation.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось подготовки имеет нулевое значение.")
            Return dataStructurePreparation
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLinePreparation) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLinePreparation)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructurePreparation.KeyParameter = strJson
        dataStructurePreparation.DWGEntity = axisLinePreparation
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLinePreparation, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructurePreparation)
        Return dataStructurePreparation
    End Function
    'функция возвращает все точки подготовки в виде словаря 
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


End Class
