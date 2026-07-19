Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports CivilEnginStructures.Pillar
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class PostcardPillar
    ' Приватные поля класса
    Private _numberPillar As Integer                  ' Номер опоры
    Private _numberSubPillar As Integer               ' Номер подопоры
    Private _number As Integer                        ' Номер откосного крыла
    Private _length As Double                         ' Длина откосного крыла
    Private _width As Double                          ' Толщина откосного крыла
    Private _lengthCabinetWall As Double              ' Длина горизонтальной части у шкафной стенки
    Private _heightCabinetWall As Double              ' Высота открылка у шкафной стенки
    Private _heightEndNozzle As Double                ' Высота открылка у торца насадки
    Private _topElevation As Double
    Private _fixedLenght As Boolean
    Private _sideObject As Pillar.SidePillarElement                    ' Тип (Left/Right)
    Private _model As String                          ' Имя модели
    Public _elementBridgePoint As PointsCollections

    ' Конструктор класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberPillar = 0
        _numberSubPillar = 0
        _number = 0
        _length = 0.0
        _width = 0.0
        _lengthCabinetWall = 0.0
        _heightCabinetWall = 0.0
        _heightEndNozzle = 0.0
        _topElevation = 0
        _fixedLenght = False
        _sideObject = Pillar.SidePillarElement.None
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
    Public Property SidePostcard() As Pillar.SidePillarElement
        Get
            Return _sideObject
        End Get
        Set(value As Pillar.SidePillarElement)
            _sideObject = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Полная длина открылка")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    Public Property Length() As Double
        Get
            Return _length
        End Get
        Set(value As Double)
            If value >= 0 Then
                _length = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина (толщина) открылка")>
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
    <Description("Длина горизонтальной части примыкающего к шкафной стенке открылка, м")>
    <Category("Свойства")>
    <DisplayName("Длина горизонтальной части")>
    Public Property LengthCabinetWall() As Double
        Get
            Return _lengthCabinetWall
        End Get
        Set(value As Double)
            If value >= 0 Then
                _lengthCabinetWall = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Вытота стенки открылка примыкающего к шкафной стенке, м")>
    <Category("Свойства")>
    <DisplayName("Высота у шк. стенки")>
    Public Property HeightCabinetWall() As Double
        Get
            Return _heightCabinetWall
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightCabinetWall = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Вытота вертикальной стенки открылка у фасада насадки, м")>
    <Category("Свойства")>
    <DisplayName("Высота по фасаду насадки")>
    Public Property HeightEndNozzle() As Double
        Get
            Return _heightEndNozzle
        End Get
        Set(value As Double)
            If value >= 0 Then
                _heightEndNozzle = value
            End If
        End Set
    End Property

    <Browsable(True)>
    <Description("Высота оккрылка у шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Высота")>
    Public Property TopElevation() As Double
        Get
            Return _topElevation
        End Get
        Set(value As Double)
            _topElevation = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Фиксированная длина (до фасада насадки), м")>
    <Category("Свойства")>
    <DisplayName("Фиксированная длина")>
    Public Property FixedLenght() As Boolean
        Get
            Return _fixedLenght
        End Get
        Set(value As Boolean)
            If value >= 0 Then
                _fixedLenght = value
            End If
        End Set
    End Property
    Public Shared Function createAxisPostcard(ByVal idBridge As String, ByVal typeUserPostcard As Pillar.SidePillarElement) As StructureElement
        Dim elementPostcard As StructureElement = New StructureElement()
        elementPostcard.Label = "Мосты и путепроводы"
        If typeUserPostcard = Pillar.SidePillarElement.Left Then
            elementPostcard.ClassObject = StructureElement.classStructure.PostcardLeftPillar
            elementPostcard.Name = StructureElement.typeObject.axisLeftPostcard
            elementPostcard.Description = "Откосное крыло левое (ось)"
        ElseIf typeUserPostcard = Pillar.SidePillarElement.Right Then
            elementPostcard.ClassObject = StructureElement.classStructure.PostcardRightPillar
            elementPostcard.Name = StructureElement.typeObject.axisRightPostcard
            elementPostcard.Description = "Откосное крыло правое (ось)"
        End If
        elementPostcard.KeyParameter = ""
        elementPostcard.IdElement = Guid.NewGuid.ToString
        elementPostcard.IdStructure = idBridge
        elementPostcard.Note = ""
        elementPostcard.DWGEntity = New DwgLine()
        Return elementPostcard
    End Function
    'ищет обратный открылок
    Public Shared Function getPoctcardPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal typeUserPostcard As Pillar.SidePillarElement, Optional ByVal numberSubPillar As Integer = 0) As List(Of StructureElement)
        Dim dataPoctcard As New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        Dim listAxisPoctcard As List(Of StructureElement) = Nothing
        If typeUserPostcard = SidePillarElement.Left Then
            If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisLeftPostcard) = True Then
                listAxisPoctcard = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisLeftPostcard)
            End If
        ElseIf typeUserPostcard = SidePillarElement.Right Then
            If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisRightPostcard) = True Then
                listAxisPoctcard = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisRightPostcard)
            End If
        End If
        If IsNothing(listAxisPoctcard) = False Then
            If listAxisPoctcard.Count > 0 Then
                For k As Integer = 0 To listAxisPoctcard.Count - 1
                    Dim tempData As StructureElement = listAxisPoctcard.Item(k)
                    If IsNothing(tempData) = False Then
                        Dim userAxisPoctcard As PostcardPillar = tempData.getPostcardPillar
                        If IsNothing(userAxisPoctcard) = False Then
                            If numberPillar = userAxisPoctcard.NumberPillar Then
                                If numberSubPillar = 0 Then numberSubPillar = userAxisPoctcard.NumberSubPillar
                                If numberSubPillar = userAxisPoctcard.NumberSubPillar Then
                                    dataPoctcard.Add(tempData)
                                End If
                            End If
                        End If
                    End If
                Next k
            End If
        End If
        Return dataPoctcard
    End Function
    Public Shared Function readPropertiesPostcard(ByVal numbPillar As Integer, ByVal numbSubPillar As Integer, ByVal DGV_Postcard As DataGridView, Optional leftPostcard As Boolean = True) As PostcardPillar
        'рисуем обратный открылок
        Dim result As PostcardPillar = New PostcardPillar
        result.NumberPillar = numbPillar
        result.NumberSubPillar = numbSubPillar
        If leftPostcard = True Then
            result.SidePostcard = Pillar.SidePillarElement.Left
        Else
            result.SidePostcard = Pillar.SidePillarElement.Right
        End If
        If DGV_Postcard.RowCount > 1 Then
            For i As Integer = 0 To DGV_Postcard.RowCount - 1
                Dim tag As String = DGV_Postcard.Rows(i).Tag
                Dim value As String = DGV_Postcard.Rows(i).Cells(1).Value
                If IsNothing(tag) = False And IsNothing(value) = False Then
                    If tag Like "bridge_postcardleft_length" Then
                        If IsNumeric(value) = True Then
                            result.Length = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение длины левого открылка.")
                        End If
                    ElseIf tag Like "bridge_postcardright_length" Then
                        If IsNumeric(value) = True Then
                            result.Length = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение длины левого открылка.")
                        End If
                    ElseIf tag Like "bridge_postcardleft_thickness" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение ширины левого открылка.")
                        End If
                    ElseIf tag Like "bridge_postcardright_thickness" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.Width = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение ширины левого открылка.")
                        End If
                    ElseIf tag Like "bridge_postcardleft_lengthb" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LengthCabinetWall = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение длины левого открылка у шкафной стенки.")
                        End If
                    ElseIf tag Like "bridge_postcardright_lengthb" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.LengthCabinetWall = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение длины левого открылка у шкафной стенки.")
                        End If
                    ElseIf tag Like "bridge_postcardleft_heightс" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightCabinetWall = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение высоты левого открылка у шкафной стенки.")
                        End If
                    ElseIf tag Like "bridge_postcardright_heightс" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightCabinetWall = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение высоты левого открылка у шкафной стенки.")
                        End If
                    ElseIf tag Like "bridge_postcardleft_heightn" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightEndNozzle = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение высоты левого открылка у торца насадки.")
                        End If
                    ElseIf tag Like "bridge_postcardright_heightn" Then
                        If IsNumeric(value) = True And value > 0 Then
                            result.HeightEndNozzle = Math.Round(Val(value), 3)
                        Else
                            MsgBox("Некорректное значение высоты левого открылка у торца насадки.")
                        End If
                    End If
                End If
            Next i
        End If
        Return result
    End Function
    'запись расчетных данных в датогрид
    Public Function writeProjectData(ByRef DGV_Postcard As DataGridView) As Boolean
        If IsNothing(DGV_Postcard) = True Then Return False
        If DGV_Postcard.RowCount > 1 Then
            Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
            If ListPointModel.Count > 0 Then
                Dim boolTopElevation As Boolean = False
                Dim boolBottomElevation As Boolean = False
                Dim boolLenght As Boolean = False
                Dim boolStepRacks As Boolean = False
                For i As Integer = 0 To DGV_Postcard.RowCount - 1
                    Dim oldTag As String = DGV_Postcard.Rows(i).Tag
                    If oldTag Like "calc-TopElevation" Then
                        DGV_Postcard.Rows(i).Cells(1).Value = TopElevation
                        boolTopElevation = True
                    End If
                Next i
                If boolTopElevation = False Then
                    Dim numberRow As Integer = DGV_Postcard.RowCount - 1
                    DGV_Postcard.Rows.Insert(numberRow)
                    DGV_Postcard.Rows(numberRow).DefaultCellStyle.ForeColor = Color.Red
                    DGV_Postcard.Rows(numberRow).Cells(0).Value = "Отметка верха, м"
                    DGV_Postcard.Rows(numberRow).Cells(1).Value = TopElevation
                    DGV_Postcard.Rows(numberRow).Tag = "calc-TopElevation"
                End If
            End If
        End If
        Return True
    End Function

    'запись данных в датогрид
    Public Function writePropertiesPostcard(ByRef DGV_Postcard As DataGridView) As Boolean
        If IsNothing(DGV_Postcard) = True Then Return False
        If DGV_Postcard.RowCount > 1 Then
            For j As Integer = 0 To DGV_Postcard.RowCount - 1
                Dim tag As String = DGV_Postcard.Rows(j).Tag
                If IsNothing(tag) = False Then
                    If tag Like "bridge_postcardleft_length" Then
                        DGV_Postcard.Rows(j).Cells(1).Value = Length
                    ElseIf tag Like "bridge_postcardleft_thickness" Then
                        DGV_Postcard.Rows(j).Cells(1).Value = Width
                    ElseIf tag Like "bridge_postcardleft_lengthb" Then
                        DGV_Postcard.Rows(j).Cells(1).Value = LengthCabinetWall
                    ElseIf tag Like "bridge_postcardleft_heightс" Then
                        DGV_Postcard.Rows(j).Cells(1).Value = HeightCabinetWall
                    ElseIf tag Like "bridge_postcardleft_heightn" Then
                        DGV_Postcard.Rows(j).Cells(1).Value = HeightEndNozzle
                    End If
                End If
            Next j
        Else
            Return False
        End If
        Return True
    End Function
    'предварительный расчет
    Public Function calculatePostcard(ByVal userNozzle As NozzlePillar) As Boolean
        Dim pointModelNozzle As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)

        If IsNothing(userNozzle) = False Then
            If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
                pointModelNozzle = userNozzle._elementBridgePoint.ListPointModel
            End If
        End If
        If pointModelNozzle.Count = 0 Then Return False
        Dim shortLineNozzle As DwgLine = New DwgLine 'короткая сторона (ширина шкафной стенки)
        Dim longLineLeftNozzle As DwgLine = New DwgLine 'длинная сторона (сторона насадки передняя)
        Dim longLineRightNozzle As DwgLine = New DwgLine 'длинная сторона (сторона насадки задняя)
        'выбираем точки для левого открылка
        If SidePostcard = Pillar.SidePillarElement.Left Then
            'берем 4 точки насадки
            For i As Integer = 0 To pointModelNozzle.Count - 1
                Dim tempPoint As PointStructure = pointModelNozzle.ElementAt(i).Value
                If tempPoint.Code Like "middlePt1" Then
                    shortLineNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineRightNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "rightPt1" Then
                    shortLineNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineLeftNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "middlePt2" Then
                    longLineRightNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "rightPt2" Then
                    longLineLeftNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                End If
            Next
        ElseIf SidePostcard = Pillar.SidePillarElement.Right Then
            'берем 4 точки насадки
            For i As Integer = 0 To pointModelNozzle.Count - 1
                Dim tempPoint As PointStructure = pointModelNozzle.ElementAt(i).Value
                If tempPoint.Code Like "middlePt2" Then
                    shortLineNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineRightNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "rightPt2" Then
                    shortLineNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                    longLineLeftNozzle.StartPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "middlePt1" Then
                    longLineRightNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                ElseIf tempPoint.Code Like "rightPt1" Then
                    longLineLeftNozzle.EndPoint = New Cad.Foundation.Vector3D(tempPoint.X, tempPoint.Y, tempPoint.Z)
                End If
            Next
        End If
        'определяем точки откосного крыла
        '1.точка в начале шкафной стенки по низу
        Dim pointLeftBottom1 As Cad.Foundation.Vector3D = shortLineNozzle.StartPoint
        Dim pointRightBottom1 As Cad.Foundation.Vector3D = MathFunction.FuncCalcPointInLine(pointLeftBottom1, longLineRightNozzle.EndPoint, Width)
        'определаем высоту первой  точки по верху
        Dim pointLeftTop1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointLeftBottom1.Pos, pointLeftBottom1.Z + HeightCabinetWall)
        Dim pointRightTop1 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointRightBottom1.Pos, pointRightBottom1.Z + HeightCabinetWall)
        '2 точка середина открылка
        Dim pointLeft2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeftBottom1.Pos, shortLineNozzle.Rotation, LengthCabinetWall)
        Dim pointRight2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointRightBottom1.Pos, shortLineNozzle.Rotation, LengthCabinetWall)
        'определаем высоту второй точки по верху
        Dim pointLeftTop2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointLeft2, pointLeftBottom1.Z + HeightCabinetWall)
        Dim pointRightTop2 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointRight2, pointRightBottom1.Z + HeightCabinetWall)
        '3 крайняя точка 
        Dim bottomLenght As Double = Length
        If FixedLenght = True Then
            bottomLenght = userNozzle.Width - userNozzle.WidthPlateCabinetWall
        End If
        Dim pointLeft3 As Vector2D = New Vector2D
        Dim pointRight3 As Vector2D = New Vector2D
        'предварительные данные по 3 точке
        Dim pointLeftBottom3 As Cad.Foundation.Vector3D = shortLineNozzle.EndPoint
        Dim pointRightBottom3 As Cad.Foundation.Vector3D = MathFunction.FuncCalcPointInLine(shortLineNozzle.EndPoint, longLineLeftNozzle.EndPoint, Width)
        If bottomLenght > 0 Then
            pointLeft3 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointLeftBottom1.Pos, shortLineNozzle.Rotation, Length)
            pointRight3 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointRightBottom1.Pos, shortLineNozzle.Rotation, Length)
            pointLeftBottom3 = MathFunction.FuncCalcPointInLine(pointLeftBottom1, pointLeftBottom3, Length)
            pointRightBottom3 = MathFunction.FuncCalcPointInLine(pointRightBottom1, pointRightBottom3, Length)
        End If
        Dim pointLeftTop3 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointLeftBottom3.Pos, pointLeftBottom3.Z + HeightEndNozzle)
        Dim pointRightTop3 As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D(pointRightBottom3.Pos, pointRightBottom3.Z + HeightEndNozzle)
        'определаем положение 2 точки по низу
        Dim elevLeftPoint2 As Double = MathFunction.FuncCalcElevationByLine(pointLeftBottom1, pointLeftBottom3, pointLeft2)
        Dim elevRightPoint2 As Double = MathFunction.FuncCalcElevationByLine(pointRightBottom1, pointRightBottom3, pointRight2)
        '=================================================================================================================================================
        'заполняем словарь
        Dim pointBridgeHand As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        Dim countLeft As Integer = 1
        Dim x As Double = Math.Round(pointRightTop1.X, 3)
        Dim y As Double = Math.Round(pointRightTop1.Y, 3)
        Dim z As Double = Math.Round(pointRightTop1.Z, 3)
        Dim h1 As Double = Math.Round(HeightCabinetWall, 3)
        Dim code As String = "rightPt1"
        Dim pointBridge As PointStructure = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1

        x = Math.Round(pointRight2.X, 3)
        y = Math.Round(pointRight2.Y, 3)
        z = Math.Round(pointRightTop1.Z, 3)
        h1 = Math.Round(pointRightTop1.Z - elevRightPoint2, 3)
        code = "rightPt0"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1

        x = Math.Round(pointRightTop3.X, 3)
        y = Math.Round(pointRightTop3.Y, 3)
        z = Math.Round(pointRightTop3.Z, 3)
        h1 = Math.Round(HeightEndNozzle, 3)
        code = "rightPt2"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'в противоположном направлении
        x = Math.Round(pointLeftTop3.X, 3)
        y = Math.Round(pointLeftTop3.Y, 3)
        z = Math.Round(pointLeftTop3.Z, 3)
        h1 = Math.Round(HeightEndNozzle, 3)
        code = "leftPt2"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1

        x = Math.Round(pointLeft2.X, 3)
        y = Math.Round(pointLeft2.Y, 3)
        z = Math.Round(pointLeftTop1.Z, 3)
        h1 = Math.Round(pointLeftTop1.Z - elevLeftPoint2, 3)
        code = "leftPt0"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        countLeft += 1

        x = Math.Round(pointLeftTop1.X, 3)
        y = Math.Round(pointLeftTop1.Y, 3)
        z = Math.Round(pointLeftTop1.Z, 3)
        h1 = Math.Round(HeightCabinetWall, 3)
        code = "leftPt1"
        pointBridge = New PointStructure(x, y, z, 0, 0, -1 * h1, 0, code)
        pointBridgeHand.Add(countLeft, pointBridge)
        _elementBridgePoint.ListPointModel = pointBridgeHand
        '=============================================================================================================
        Dim middleStartPoint As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(pointLeftTop1, pointRightTop1)
        Dim middleEndPoint As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(pointLeftTop3, pointRightTop3)
        _elementBridgePoint.StartAxisPoint = middleStartPoint
        _elementBridgePoint.EndAxisPoint = middleEndPoint
        TopElevation = Math.Round(pointLeftTop1.Z, 3)
        _elementBridgePoint.CenterTopPoint = middleStartPoint
        Return True
    End Function
    'рисование оси откосного крыла
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
        Dim axisLinePostcard As DwgLine = Nothing
        Dim dataStructurePostcard As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim listAxis As List(Of StructureElement) = getPoctcardPillar(dictionaryBridgeElements, NumberPillar, SidePostcard, NumberSubPillar)
        If listAxis.Count = 0 Then
            dataStructurePostcard = createAxisPostcard(idBridge, SidePostcard)
        ElseIf listAxis.Count = 1 Then
            dataStructurePostcard = listAxis.Item(0)
        Else
            dataStructurePostcard = StructureElement.isValidateDataStructure(listAxis)
        End If
        If IsNothing(dataStructurePostcard) Then Return Nothing
        axisLinePostcard = dataStructurePostcard.DWGEntity
        If IsNothing(axisLinePostcard) = True Then axisLinePostcard = New DwgLine
        If axisLinePostcard.Length = 0 Then
            Dim layerAxisNozzle As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorAxisNozzle As CadColor = New CadColor(7)
            Dim nameTypeLineAxisNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineAxisNozzle As Integer = 1
            Dim widthTypeLineAxisNozzle As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleAxisNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Откосное крыло (ось)")
            styleAxisNozzle.setObjectStyle(axisLinePostcard)
        End If
        'ось обратного открылка
        axisLinePostcard.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLinePostcard.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxisNozzle As Double = (axisLinePostcard.StartPoint.Pos - axisLinePostcard.EndPoint.Pos).Length
        If lenghtAxisNozzle = 0 Then
            MsgBox("Ось откосного крыла имеет нулевое значение. ")
            Return dataStructurePostcard
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLinePostcard) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLinePostcard)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataStructurePostcard.KeyParameter = strJson
        dataStructurePostcard.DWGEntity = axisLinePostcard
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLinePostcard, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructurePostcard)
        Return dataStructurePostcard
    End Function
    'функция возвращает все точки насадки в виде словаря 
    Public Function getProjectionPoint(ByVal matrixTransform As TransformCivilStructure) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim transformTop As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        Dim transformBottom As Dictionary(Of String, ProjectionPoint) = New Dictionary(Of String, ProjectionPoint)
        If IsNothing(_elementBridgePoint.ListPointModel) = True Then
            Return result
        End If
        If _elementBridgePoint.ListPointModel.Count = 0 Then
            Return result
        End If
        'записываем результат
        Dim ListPointModel As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
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
