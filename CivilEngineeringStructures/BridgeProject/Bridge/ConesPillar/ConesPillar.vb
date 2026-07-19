Imports System.ComponentModel
Imports CivilEnginStructures.Bridges
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization.Geometry

Public Class ConesPillar
    ' Поля класса
    Private _numberPillar As String             'Номер конуса (1-первая опора, 2- последняя опора
    Private _lenghtLeftTop As Double            'длина конуса до начала откосного крыла слева
    Private _lenghtRightTop As Double           'длина конуса до начала откосного крыла справа
    Private _offsetLeftHand As Double
    Private _offsetRightHand As Double
    Private _lenghtLeftHand As Double
    Private _lenghtRightHand As Double
    Private _topRadius As Double
    Private _bottomRadius As Double
    Private _verticalOffsetNozzle As Double
    Private _slope As Double
    Private _horizontalOffsetPlate As Double
    Private _elevationLeftEgeStart As Double
    Private _elevationRightEgeStart As Double
    Private _elevationLeftMiddleHand As Double
    Private _elevationRightMiddleHand As Double
    Private _elevationGround As Double
    Private _listCounters As List(Of DwgPolyline3D)                   'список линий
    Public Sub New()
        _numberPillar = 0
        _listCounters = New List(Of DwgPolyline3D)
    End Sub

    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    Public Property NumberPillar() As Integer
        Get
            Return _numberPillar
        End Get
        Set(value As Integer)
            _numberPillar = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Расстояние от начала конуса до края откосного крыла слева, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние слева до края крыла")>
    Public Property LenghtLeftTop() As Double
        Get
            Return _lenghtLeftTop
        End Get
        Set(value As Double)
            _lenghtLeftTop = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Расстояние от начала конуса до края откосного крыла справа, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние справа до края крыла")>
    Public Property LenghtRightTop() As Double
        Get
            Return _lenghtRightTop
        End Get
        Set(value As Double)
            _lenghtRightTop = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Выпуск конуса в левое откосное крыло, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск на левое крыло")>
    Public Property LenghtLeftHand() As Double
        Get
            Return _lenghtLeftHand
        End Get
        Set(value As Double)
            _lenghtLeftHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Выпуск конуса в правое откосное крыло, м")>
    <Category("Свойства")>
    <DisplayName("Выпуск на правое крыло")>
    Public Property LenghtRightHand() As Double
        Get
            Return _lenghtRightHand
        End Get
        Set(value As Double)
            _lenghtRightHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение по горизонтали от левого открылка")>
    <Category("Свойства")>
    <DisplayName("Смещение влево")>
    Public Property OffsetLeftHand() As Double
        Get
            Return _offsetLeftHand
        End Get
        Set(value As Double)
            _offsetLeftHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение по горизонтали от правого открылка")>
    <Category("Свойства")>
    <DisplayName("Смещение вправо")>
    Public Property OffsetRightHand() As Double
        Get
            Return _offsetRightHand
        End Get
        Set(value As Double)
            _offsetRightHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Вертикальное смещение края конуса у насадки, м")>
    <Category("Свойства")>
    <DisplayName("Вертикальное смещение у насадки")>
    Public Property VerticalOffsetNozzle() As Double
        Get
            Return _verticalOffsetNozzle
        End Get
        Set(value As Double)
            _verticalOffsetNozzle = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Ширина горизонтальной площадки у насадки")>
    <Category("Свойства")>
    <DisplayName("Ширина площадки")>
    Public Property HorizontalOffsetPlate() As Double
        Get
            Return _horizontalOffsetPlate
        End Get
        Set(value As Double)
            _horizontalOffsetPlate = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Заложение откоса")>
    <Category("Свойства")>
    <DisplayName("Заложение откоса")>
    Public Property Slope() As Double
        Get
            Return _slope
        End Get
        Set(value As Double)
            _slope = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Радиус закругления основания конуса, м")>
    <Category("Свойства")>
    <DisplayName("Радиус закругления основания конуса")>
    Public Property BottomRadius() As Double
        Get
            Return _bottomRadius
        End Get
        Set(value As Double)
            _bottomRadius = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка обочины в начале конуса слева, м")>
    <Category("Свойства")>
    <DisplayName("Отметка обочины в начале конуса слева")>
    Public Property ElevationLeftEgeStart() As Double
        Get
            Return _elevationLeftEgeStart
        End Get
        Set(value As Double)
            _elevationLeftEgeStart = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка обочины в начале конуса справа, м")>
    <Category("Свойства")>
    <DisplayName("Отметка обочины в начале конуса справа")>
    Public Property ElevationRightEgeStart() As Double
        Get
            Return _elevationRightEgeStart
        End Get
        Set(value As Double)
            _elevationRightEgeStart = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка обочины в конце закругления конуса у откосного крыла слева, м")>
    <Category("Свойства")>
    <DisplayName("Отметка обочины конца закругления слева")>
    Public Property ElevationLeftMiddleHand() As Double
        Get
            Return _elevationLeftMiddleHand
        End Get
        Set(value As Double)
            _elevationLeftMiddleHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка обочины в конце закругления конуса у откосного крыла справа, м")>
    <Category("Свойства")>
    <DisplayName("Отметка обочины конца закругления слева")>
    Public Property ElevationRightMiddleHand() As Double
        Get
            Return _elevationRightMiddleHand
        End Get
        Set(value As Double)
            _elevationRightMiddleHand = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка земли основания конуса")>
    <Category("Свойства")>
    <DisplayName("Отметка земли")>
    Public Property ElevationGround() As Double
        Get
            Return _elevationGround
        End Get
        Set(value As Double)
            _elevationGround = value
        End Set
    End Property
    Public Function calulateFirstLine(ByVal userNozzle As NozzlePillar, ByVal userLeftHand As HandPillar, ByVal userRightHand As HandPillar, Optional ByVal countArcSegments As Integer = 5, Optional projectSurface As Surface = Nothing, Optional egSurface As Surface = Nothing) As List(Of DwgPolyline3D)
        Dim result As List(Of DwgPolyline3D) = New List(Of DwgPolyline3D)
        If IsNothing(userNozzle) = True Then
            Return result
        End If
        If IsNothing(userLeftHand) = True Then
            Return result
        End If
        If IsNothing(userRightHand) = True Then
            Return result
        End If
        If userNozzle._elementBridgePoint.ListPointModel.Count = 0 Then
            Return result
        End If
        Dim poly3d1 As DwgPolyline3D = New DwgPolyline3D 'верх откоса
        Dim poly3d2 As DwgPolyline3D = New DwgPolyline3D 'середина
        Dim poly3d3 As DwgPolyline3D = New DwgPolyline3D 'середина горизонтальная площадка
        Dim poly3d4 As DwgPolyline3D = New DwgPolyline3D 'земля
        'точки верхние насадки
        Dim rightNozzlePt1 As Cad.Foundation.Vector3D = userNozzle.getPointByCode("rightPt1", True)
        Dim rightNozzlePt2 As Cad.Foundation.Vector3D = userNozzle.getPointByCode("rightPt2", True)
        If (rightNozzlePt1.Pos - rightNozzlePt2.Pos).Length < 0.1 Then
            Return result
        End If
        'крайние точки откосного крыла левого
        If userLeftHand._elementBridgePoint.ListPointModel.Count = 0 Then
            Return result
        End If
        Dim leftHandPt1 As Cad.Foundation.Vector3D = userLeftHand.getPointByCode("leftPt1")
        Dim leftHandPt2 As Cad.Foundation.Vector3D = userLeftHand.getPointByCode("leftPt2")
        If (leftHandPt1.Pos - leftHandPt2.Pos).Length < 0.1 Then
            Return result
        End If
        'внешняя сторона откосного крыла левого
        Dim line1 As DwgLine = New DwgLine()
        line1.StartPoint = leftHandPt1
        line1.EndPoint = leftHandPt2
        Dim rotationLeftLine As Double = line1.Rotation 'от насадки
        Dim reverseRotationLeftLine As Double = MathFunction.reverseAngle(rotationLeftLine) 'к насадке
        Dim normalLeftAngle As Double = MathFunction.normalAngle(rotationLeftLine)
        Dim reverseNormalLeftAngle As Double = MathFunction.reverseAngle(normalLeftAngle)
        'крайние точки откосного крыла правого
        If userRightHand._elementBridgePoint.ListPointModel.Count = 0 Then
            Return result
        End If
        Dim rightHandPt1 As Cad.Foundation.Vector3D = userRightHand.getPointByCode("leftPt1")
        Dim rightHandPt2 As Cad.Foundation.Vector3D = userRightHand.getPointByCode("leftPt2")
        If (rightHandPt1.Pos - rightHandPt2.Pos).Length < 0.1 Then
            Return result
        End If
        'внешняя сторона откосного крыла правого
        Dim line2 As DwgLine = New DwgLine()
        line2.StartPoint = rightHandPt1
        line2.EndPoint = rightHandPt2
        Dim rotationRightLine As Double = line2.Rotation 'от насадки
        Dim reverseRotationRightLine As Double = MathFunction.reverseAngle(rotationRightLine) 'к насадке
        Dim normalRightAngle As Double = MathFunction.normalAngle(rotationRightLine)
        Dim reverseNormalRightAngle As Double = MathFunction.reverseAngle(normalRightAngle)
        '=============================================================================================================================
        'левое уширение
        'смещаем влево крайнюю точку левого откосного крыла
        Dim offsetPointPt2LeftHand As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftHandPt2.Pos, reverseNormalLeftAngle, OffsetLeftHand)
        'вычисляем точку начала конуса
        Dim pointStartLeftTopCole As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(offsetPointPt2LeftHand, rotationLeftLine, LenghtLeftTop)
        'вычисляем точки конца закругления на откосном круле
        Dim pointEndLeftTopCole As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftHandPt2.Pos, reverseRotationLeftLine, LenghtLeftHand)
        'вычисляем вспомогательную точку
        Dim TempPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointEndLeftTopCole, normalLeftAngle, 1)
        'вписываем окружность влево
        Dim listLeftTopPoint As List(Of Vector2D) = MathFunction.InscribeCircleBetweenSegments(pointStartLeftTopCole, offsetPointPt2LeftHand, TempPoint, pointEndLeftTopCole, OffsetLeftHand, countArcSegments)
        '=============================================================================================================================
        'правое уширение
        'делаем офсет внешней стороны крыла левого
        Dim offsetPointPt2RightHand As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightHandPt2, normalRightAngle, OffsetRightHand)
        'вычисляем точку начала конуса
        Dim pointStartRightTopCole As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(offsetPointPt2RightHand, rotationRightLine, LenghtRightTop)
        'вычисляем точки конца закругления на откосном круле
        Dim pointEndRightTopCole As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightHandPt2, reverseRotationRightLine, OffsetRightHand)
        'вспомогательная точка
        TempPoint = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointEndRightTopCole, reverseNormalRightAngle, 1)
        'вписываем окружность вправо
        Dim listRightTopPoint As List(Of Vector2D) = MathFunction.InscribeCircleBetweenSegments(TempPoint, pointEndRightTopCole, pointStartRightTopCole, offsetPointPt2RightHand, OffsetRightHand, countArcSegments)
        'определяем высоты найденных точек
        Dim elevPointLeft1 As Double = ElevationLeftEgeStart
        Try
            elevPointLeft1 = projectSurface.GetElevation(pointStartLeftTopCole)
        Catch ex As System.InvalidOperationException
        End Try
        Dim elevPointRight1 As Double = ElevationRightEgeStart
        Try
            elevPointLeft1 = projectSurface.GetElevation(pointStartRightTopCole)
        Catch ex As System.InvalidOperationException
        End Try
        '2 точка
        Dim elevPointLeft2 As Double = ElevationLeftMiddleHand
        Try
            elevPointLeft2 = projectSurface.GetElevation(offsetPointPt2LeftHand)
        Catch ex As System.InvalidOperationException
        End Try
        Dim elevPointRight2 As Double = ElevationRightMiddleHand
        Try
            elevPointRight2 = projectSurface.GetElevation(offsetPointPt2RightHand)
        Catch ex As System.InvalidOperationException
        End Try
        '3. ищем точку на открылке
        Dim elevPointLeft3 As Double = ElevationLeftMiddleHand
        Try
            elevPointLeft3 = projectSurface.GetElevation(pointEndLeftTopCole)
        Catch ex As System.InvalidOperationException
        End Try
        Dim elevPointRight3 As Double = ElevationRightMiddleHand
        Try
            elevPointRight3 = projectSurface.GetElevation(pointEndRightTopCole)
        Catch ex As System.InvalidOperationException
        End Try
        '4. ищем точку на насадке
        Dim pointLeft4 As Vector2D = rightNozzlePt1.Pos
        Dim elevPointLeft4 As Double = rightNozzlePt1.Z - VerticalOffsetNozzle
        Dim pointRight4 As Vector2D = rightNozzlePt2.Pos
        Dim elevPointRight4 As Double = rightNozzlePt2.Z - VerticalOffsetNozzle
        'апроксимируем дугу
        poly3d1.Add(New Vector3D(pointStartLeftTopCole, elevPointLeft1))
        If listLeftTopPoint.Count > 0 Then
            For k As Integer = 0 To listLeftTopPoint.Count - 1
                If k = 0 Then
                    Dim stDist As Double = (listLeftTopPoint.Item(0) - offsetPointPt2LeftHand).Length
                    If stDist > 0.01 Then
                        'poly3d1.Add(New Vector3D(offsetPointPt2LeftHand, elevPointLeft2))
                        poly3d1.Add(New Vector3D(listLeftTopPoint.Item(k), elevPointLeft2))
                        offsetPointPt2LeftHand = listLeftTopPoint.Item(k)
                    Else
                        poly3d1.Add(New Vector3D(offsetPointPt2LeftHand, elevPointLeft2))
                    End If
                ElseIf k = listLeftTopPoint.Count - 1 Then
                    Dim stDist As Double = (listLeftTopPoint.Item(k) - pointEndLeftTopCole).Length
                    If stDist > 0.01 Then
                        poly3d1.Add(New Vector3D(listLeftTopPoint.Item(k), elevPointLeft2))
                        'poly3d1.Add(New Vector3D(pointEndLeftTopCole, elevPointLeft2))
                        pointEndLeftTopCole = listLeftTopPoint.Item(k)
                    Else
                        poly3d1.Add(New Vector3D(pointEndLeftTopCole, elevPointLeft2))
                    End If
                Else
                    poly3d1.Add(New Vector3D(listLeftTopPoint.Item(k), elevPointLeft2))
                End If
            Next k
        End If
        'записываем точки насадки
        poly3d1.Add(New Vector3D(pointLeft4, elevPointLeft4))
        poly3d1.Add(New Vector3D(pointRight4, elevPointRight4))
        'апроксимируем дугу
        If listRightTopPoint.Count > 0 Then
            For k As Integer = 0 To listRightTopPoint.Count - 1
                If k = 0 Then
                    Dim stDist As Double = (listRightTopPoint.Item(0) - pointEndRightTopCole).Length
                    If stDist > 0.01 Then
                        'poly3d1.Add(New Vector3D(pointEndRightTopCole, elevPointRight2))
                        poly3d1.Add(New Vector3D(listRightTopPoint.Item(k), elevPointRight2))
                        pointEndRightTopCole = listRightTopPoint.Item(k)
                    Else
                        poly3d1.Add(New Vector3D(pointEndRightTopCole, elevPointLeft2))
                    End If
                ElseIf k = listRightTopPoint.Count - 1 Then
                    Dim stDist As Double = (listRightTopPoint.Item(k) - offsetPointPt2RightHand).Length
                    If stDist > 0.01 Then
                        poly3d1.Add(New Vector3D(listRightTopPoint.Item(k), elevPointRight2))
                        offsetPointPt2RightHand = listRightTopPoint.Item(k)
                        'poly3d1.Add(New Vector3D(offsetPointPt2RightHand, elevPointRight2))
                    Else
                        poly3d1.Add(New Vector3D(offsetPointPt2RightHand, elevPointRight2))
                    End If
                Else
                    poly3d1.Add(New Vector3D(listRightTopPoint.Item(k), elevPointRight2))
                End If
            Next k
        End If
        poly3d1.Add(New Vector3D(pointStartRightTopCole, elevPointRight1))
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'средняя линия
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim lenghtSlopeLeftCone As Double = (elevPointLeft2 - ElevationGround) * Slope
        Dim lenghtSlopeRightCone As Double = (elevPointRight2 - ElevationGround) * Slope
        'вычисляем пересечение линии насадки и офсетных линий обратного открылка
        Dim pointLastLeftGround As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointStartLeftTopCole, reverseNormalLeftAngle, lenghtSlopeLeftCone)
        Dim pointLastRightGround As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(pointStartRightTopCole, normalRightAngle, lenghtSlopeRightCone)
        Dim pointLastLeftGround3d As Vector3D = New Vector3D(pointLastLeftGround, ElevationGround)
        Dim pointLastRightGround3d As Vector3D = New Vector3D(pointLastRightGround, ElevationGround)
        Dim pointLastLeftTopCone3d As Vector3D = New Vector3D(pointStartLeftTopCole, elevPointLeft1)
        Dim pointLastRightTopCone3d As Vector3D = New Vector3D(pointStartRightTopCole, elevPointLeft1)
        'левый откос (вичисляем возицию точки с высотой насадки)
        Dim newLineLeft As DwgLine = New DwgLine()
        newLineLeft.StartPoint = pointLastLeftTopCone3d
        newLineLeft.EndPoint = pointLastLeftGround3d
        Dim middleLeftPoint1 As Vector2D = MathFunction.FuncCalcPositionLineByElevation(newLineLeft.StartPoint, newLineLeft.EndPoint, elevPointLeft4)
        Dim l1 As Double = (pointStartLeftTopCole - offsetPointPt2LeftHand).Length
        Dim middleLeftPoint2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleLeftPoint1, reverseRotationLeftLine, l1)

        Dim newLineRight As DwgLine = New DwgLine()
        newLineRight.StartPoint = pointLastRightTopCone3d
        newLineRight.EndPoint = pointLastRightGround3d
        Dim middleRightPoint1 As Vector2D = MathFunction.FuncCalcPositionLineByElevation(newLineRight.StartPoint, newLineRight.EndPoint, elevPointRight4)
        Dim l2 As Double = (pointStartRightTopCole - offsetPointPt2RightHand).Length
        Dim middleRightPoint2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleRightPoint1, reverseRotationRightLine, l2)

        Dim lineNozzle As DwgLine = New DwgLine
        lineNozzle.StartPoint = rightNozzlePt1
        lineNozzle.EndPoint = rightNozzlePt2
        Dim boolExtend As Boolean = BridgeGeometry.extendBeam(lineNozzle, HorizontalOffsetPlate, HorizontalOffsetPlate)
        Dim normalAngleNozzle As Double = MathFunction.normalAngle(lineNozzle.Rotation)
        Dim middleLeftNozzle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineNozzle.StartPoint.Pos, normalAngleNozzle, HorizontalOffsetPlate)
        Dim middleRightNozzle As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineNozzle.EndPoint.Pos, normalAngleNozzle, HorizontalOffsetPlate)
        Dim r1 As Double = (middleLeftPoint1 - pointStartLeftTopCole).Length
        Dim listMiddleTopPoint1 As List(Of Vector2D) = MathFunction.InscribeCircleBetweenSegments(middleLeftPoint1, middleLeftPoint2, middleRightNozzle, middleLeftNozzle, r1, countArcSegments)
        Dim r2 As Double = (middleRightPoint1 - pointStartRightTopCole).Length
        Dim listMiddleTopPoint2 As List(Of Vector2D) = MathFunction.InscribeCircleBetweenSegments(middleLeftNozzle, middleRightNozzle, middleRightPoint1, middleRightPoint2, r2, countArcSegments)
        'добавляем первую точку
        poly3d2.Add(New Vector3D(middleLeftPoint1, elevPointLeft4))
        If listMiddleTopPoint1.Count > 0 Then
            For k As Integer = 0 To listMiddleTopPoint1.Count - 1
                If k = 0 Then
                    Dim stDist As Double = (listMiddleTopPoint1.Item(0) - middleLeftPoint2).Length
                    If stDist > 0.01 Then
                        poly3d2.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
                        middleLeftPoint2 = listMiddleTopPoint1.Item(k)
                    Else
                        poly3d2.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
                    End If
                ElseIf k = listMiddleTopPoint1.Count - 1 Then
                    Dim stDist As Double = (listMiddleTopPoint1.Item(k) - middleLeftNozzle).Length
                    If stDist > 0.01 Then
                        poly3d2.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
                        middleLeftNozzle = listMiddleTopPoint1.Item(k)
                    Else
                        poly3d2.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
                    End If
                Else
                    poly3d2.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
                End If
            Next k
        End If
        'poly3d2.Add(New Vector3D(middleLeftNozzle, elevPointLeft4))
        'poly3d2.Add(New Vector3D(middleRightNozzle, elevPointRight4))
        If listMiddleTopPoint2.Count > 0 Then
            For k As Integer = 0 To listMiddleTopPoint2.Count - 1
                If k = 0 Then
                    Dim stDist As Double = (listMiddleTopPoint2.Item(0) - middleRightPoint2).Length
                    If stDist > 0.01 Then
                        poly3d2.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
                        middleRightNozzle = listMiddleTopPoint2.Item(k)
                    Else
                        poly3d2.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
                    End If
                ElseIf k = listMiddleTopPoint2.Count - 1 Then
                    Dim stDist As Double = (listMiddleTopPoint2.Item(k) - middleRightPoint2).Length
                    If stDist > 0.01 Then
                        poly3d2.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
                        middleRightPoint2 = listMiddleTopPoint2.Item(k)
                    Else
                        poly3d2.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
                    End If
                Else
                    poly3d2.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
                End If
            Next k
        End If
        poly3d2.Add(New Vector3D(middleRightPoint1, elevPointRight4))
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'средняя линия горизонтальной площадки
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        listMiddleTopPoint1 = MathFunction.InscribeCircleBetweenSegments(middleLeftPoint1, middleLeftPoint2, rightNozzlePt2, rightNozzlePt1, r1, countArcSegments)
        listMiddleTopPoint2 = MathFunction.InscribeCircleBetweenSegments(rightNozzlePt1, rightNozzlePt2, middleRightPoint1, middleRightPoint2, r2, countArcSegments)
        poly3d3.Add(New Vector3D(middleLeftPoint2, elevPointLeft4))
        If listMiddleTopPoint1.Count > 0 Then
            For k As Integer = 1 To listMiddleTopPoint1.Count - 2
                poly3d3.Add(New Vector3D(listMiddleTopPoint1.Item(k), elevPointLeft4))
            Next k
        End If
        poly3d3.Add(New Vector3D(rightNozzlePt1, elevPointLeft4))
        poly3d3.Add(New Vector3D(rightNozzlePt2, elevPointRight4))
        If listMiddleTopPoint2.Count > 0 Then
            For k As Integer = 1 To listMiddleTopPoint2.Count - 2
                poly3d3.Add(New Vector3D(listMiddleTopPoint2.Item(k), elevPointRight4))
            Next k
        End If
        poly3d3.Add(New Vector3D(middleRightPoint2, elevPointRight4))
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'линия в уровне земли
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim pointLeftGround1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleLeftPoint1, reverseNormalLeftAngle, lenghtSlopeLeftCone - r1)
        Dim pointLeftGround2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleLeftPoint2, reverseNormalLeftAngle, lenghtSlopeLeftCone - r1)

        Dim pointRightGround1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleRightPoint1, normalRightAngle, lenghtSlopeRightCone - r2)
        Dim pointRightGround2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleRightPoint2, normalRightAngle, lenghtSlopeRightCone - r2)

        Dim lenOffset2 As Double = lenghtSlopeRightCone - r2
        Dim pointMiddleGround1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleLeftNozzle, normalAngleNozzle, lenOffset2)
        Dim pointMiddleGround2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middleRightNozzle, normalAngleNozzle, lenOffset2)

        listMiddleTopPoint1 = MathFunction.InscribeCircleBetweenSegments(pointLeftGround1, pointLeftGround2, pointMiddleGround2, pointMiddleGround1, r1 * 2, countArcSegments)
        listMiddleTopPoint2 = MathFunction.InscribeCircleBetweenSegments(pointMiddleGround1, pointMiddleGround2, pointRightGround2, pointRightGround1, r2 * 2, countArcSegments)

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'назначаем высоты
        poly3d4.Add(New Vector3D(pointLeftGround1, ElevationGround))
        For i As Integer = 0 To listMiddleTopPoint1.Count - 1
            Dim vert As Vector3D = New Vector3D(listMiddleTopPoint1.Item(i), ElevationGround)
            poly3d4.Add(vert)
        Next
        For i As Integer = 0 To listMiddleTopPoint2.Count - 1
            Dim vert As Vector3D = New Vector3D(listMiddleTopPoint2.Item(i), ElevationGround)
            poly3d4.Add(vert)
        Next
        poly3d4.Add(New Vector3D(pointRightGround1, ElevationGround))
        result.Add(poly3d1)
        result.Add(poly3d2)
        result.Add(poly3d3)
        result.Add(poly3d4)
        Return result
    End Function
End Class
