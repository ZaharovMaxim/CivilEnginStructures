Imports System.Threading.Tasks
Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Public Class CalculationBeams
    Public Enum rectoreBeam
        siteMonolit
        pointSupport
        fullBeam
    End Enum
    Public Shared Function getConditionalRow(ByVal numberRow As Integer) As String
        Dim result As String = ""
        If numberRow < 0 Then
            result = "Л-" & Math.Abs(numberRow)
        ElseIf numberRow = 0 Then
            result = "Ось"
        Else
            result = "П-" & numberRow
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция корректирует балку по высоте
    Public Shared Function correctElevation(ByRef entityBeamI As DwgLine, ByVal userBeam As BeamI, ByVal surf As Surface, ByVal offsetSurface As Double) As Boolean
        correctElevation = False
        If IsNothing(entityBeamI) = True Then
            Return False
        ElseIf entityBeamI.Length = 0 Then
            Return False
        End If
        If IsNothing(surf) = True Then Return False
        Dim startPointBeam As Vector3D = entityBeamI.StartPoint
        Dim endPointBeam As Vector3D = entityBeamI.EndPoint
        For i As Integer = 0 To 100
            'восстанавливаем балку
            Dim startSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim endSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim boolRestoreBeam As Boolean = restoreElementsBeam(entityBeamI, userBeam, startSectionPoint3d, endSectionPoint3d, rectoreBeam.siteMonolit)
            If boolRestoreBeam = True Then
                Try
                    'корректируем балку по высоте
                    Dim elev1 As Double = surf.GetElevation(startSectionPoint3d.Item(0).Pos)
                    Dim elevF1 As Double = startSectionPoint3d.Item(0).Z
                    Dim elevTopStart1 As Double = elev1 - (offsetSurface + startSectionPoint3d.Item(0).Z)

                    Dim elev2 As Double = surf.GetElevation(startSectionPoint3d.Item(1).Pos)
                    Dim elevF2 As Double = startSectionPoint3d.Item(1).Z
                    Dim elevTopStart2 As Double = elev2 - (offsetSurface + startSectionPoint3d.Item(1).Z)

                    Dim elev3 As Double = surf.GetElevation(endSectionPoint3d.Item(0).Pos)
                    Dim elevF3 As Double = endSectionPoint3d.Item(0).Z
                    Dim elevTopStart3 As Double = elev3 - (offsetSurface + endSectionPoint3d.Item(0).Z)

                    Dim elevF4 As Double = endSectionPoint3d.Item(1).Z
                    Dim elev4 As Double = surf.GetElevation(endSectionPoint3d.Item(1).Pos)
                    Dim elevTopStart4 As Double = elev4 - (offsetSurface + endSectionPoint3d.Item(1).Z)

                    'Dim listZ As List(Of Double) = New List(Of Double) From {elevTopStart1, elevTopStart2, elevTopStart3, elevTopStart4}
                    'Dim minZ As Double = listZ.Min
                    'startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + minZ)
                    'endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + minZ)
                    If elevTopStart1 < 0 And elevTopStart2 < 0 Then
                        If Math.Abs(elevTopStart1) > Math.Abs(elevTopStart2) Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        End If
                    ElseIf elevTopStart1 > 0 And elevTopStart2 > 0 Then
                        If elevTopStart1 > elevTopStart2 Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        End If
                    ElseIf elevTopStart1 < 0 And elevTopStart2 > 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                    ElseIf elevTopStart1 > 0 And elevTopStart2 < 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                    End If

                    If elevTopStart3 < 0 And elevTopStart4 < 0 Then
                        If Math.Abs(elevTopStart3) > Math.Abs(elevTopStart4) Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        End If
                    ElseIf elevTopStart3 > 0 And elevTopStart4 > 0 Then
                        If elevTopStart3 > elevTopStart4 Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        End If
                    ElseIf elevTopStart3 < 0 And elevTopStart4 > 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                    ElseIf elevTopStart3 > 0 And elevTopStart4 < 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                    End If
                    entityBeamI.StartPoint = startPointBeam
                    entityBeamI.EndPoint = endPointBeam
                    Return True
                Catch ex As System.NullReferenceException
                    Return False
                Catch ex As System.InvalidOperationException
                    Return False
                Catch ex As System.ArgumentOutOfRangeException
                    Return False
                End Try
            End If
        Next i
    End Function
    '=========================================================================================================
    'функция восстанавливает габарит балки
    Public Shared Function restoreElementsBeam(ByVal entityBeamI As DwgLine, ByVal userBeam As BeamI, ByRef startSectionBeam As List(Of Vector3D), ByRef endSectionBeam As List(Of Vector3D), ByVal type As rectoreBeam) As Boolean
        restoreElementsBeam = False
        If IsNothing(entityBeamI) = True Then Exit Function
        If IsNothing(userBeam) = True Then Exit Function
        '1. ось балки по верху
        Dim axisLineTopBeam As DwgLine = New DwgLine()
        axisLineTopBeam.StartPoint = entityBeamI.StartPoint
        axisLineTopBeam.EndPoint = entityBeamI.EndPoint
        '2. ось балки по низу
        Dim axisLineDownBeam As DwgLine = New DwgLine()
        axisLineDownBeam.StartPoint = entityBeamI.StartPoint
        axisLineDownBeam.EndPoint = entityBeamI.EndPoint
        'удлинняем балку на величину участков опирания балок 
        If type = rectoreBeam.fullBeam Then
            Dim boolExtBearm As Boolean = MathFunction.FuncExtendPos(axisLineTopBeam.StartPoint, axisLineTopBeam.EndPoint, userBeam.a, userBeam.b)
        ElseIf type = rectoreBeam.siteMonolit Then
            Dim tempStartDist As Double = userBeam.a - userBeam.startLenghtMonolith
            Dim tempEndDist As Double = userBeam.b - userBeam.endLenghtMonolith
            Dim boolExtBearm As Boolean = MathFunction.FuncExtendPos(axisLineTopBeam.StartPoint, axisLineTopBeam.EndPoint, tempStartDist, tempEndDist)
        End If
        Dim boolExtDownBearm As Boolean = MathFunction.FuncExtendPos(axisLineDownBeam.StartPoint, axisLineDownBeam.EndPoint, userBeam.a, userBeam.b)
        Dim hBeam As Double = userBeam.height 'полная высота балки
        Dim offsetLeftPlate As Double = userBeam.widthTopPlateLeft
        Dim offsetRightPlate As Double = userBeam.widthTopPlateRight
        If userBeam.startLenghtMonolith > 0 And userBeam.endLenghtMonolith > 0 Then
            If type = rectoreBeam.fullBeam Then
                If Not (type = rectoreBeam.pointSupport) Then
                    hBeam = userBeam.height - userBeam.heightTopPlate
                    offsetLeftPlate = userBeam.WidthTop / 2
                    offsetRightPlate = userBeam.WidthTop / 2
                End If
            End If
        End If
        '1.делаем смещение балки вверх
        Dim b As Double = axisLineTopBeam.EndPoint.Z - axisLineTopBeam.StartPoint.Z
        Dim c As Double = axisLineTopBeam.Length
        Dim i As Double = Math.Asin(b / c)
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, hBeam)
        'получаем новые координаты верха балки
        Dim startPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(axisLineTopBeam.StartPoint.X, axisLineTopBeam.StartPoint.Y), axisLineTopBeam.Rotation, deltaXZ.X)
        Dim endPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(axisLineTopBeam.EndPoint.X, axisLineTopBeam.EndPoint.Y), axisLineTopBeam.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPosTopBeam3d As Vector3D = New Vector3D(startPosUpBeam2d.X, startPosUpBeam2d.Y, axisLineTopBeam.StartPoint.Z + deltaXZ.Y)
        Dim endPosTopBeam3d As Vector3D = New Vector3D(endPosUpBeam2d.X, endPosUpBeam2d.Y, axisLineTopBeam.EndPoint.Z + deltaXZ.Y)
        axisLineTopBeam.StartPoint = startPosTopBeam3d
        axisLineTopBeam.EndPoint = endPosTopBeam3d
        '====================================================================================================================
        Dim pointStart1 As Vector3D = Nothing
        Dim pointStart2 As Vector3D = Nothing
        Dim pointStart3 As Vector3D = Nothing
        Dim pointStart4 As Vector3D = Nothing
        Dim pointEnd1 As Vector3D = Nothing
        Dim pointEnd2 As Vector3D = Nothing
        Dim pointEnd3 As Vector3D = Nothing
        Dim pointEnd4 As Vector3D = Nothing
        'смещение низ право
        Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineDownBeam.Offset(dbCollection1, userBeam.widthBottom / 2)
        If dbCollection1.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection1.Item(0)
            pointStart1 = tempLine.StartPoint
            pointEnd1 = tempLine.EndPoint
        Else
            Return False
        End If
        '====================================================================================================================
        'смещение низ лево
        Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineDownBeam.Offset(dbCollection2, -1 * userBeam.widthBottom / 2)
        If dbCollection2.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection2.Item(0)
            pointStart2 = tempLine.StartPoint
            pointEnd2 = tempLine.EndPoint
        Else
            Return False
        End If
        '====================================================================================================================
        'смещение верх право
        Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection3, offsetRightPlate)
        If dbCollection3.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection3.Item(0)
            pointStart3 = tempLine.StartPoint
            pointEnd3 = tempLine.EndPoint
        Else
            Return False
        End If
        '====================================================================================================================
        'смещение верх лево
        Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection4, -1 * offsetLeftPlate)
        If dbCollection4.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection4.Item(0)
            pointStart4 = tempLine.StartPoint
            pointEnd4 = tempLine.EndPoint
        Else
            Return False
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начальное сечение балки
        startSectionBeam = New List(Of Vector3D)
        startSectionBeam.Add(pointStart4) 'верх лево
        startSectionBeam.Add(pointStart3) 'верх право
        startSectionBeam.Add(pointStart2) 'низ лево
        startSectionBeam.Add(pointStart1) 'низ право
        'конечное сечение балки
        endSectionBeam = New List(Of Vector3D)
        endSectionBeam.Add(pointEnd4) 'верх лево
        endSectionBeam.Add(pointEnd3) 'верх право
        endSectionBeam.Add(pointEnd2) 'низ лево
        endSectionBeam.Add(pointEnd1) 'низ право
        Return True
    End Function
    '=========================================================================================================
    'функция проверяет и корректирует балку еcли она против направления пикетажа
    Public Shared Function correctionAxisDirection(ByRef entityBeamI As DwgLine, ByRef align As Alignment) As Boolean
        correctionAxisDirection = False
        If IsNothing(entityBeamI) = True Then
            Return False
        End If
        If IsNothing(align) = True Then
            Return False
        End If
        If entityBeamI.Length = 0 Then
            Return False
        End If
        'находим пикеты начала и конца предудущей балки
        Try
            Dim pkStart As Double = 0
            Dim offStart As Double = 0
            Dim boolFindPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(entityBeamI.StartPoint.Pos, pkStart, offStart)
            Dim pkEnd As Double = 0
            Dim offend As Double = 0
            Dim boolFindPk1 As Boolean = align.Plan.CompoundLine.PosToStaOffset(entityBeamI.EndPoint.Pos, pkEnd, offend)
            If boolFindPk = True And boolFindPk1 = True Then
                If pkStart > pkEnd Then
                    Dim tempStartPoint As Vector3D = entityBeamI.StartPoint
                    Dim tempEndPoint As Vector3D = entityBeamI.EndPoint
                    entityBeamI.StartPoint = tempEndPoint
                    entityBeamI.EndPoint = tempStartPoint
                    Return True
                End If
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    '==========================================================================================================
    'функция корректирует балку по длине (lenghtShortBeam - это теоретическая длина от точки опиратия a, до точки опирания b)
    Public Shared Function correctionLenght(ByRef axisLineBeam As DwgLine, ByVal axisPlacementBeams As Polyline3D, ByVal lenghtShortBeam As Double) As Boolean
        correctionLenght = False
        If IsNothing(axisPlacementBeams) = True Then Exit Function
        If IsNothing(axisLineBeam) = True Then Exit Function
        If axisLineBeam.Length = 0 Then Exit Function
        If lenghtShortBeam > 0 Then
            'делаем максимуи 10 итераций (по идее достаточно 2)
            For i As Integer = 0 To 10
                Try
                    Dim startPK As Double = 0
                    Dim off As Double = 0
                    Dim startPoint3d As Vector3D = axisLineBeam.StartPoint
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPlacementBeams, axisLineBeam.StartPoint.Pos, startPK, off)
                    If boolStartPoint = True Then
                        Dim tempPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPlacementBeams, startPK, 0)
                        startPoint3d = New Vector3D(tempPoint, startPoint3d.Z)
                    End If
                    Dim endPK As Double = 0
                    Dim endPoint3d As Vector3D = axisLineBeam.EndPoint
                    Dim boolEndPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPlacementBeams, axisLineBeam.EndPoint.Pos, endPK, off)
                    If boolEndPoint = True Then
                        Dim tempPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPlacementBeams, endPK, 0)
                        endPoint3d = New Vector3D(tempPoint, endPoint3d.Z)
                    End If
                    axisLineBeam.StartPoint = startPoint3d
                    axisLineBeam.EndPoint = endPoint3d
                    Dim newLenght As Double = Math.Round((endPoint3d - startPoint3d).Length, 3)
                    Dim deltaLenght As Double = Math.Round(lenghtShortBeam - newLenght, 3)
                    If deltaLenght > 0.0015 Then
                        Dim boolExtendBeam As Boolean = MathFunction.FuncExtendPos(axisLineBeam.StartPoint, axisLineBeam.EndPoint, 0, deltaLenght, 3)
                    Else
                        Return True
                    End If
                Catch ex As Exception
                    Return False
                End Try
            Next
        End If
    End Function
    '==========================================================================================================
    'функция строит ось промежуточной балки
    Public Shared Function createMiddleAxisBeam(ByRef entityBeamI As DwgLine, ByVal startAxisPillar As DwgLine, ByVal deltaLenghtStartAxisPillars As Double, ByVal axisPolyline3D As Polyline3D, ByVal leftBeam As Boolean, Optional lenghtShortBeam As Double = 0, Optional endAxisBeamsPillar As DwgLine = Nothing, Optional endAxisPillar As DwgLine = Nothing, Optional surf As Surface = Nothing, Optional ByVal userBeam As BeamI = Nothing) As Boolean
        If IsNothing(startAxisPillar) = True Then
            Return False
        End If
        If IsNothing(axisPolyline3D) = True Then
            Return False
        End If
        If IsNothing(entityBeamI) = True Then
            entityBeamI = New DwgLine
        End If

        If IsNothing(userBeam) = True Then
            Dim dataBeam As StructureElement = Nothing
            Dim boolReadData As Boolean = FuncXRecords.getXRecords(entityBeamI, dataBeam)
            If IsNothing(dataBeam) = True Then Exit Function
            userBeam = dataBeam.getBeamI()
            If IsNothing(userBeam) = True Then Exit Function
        End If
        Dim angle As Double = startAxisPillar.Rotation
        Dim angleReverse As Double = startAxisPillar.Rotation + Math.PI
        If angleReverse > Math.PI * 2 Then
            angleReverse -= Math.PI * 2
        End If
        '==============================================================================================================================
        'крайняя левая балка
        Dim boolFindIntersectPoint As Boolean = False
        Dim startIntersectPoint2D As Vector2D = New Vector2D(0, 0)
        Dim startPointPrBeam As Vector3D = New Vector3D(0, 0, 0)
        Dim endPointPrBeam As Vector3D = New Vector3D(0, 0, 0)
        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPolyline3D, startAxisPillar.StartPoint.Pos, startAxisPillar.EndPoint.Pos)
        If pointIntersectCollection.Count > 0 Then
            startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
            Dim startDistPrBeam As Double = -1
            Dim off As Double = -1
            Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
            If boolDist = True Then
                startDistPrBeam += deltaLenghtStartAxisPillars
                startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
            End If
            'вычисляем верх балки
            Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
            Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
            Dim startAlignPointLeftElevation As Double = 0
            Dim startAlignPointRightElevation As Double = 0
            Try
                startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                If startAlignPointLeftElevation > startAlignPointRightElevation Then
                    startAlignPointLeftElevation = startAlignPointRightElevation
                End If
            Catch ex As System.NullReferenceException
            End Try
            startPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            boolFindIntersectPoint = True
        Else
            'нет пересечения пытаемся удлиннить ось
            Dim newAxisPillar As DwgLine = New DwgLine
            newAxisPillar.StartPoint = startAxisPillar.StartPoint
            newAxisPillar.EndPoint = startAxisPillar.EndPoint
            Dim boolExtendLine As Boolean = BridgeGeometry.extendBeam(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
            'заново вытаемся найти пересечение
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
            If pointIntersectCollection.Count > 0 Then
                startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                Dim startDistPrBeam As Double = -1
                Dim off As Double = -1
                Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                If boolDist = True Then
                    startDistPrBeam += deltaLenghtStartAxisPillars
                    startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
                End If
                'вычисляем верх балки
                Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
                Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
                Dim startAlignPointLeftElevation As Double = 0
                Dim startAlignPointRightElevation As Double = 0
                Try
                    startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                    startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                    If startAlignPointLeftElevation > startAlignPointRightElevation Then
                        startAlignPointLeftElevation = startAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                startPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                boolFindIntersectPoint = True
            End If
        End If
        If boolFindIntersectPoint = False Then
            Return False
        End If
        '=========================================================================================================================
        'вычисляем второй конец балки
        If lenghtShortBeam > 0 Then 'фиксированная длина
            'userBeam.lenght = lenghtShortBeam
            endPointPrBeam = calculateEndPointAxisBeam(axisPolyline3D, startPointPrBeam, lenghtShortBeam, userBeam, surf)
        ElseIf IsNothing(endAxisBeamsPillar) = False Then
            angle = endAxisBeamsPillar.Rotation
            angleReverse = endAxisBeamsPillar.Rotation + Math.PI
            If angleReverse > Math.PI * 2 Then
                angleReverse -= Math.PI * 2
            End If
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, endAxisBeamsPillar.StartPoint.Pos, endAxisBeamsPillar.EndPoint.Pos)
            If pointIntersectCollection.Count > 0 Then
                startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                'вычисляем верх балки
                Dim startDistPrBeam As Double = -1
                Dim off As Double = -1
                Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
                Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
                Dim startAlignPointLeftElevation As Double = 0
                Dim startAlignPointRightElevation As Double = 0
                Try
                    startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                    startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                    If startAlignPointLeftElevation > startAlignPointRightElevation Then
                        startAlignPointLeftElevation = startAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            Else
                'нет пересечения пытаемся удлиннить ось
                Dim newAxisPillar As DwgLine = New DwgLine
                newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                Dim boolExtendLine As Boolean = BridgeGeometry.extendBeam(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                If pointIntersectCollection.Count > 0 Then
                    startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                    'вычисляем верх балки
                    Dim startDistPrBeam As Double = -1
                    Dim off As Double = -1
                    Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                    Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.offsetSurface / 2)
                    Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.offsetSurface / 2)
                    Dim startAlignPointLeftElevation As Double = 0
                    Dim startAlignPointRightElevation As Double = 0
                    Try
                        startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                        startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                        If startAlignPointLeftElevation > startAlignPointRightElevation Then
                            startAlignPointLeftElevation = startAlignPointRightElevation
                        End If
                    Catch ex As System.NullReferenceException
                    End Try
                    endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                End If
            End If
            If boolFindIntersectPoint = False Then
                Return False
            End If
        ElseIf IsNothing(endAxisPillar) = False Then
            'смотрим зазор
            Dim dataEndPillar As StructureElement = Nothing
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(endAxisPillar, dataEndPillar)
            Dim userAxisPillar As Pillar = dataEndPillar.getPillar()
            If IsNothing(userAxisPillar) = False Then
                Dim zazor As Double = 0
                If leftBeam = True Then
                    zazor = userAxisPillar.Clearence
                Else
                    zazor = userAxisPillar.rightClearence

                End If
                pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, endAxisPillar.StartPoint.Pos, endAxisPillar.EndPoint.Pos)
                If pointIntersectCollection.Count > 0 Then
                    startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                    Dim startDistPrBeam As Double = -1
                    Dim off As Double = -1
                    Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                    If boolDist = True Then
                        startDistPrBeam = startDistPrBeam - userBeam.b - zazor / 2
                        'перевычисляем точку
                        startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
                        Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
                        Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
                        'вычисляем верх балки
                        Dim startAlignPointLeftElevation As Double = 0
                        Dim startAlignPointRightElevation As Double = 0
                        Try
                            startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                            startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                            If startAlignPointLeftElevation > startAlignPointRightElevation Then
                                startAlignPointLeftElevation = startAlignPointRightElevation
                            End If
                        Catch ex As System.NullReferenceException
                        End Try
                        endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                    End If
                Else
                    'нет пересечения пытаемся удлиннить ось
                    Dim newAxisPillar As DwgLine = New DwgLine
                    newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                    newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                    Dim boolExtendLine As Boolean = BridgeGeometry.extendBeam(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                    pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                    If pointIntersectCollection.Count > 0 Then
                        startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                        'вычисляем верх балки
                        Dim startDistPrBeam As Double = -1
                        Dim off As Double = -1
                        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                        Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
                        Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
                        Dim startAlignPointLeftElevation As Double = 0
                        Dim startAlignPointRightElevation As Double = 0
                        Try
                            startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.offsetSurface 'высота верха лево
                            startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.offsetSurface 'высота верха право
                            If startAlignPointLeftElevation > startAlignPointRightElevation Then
                                startAlignPointLeftElevation = startAlignPointRightElevation
                            End If
                        Catch ex As System.NullReferenceException
                        End Try
                        endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                    End If
                End If
            End If
            If boolFindIntersectPoint = False Then
                Return False
            End If
        End If
        entityBeamI.StartPoint = startPointPrBeam
        entityBeamI.EndPoint = endPointPrBeam
        Return True
    End Function
    '==========================================================================================================
    'функция ищет вторую точку опирания балки, без проверки допуска по зазору
    Public Shared Function calculateEndPointAxisBeam(ByVal axisPolyline3D As Polyline3D, ByVal startPointPrBeam As Vector3D, ByVal lenghtBeam As Double, ByVal userBeam As BeamI, ByVal surf As Surface) As Vector3D
        Dim result As Vector3D = New Vector3D(0, 0, 0)
        Dim startDistPrBeam As Double = -1
        Dim off As Double = -1
        If IsNothing(axisPolyline3D) = True Then
            Return result
        End If
        If IsNothing(userBeam) = True Then
            Return result
        End If
        'находим расстояние на полилинии
        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startPointPrBeam.Pos, startDistPrBeam, off)
        If boolDist = True Then
            Dim newDistEndPointPrBeam As Double = startDistPrBeam + lenghtBeam
            For i As Integer = 0 To 10
                Dim centerPointBeam As Vector2D = Nothing
                Dim leftPointBeam As Vector2D = Nothing
                Dim rightPointBeam As Vector2D = Nothing
                Try
                    centerPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, 0)
                    leftPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, -1 * userBeam.widthTopPlateLeft)
                    rightPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, userBeam.widthTopPlateRight)
                Catch ex As System.ArgumentOutOfRangeException
                    Return New Vector3D(-1, -1, -1)
                End Try
                Dim endAlignPointLeftElevation As Double = 0
                Dim endAlignPointRightElevation As Double = 0
                Try
                    endAlignPointLeftElevation = surf.GetElevation(leftPointBeam) - userBeam.offsetSurface 'высота верха лево
                    endAlignPointRightElevation = surf.GetElevation(rightPointBeam) - userBeam.offsetSurface  'высота верха право
                    If endAlignPointLeftElevation > endAlignPointRightElevation Then
                        endAlignPointLeftElevation = endAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                Dim tempPointNew As Vector3D = New Vector3D(centerPointBeam, endAlignPointLeftElevation)
                Dim tempDist As Double = (startPointPrBeam - tempPointNew).Length
                Dim dLenght As Double = lenghtBeam - tempDist
                If Math.Abs(dLenght) < 0.001 Then
                    result = tempPointNew
                    Exit For
                Else
                    newDistEndPointPrBeam = newDistEndPointPrBeam + dLenght
                End If
            Next i
        End If
        Return result
    End Function
    '==========================================================================================================
    'коррекция горизонтального расстояния за угол наклона линии
    Public Shared Function correctionLenghtBeamToElevation(ByVal axisPline As DwgPolyline, ByVal startPoint As Vector3D, ByVal radius As Double, ByVal surf As Surface, ByVal heightBeam As Double, ByVal dEarth As Double) As Vector3D
        Dim axisPline3D As IPolyline3D = New Polyline3D()
        axisPline.GetPolyline(axisPline3D)
        Dim startdist As Double = -1
        Dim off As Double = -1
        'находим расстояние на полилинии
        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startPoint.Pos, startdist, off)
        If boolDist = True Then
            Dim findDist As Double = radius
            For i As Integer = 0 To 10
                Dim delta As Double = startdist + findDist
                Dim tempPoint As Vector2D = Nothing
                Try
                    tempPoint = PolylineExtentions.StaOffsetToPos(axisPline3D, delta, 0)
                Catch ex As System.ArgumentOutOfRangeException
                    Return New Vector3D(-1, -1, -1)
                End Try

                Dim elev As Double = startPoint.Z
                Try
                    elev = surf.GetElevation(tempPoint)
                Catch ex As System.NullReferenceException
                    Return New Vector3D(-1, -1, -1)
                Catch ex As System.InvalidOperationException
                    Return New Vector3D(-1, -1, -1)
                End Try

                elev = elev - heightBeam - dEarth
                Dim tempPointNew As Vector3D = New Vector3D(tempPoint.X, tempPoint.Y, elev)
                Dim tempDist As Double = (startPoint - tempPointNew).Length
                Dim dLenght As Double = radius - tempDist
                If Math.Abs(dLenght) < 0.001 Then
                    Return tempPointNew
                Else
                    findDist = findDist + dLenght
                End If
            Next
        Else
            Return New Vector3D(-1, -1, -1)
        End If
    End Function
    'находит балку с минимальной высотой (список - 2 пролета)
    Public Shared Function getBeamToMinElevation(ByVal listBeamsProlet As List(Of Dictionary(Of Integer, StructureElement))) As Double
        Dim minElevationBottomBeam As Double = 999999 'отметка самой нижней балки в пролете
        If IsNothing(listBeamsProlet) = True Then Exit Function
        If listBeamsProlet.Count > 0 Then
            For i As Integer = 0 To listBeamsProlet.Count - 1
                Dim dictBeams As Dictionary(Of Integer, StructureElement) = listBeamsProlet.Item(i)
                If IsNothing(dictBeams) = False Then
                    If dictBeams.Count > 0 Then
                        For j As Integer = 0 To dictBeams.Count - 1
                            Dim dataBeam As StructureElement = dictBeams.ElementAt(j).Value
                            Dim axisLineBeam As DwgLine = dataBeam.DWGEntity
                            If IsNothing(axisLineBeam) = False Then
                                If axisLineBeam.Length > 0 Then
                                    Dim elevStartBeam As Double = axisLineBeam.StartPoint.Z
                                    Dim elevEndBeam As Double = axisLineBeam.EndPoint.Z
                                    'если это следующий пролет
                                    If i = 1 Then
                                        If elevStartBeam < minElevationBottomBeam Then
                                            minElevationBottomBeam = elevStartBeam
                                        End If
                                    Else
                                        If elevEndBeam < minElevationBottomBeam Then
                                            minElevationBottomBeam = elevEndBeam
                                        End If
                                    End If
                                End If
                            End If
                        Next j
                    End If
                End If
            Next i
        End If
        If minElevationBottomBeam <> 999999 Then
            minElevationBottomBeam = Math.Round(minElevationBottomBeam, 3)
        End If
        Return minElevationBottomBeam
    End Function
    'функция ищет крайние балки мостового сооружения в выбранной опоре (0-левая крайняя предыдущего пролета, 1-правая крайняя предыдущего пролета, 2- левая крайняя следующего пролета, 3 - правая крайняя следующего пролета)
    Public Shared Function getExtrmBeamsToPillar(ByVal listBeamsProlet As List(Of Dictionary(Of Integer, StructureElement))) As List(Of StructureElement)
        Dim result As List(Of StructureElement) = New List(Of StructureElement) From {Nothing, Nothing, Nothing, Nothing}
        If IsNothing(listBeamsProlet) = False Then
            If listBeamsProlet.Count > 0 Then
                Dim lastBeamDictionary As Dictionary(Of Integer, StructureElement) = listBeamsProlet.Item(0)
                If lastBeamDictionary.Count > 0 Then
                    result.Item(0) = lastBeamDictionary.First.Value
                    result.Item(1) = lastBeamDictionary.Last.Value
                End If
            End If
            If listBeamsProlet.Count > 1 Then
                Dim nextBeamDictionary As Dictionary(Of Integer, StructureElement) = listBeamsProlet.Item(1)
                If nextBeamDictionary.Count > 0 Then
                    result.Item(2) = nextBeamDictionary.First.Value
                    result.Item(3) = nextBeamDictionary.Last.Value
                End If
            End If
        End If
        Return result
    End Function
    'функция возвращает координаты точек опирания балок для выбранной опоры
    Public Shared Function getPointABeams(ByVal listBeamsProlet As List(Of Dictionary(Of Integer, StructureElement)), ByVal numberPillar As Integer) As List(Of Dictionary(Of Integer, Vector3D))
        Dim dictPointA As List(Of Dictionary(Of Integer, Vector3D)) = New List(Of Dictionary(Of Integer, Vector3D))
        Dim dictLastPoint As Dictionary(Of Integer, Vector3D) = New Dictionary(Of Integer, Vector3D)
        Dim dictNextPoint As Dictionary(Of Integer, Vector3D) = New Dictionary(Of Integer, Vector3D)
        If IsNothing(listBeamsProlet) = True Then
            Return dictPointA
        End If
        If listBeamsProlet.Count > 0 Then
            For i As Integer = 0 To 1
                Dim tempDictBeams As Dictionary(Of Integer, StructureElement) = listBeamsProlet.Item(i)
                If IsNothing(tempDictBeams) = True Then Continue For
                If tempDictBeams.Count = 0 Then Continue For
                For j As Integer = 0 To tempDictBeams.Count - 1
                    Dim dataBeam As StructureElement = tempDictBeams.Item(j)
                    If IsNothing(dataBeam) = True Then Continue For
                    Dim axisLineBeam As DwgLine = dataBeam.DWGEntity
                    If IsNothing(axisLineBeam) = True Then Continue For
                    If axisLineBeam.Length = 0 Then Continue For
                    Dim userBeam As BeamI = dataBeam.getBeamI()
                    If IsNothing(userBeam) = True Then Continue For
                    Dim numRowBeam As Integer = userBeam.numberRow
                    If i = 0 Then
                        If dictLastPoint.Count = 0 Then
                            dictLastPoint.Add(numRowBeam, axisLineBeam.EndPoint)
                        Else
                            If dictLastPoint.ContainsKey(numRowBeam) = False Then
                                dictLastPoint.Add(numRowBeam, axisLineBeam.EndPoint)
                            End If
                        End If
                    Else
                        If dictNextPoint.Count = 0 Then
                            dictNextPoint.Add(numRowBeam, axisLineBeam.StartPoint)
                        Else
                            If dictNextPoint.ContainsKey(numRowBeam) = False Then
                                dictNextPoint.Add(numRowBeam, axisLineBeam.StartPoint)
                            End If
                        End If
                    End If
                Next j
            Next i
        End If
        dictPointA.Add(dictLastPoint)
        dictPointA.Add(dictNextPoint)
        Return dictPointA
    End Function


    'функция возвращает предыдущую балку
    '=========================================================================================================
    Public Shared Function getPreviousBeam(ByVal axisLineBeam As DwgLine, ByVal dictionaryBeams As List(Of StructureElement)) As StructureElement
        If IsNothing(axisLineBeam) = True Then Return Nothing
        If axisLineBeam.Length = 0 Then Return Nothing
        If IsNothing(dictionaryBeams) = True Then Return Nothing
        If dictionaryBeams.Count = 0 Then Return Nothing
        Dim startPoint As Vector2D = axisLineBeam.StartPoint.Pos
        Dim result As StructureElement = Nothing
        Dim lenghtBeam As Double = 9999999
        For i As Integer = 0 To dictionaryBeams.Count - 1
            Dim tempDataBeam As StructureElement = dictionaryBeams.ElementAt(i)
            If IsNothing(tempDataBeam.DWGEntity) = False Then
                Dim tempAxisBeam As DwgLine = tempDataBeam.DWGEntity
                If tempAxisBeam.Length > 0 Then
                    Dim endPoint As Vector2D = tempAxisBeam.EndPoint.Pos
                    Dim tempdist As Double = (endPoint - startPoint).Length
                    If tempdist < lenghtBeam Then
                        result = tempDataBeam
                        lenghtBeam = tempdist
                    End If
                End If
            End If
        Next
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'поиск объектов в общей библиотеке элементов
    'ищет ось балки
    Public Shared Function getAxisBeam(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As StructureElement
        Dim dataAxisBeam As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberProlet < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
            Dim listAxisBeam = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisBeam)
            If IsNothing(listAxisBeam) = False Then
                If listAxisBeam.Count > 0 Then
                    For k As Integer = 0 To listAxisBeam.Count - 1
                        Dim tempData As StructureElement = listAxisBeam.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisBeam As BeamI = tempData.getBeamI()
                            If IsNothing(userAxisBeam) = False Then
                                If numberProlet = userAxisBeam.numberProlet And numberRow = userAxisBeam.numberRow Then
                                    dataAxisBeam = tempData
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataAxisBeam
    End Function


End Class
