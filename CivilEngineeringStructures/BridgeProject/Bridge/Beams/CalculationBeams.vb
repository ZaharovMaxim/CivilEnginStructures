Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Excel
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Runtime.Communications
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Public Class CalculationBeams
    Public Enum rectoreBeam
        siteMonolit
        pointSupport
        fullBeam
    End Enum
    Friend Shared Sub ApplyTrajectoryOffsets(beam As BeamI, trajectory As TrajectoryPlacementBeams)
        beam.offsetSurface = trajectory.OffsetProjectSurface
        beam.axisOffset = trajectory.OffsetProjectAlignment
    End Sub

    Friend Shared Sub EnsureElevationCalculated(succeeded As Boolean, beam As BeamI)
        If succeeded Then Return
        Dim spanNumber As Integer = If(beam Is Nothing, 0, beam.numberProlet)
        Dim rowNumber As Integer = If(beam Is Nothing, 0, beam.numberRow)
        Throw New BuildStageException(
            "Коррекция отметок балки",
            "Не удалось определить отметки балки: пролёт " & spanNumber &
                ", ряд " & rowNumber & ". Проверьте покрытие проектной поверхности.",
            "Проверьте, что вся балка находится в границах проектной поверхности.",
            "CalculationBeams.correctElevation")
    End Sub
    '==========================================================================================================
    'функция ищет ось балки
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
    '=========================================================================================================
    'раскладка балок с фиксированным зазором
    Public Shared Function calculatePlacementFixedBeams(ByRef beamsInRow As List(Of StructureElement), ByVal lineCabinetWall As DwgLine, dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByRef traectoryPlacementBeams As Dictionary(Of Integer, StructureElement), ByVal projectAlignment As Alignment, ByVal projectSurface As Surface, ByVal userBridge As Bridges) As Boolean
        Dim result As Boolean = False
        If IsNothing(beamsInRow) = True Then
            MsgBox("Список балок пустой. Сооружение не построено!")
            Return False
        End If
        If beamsInRow.Count = 0 Then
            MsgBox("Список балок пустой. Сооружение не построено!")
            Return False
        End If
        'определяем начальную опору
        Dim startPlacementNumberPillar As Integer = Pillar.getDefinitNumberPillar(dictAxisPillar)
        If startPlacementNumberPillar = 0 Then
            MsgBox("Определяющая опора для раскладки балок не выбрана. Сооружение не построено.")
            Return False
        End If
        'роверяем наличие проектной поверхности и ее содержимое
        If IsNothing(projectAlignment) = True Then
            MsgBox("Проектная ось трассы автодороги не найдена. Сооружение не построено!")
            Return False
        End If
        If projectAlignment.Plan.CompoundLine.Length = 0 Then
            MsgBox("Проектная ось трассы автодороги имеет нулевую длину. Сооружение не построено!")
            Return False
        End If
        'роверяем наличие проектной поверхности и ее содержимое
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена. Сооружение не построено!")
            Return False
        End If
        If projectSurface.Triangles.Count = 0 Then
            MsgBox("Проектная поверхность пустая. Сооружение не построено!")
            Return False
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'временный объект его удлинняем для гарантированного пересечения с осью раскладки балок
        Dim boolSelectLineCabinetWall As Boolean = False
        Dim tempLineAxisPillar As DwgLine = New DwgLine()
        If IsNothing(lineCabinetWall) = False Then
            If lineCabinetWall.Length > 0 Then
                tempLineAxisPillar.StartPoint = lineCabinetWall.StartPoint
                tempLineAxisPillar.EndPoint = lineCabinetWall.EndPoint
                boolSelectLineCabinetWall = True
            End If
        End If
        If tempLineAxisPillar.Length = 0 Then
            If dictAxisPillar.ContainsKey(startPlacementNumberPillar) = True Then
                Dim listAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(startPlacementNumberPillar)
                If IsNothing(listAxisPillar) = False Then
                    If listAxisPillar.Count > 1 Then
                        Dim dataPillar As StructureElement = listAxisPillar(1)
                        If IsNothing(dataPillar.DWGEntity) = False Then
                            tempLineAxisPillar = dataPillar.DWGEntity
                        End If
                    End If
                End If
            End If
        End If
        Dim alignmentPolyline As New Polyline3D()
        projectAlignment.Plan.CompoundLine.ToPolyLine(alignmentPolyline)
        Dim intersections As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(
            alignmentPolyline, tempLineAxisPillar.StartPoint.Pos, tempLineAxisPillar.EndPoint.Pos)
        If intersections Is Nothing OrElse Not intersections.Any() Then Return False
        Dim currentStation As Double = 0
        Dim currentOffset As Double = 0
        If Not projectAlignment.Plan.CompoundLine.PosToStaOffset(
            intersections.First(), currentStation, currentOffset) Then Return False
        Dim startPlacementPk As Double = userBridge.startPlacementPosition + userBridge.HorizontalOffset
        If Not Pillar.moveAxisPillarToStation(projectAlignment, tempLineAxisPillar, startPlacementPk) Then Return False
        Dim boolExtend As Boolean = BridgeGeometry.extendLine(tempLineAxisPillar, userBridge.LeftStructureWidth, userBridge.RightStructureWidth)
        'смещение от оси
        Dim offsetAxisPlacementBeams As Double = 0
        'Dim dictAxisPlacementBeams As Dictionary(Of Double, List(Of Polyline3D)) = New Dictionary(Of Double, List(Of Polyline3D))
        'предыдущая балка
        Dim prevUserBeam As BeamI = Nothing
        Dim axisLineBeamI As DwgLine = New DwgLine()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'последовательно достаем балки в прямом направлнии
        For i As Integer = 0 To beamsInRow.Count - 1
            Try
                'достаем данные балки
                Dim dataBeamI As StructureElement = beamsInRow.ElementAt(i)
                If IsNothing(dataBeamI) = True Then
                    Return False
                End If
                Dim userBeamI As BeamI = dataBeamI.getBeamI()
                If IsNothing(userBeamI) = True Then
                    Return False
                End If
                'определяем номер пролета выбранной балки
                Dim numberProlet As Integer = userBeamI.numberProlet
                If numberProlet >= startPlacementNumberPillar Then
                    Dim monolitBeam1 As Double = userBeamI.startLenghtMonolith 'участки омоноличивания
                    Dim monolitBeam2 As Double = userBeamI.endLenghtMonolith
                    Dim clearence As Double = userBeamI.clearence 'проектный зазор
                    If dictAxisPillar.ContainsKey(numberProlet) = True Then
                        Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet)
                        Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                        Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                        'для левого ряда (меньше или равно 0)
                        clearence = userAxisPillar.Clearence
                        If userBeamI.numberRow > 0 Then
                            If userAxisPillar.RightClearence > 0 Then
                                clearence = userAxisPillar.RightClearence
                            End If
                        End If
                        monolitBeam1 = userAxisPillar.SiteMonolit
                        userBeamI.startLenghtMonolith = monolitBeam1
                    End If
                    'ищем второй участок омоноличивания
                    If dictAxisPillar.ContainsKey(numberProlet + 1) = True Then
                        Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet + 1)
                        Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                        Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                        monolitBeam2 = userAxisPillar.SiteMonolit
                        userBeamI.endLenghtMonolith = monolitBeam2
                    End If
                    'зазор должен быть больше 0
                    If clearence <= 0 Then Return False
                    If clearence > 0 Then
                        'находим траекторию раскладки балок
                        Dim dataTraectoryPlacementBeams As StructureElement = Nothing
                        Dim userTraectoryPlacementBeams As TrajectoryPlacementBeams = Nothing
                        If traectoryPlacementBeams.ContainsKey(userBeamI.numberRow) = True Then
                            dataTraectoryPlacementBeams = traectoryPlacementBeams.Item(userBeamI.numberRow)
                            userTraectoryPlacementBeams = dataTraectoryPlacementBeams.getAxisPlacementBeams()
                        End If
                        If IsNothing(userTraectoryPlacementBeams) = True Then
                            MsgBox("Траектория для раскладки балок ряда: " & userBeamI.numberRow & ", не найдена. Сооружение не построено.")
                            Return False
                        End If
                        ApplyTrajectoryOffsets(userBeamI, userTraectoryPlacementBeams)
                        'формируем ось раскладки балок
                        Dim axisPlacementBeams As Polyline3D = New Polyline3D
                        Dim polylinePlacementBeams As DwgPolyline = dataTraectoryPlacementBeams.DWGEntity
                        polylinePlacementBeams.GetPolyline(axisPlacementBeams)
                        'если это первый пролет в этом циклк
                        If IsNothing(prevUserBeam) = True Then
                            'делаем пересечение осей и находим начальную точку раскладки балок
                            Dim startPointPlacementBeams3d As Vector3D = New Vector3D(0, 0, 0)
                            Dim startPointPlacementBeams2d As Vector3D = New Vector3D(0, 0, 0)
                            Dim startPK As Double = -9999
                            Dim startOff As Double = -999
                            Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlacementBeams, tempLineAxisPillar.StartPoint.Pos, tempLineAxisPillar.EndPoint.Pos)
                            If pointIntersectCollection.Count = 0 Then Return False
                            If pointIntersectCollection.Count > 0 Then
                                'высота начальной точки раскладки по низу балки
                                startPointPlacementBeams2d = pointIntersectCollection(0)
                                Dim boolFindPk As Boolean = axisPlacementBeams.PosToStaOffset(startPointPlacementBeams2d, startPK, startOff)
                                If boolFindPk = False Then
                                    MsgBox("Не удалось найти пересечение оси раскладки балок и оси опоры для балки: Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                    Return False
                                End If
                            End If
                            'если раскладка идет от линии шкафной стенки, то отодвигаем балку на величину участка опирания балки
                            If boolSelectLineCabinetWall = True Then
                                startPK += userBeamI.a
                                startPointPlacementBeams2d = axisPlacementBeams.StaOffsetToPos(startPK, 0)
                            End If
                            'находим высоту начальной точки
                            Dim elevStartPoint As Double = 0
                            Dim boolFindElevation As Boolean = FuncSurface.getElevationToSurface(projectSurface, startPointPlacementBeams2d, elevStartPoint)
                            If boolFindElevation = False Then
                                MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                Return False
                            End If
                            startPointPlacementBeams3d = New Vector3D(startPointPlacementBeams2d, elevStartPoint)
                            'определяем длину балки
                            Dim shortLenghtBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                            If shortLenghtBeam <= 0 Then
                                MsgBox("Проектная балка имеет нулевую длину. " & "Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                Return False
                            End If
                            'вычисляем второй конец раскладки балки
                            Dim endPointPlacementBeams As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPlacementBeams, startPointPlacementBeams3d, shortLenghtBeam, projectSurface, userBeamI.height, userBeamI.offsetSurface)
                            If endPointPlacementBeams.X = -1 And endPointPlacementBeams.Y = -1 And endPointPlacementBeams.Z = -1 Then
                                MsgBox("Ошибка в определении конечной точки опирания балки в пролете №: " & userBeamI.numberProlet & "Сооржение не построено!!!")
                                Return False
                            End If
                            'создаем новую линию балки
                            axisLineBeamI.StartPoint = startPointPlacementBeams3d
                            axisLineBeamI.EndPoint = endPointPlacementBeams
                            If IsNothing(lineCabinetWall) = True And numberProlet > 1 Then
                                'раскладка от средней оси опоры
                                clearence = clearence / 2
                            ElseIf IsNothing(lineCabinetWall) = True And numberProlet = 1 Then
                                clearence = 0
                            End If
                            'считаем минимальный зазор
                            For k = 0 To 10
                                Dim listPointStartBeam As List(Of Vector3D) = New List(Of Vector3D)
                                Dim listPointEndBeam As List(Of Vector3D) = New List(Of Vector3D)
                                Dim boolRectoreBeams As Boolean = restoreElementsBeam(axisLineBeamI, userBeamI, listPointStartBeam, listPointEndBeam, rectoreBeam.fullBeam)
                                If listPointStartBeam.Count = 4 And listPointEndBeam.Count = 4 Then
                                    Dim distLineMove As Double = 9999
                                    If clearence > 0 Then
                                        For j As Integer = 0 To 3
                                            'расчет зазора от вертикальной плоскости
                                            Dim tempmove As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(tempLineAxisPillar.StartPoint.Pos, tempLineAxisPillar.EndPoint.Pos, listPointStartBeam.Item(j), listPointEndBeam.Item(j))
                                            If tempmove < distLineMove Then
                                                distLineMove = tempmove
                                            End If
                                        Next j
                                        If distLineMove <> clearence Then
                                            distLineMove = clearence - distLineMove
                                        Else
                                            Exit For
                                        End If
                                        'переносим линию
                                        Dim boolMoveLine As Boolean = BridgeGeometry.moveLine(axisLineBeamI, distLineMove)
                                    End If
                                    'корректируем балку в плане и в высоте
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(axisLineBeamI, axisPlacementBeams, theoryShortLineBeam)
                                    'опускаем балку на нужную высоту
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                                    If Not boolElevBeam Then Return False
                                    If Math.Abs(theoryShortLineBeam - axisLineBeamI.Length) <= 0.0005 Then
                                        Exit For
                                    End If
                                End If
                            Next k
                            userBeamI._elementBridgePoint.StartAxisPoint = axisLineBeamI.StartPoint
                            userBeamI._elementBridgePoint.EndAxisPoint = axisLineBeamI.EndPoint
                        Else
                            'уже есть предыдущая балка, значит нужно найти точку пересечения оси раскладки и оси балки
                            Dim startAlignPoint3d As Vector3D = prevUserBeam._elementBridgePoint.EndAxisPoint
                            Dim pkStartPoint As Double = 0
                            Dim off As Double = 0
                            Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPlacementBeams, startAlignPoint3d.Pos, pkStartPoint, off)
                            If boolStartPoint = True Then
                                pkStartPoint = pkStartPoint + prevUserBeam.b + userBeamI.clearence + userBeamI.a
                                Dim startAlignPoint As Vector2D = axisPlacementBeams.StaOffsetToPos(pkStartPoint, 0)
                                Dim elevationST As Double = 0
                                Dim boolFindElevation As Boolean = FuncSurface.getElevationToSurface(projectSurface, startAlignPoint, elevationST)
                                If boolFindElevation = False Then
                                    MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                    Return False
                                End If
                                Dim distEndPointPr As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPlacementBeams, startAlignPoint3d, distEndPointPr, projectSurface, userBeamI.height, userBeamI.offsetSurface)
                                Dim LineShortBeam As DwgLine = dataBeamI.DWGEntity
                                LineShortBeam.StartPoint = startAlignPoint3d
                                LineShortBeam.EndPoint = endAlignPoint3d
                                userBeamI._elementBridgePoint.StartAxisPoint = startAlignPoint3d
                                userBeamI._elementBridgePoint.EndAxisPoint = endAlignPoint3d
                                '==============================================================================================================
                                'оформляем балку
                                For k As Integer = 0 To 10
                                    Dim userListClearence As List(Of Double) = New List(Of Double)
                                    Dim minZazor As Double = getClearenceBeams(prevUserBeam, userBeamI, userListClearence)
                                    Dim deltaTrimBeam As Double = Math.Round(clearence - minZazor, 4)
                                    If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                        Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(LineShortBeam, deltaTrimBeam)
                                    End If
                                    userBeamI._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                    userBeamI._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPlacementBeams, theoryShortLineBeam)
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeamI, projectSurface, userBeamI.offsetSurface)
                                    If Not boolElevBeam Then Return False
                                    userBeamI._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                    userBeamI._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                    If Math.Abs(theoryShortLineBeam - LineShortBeam.Length) <= 0.001 And Math.Abs(deltaTrimBeam) <= 0.001 Then
                                        Exit For
                                    End If
                                Next k
                            Else
                                MsgBox("Не удалось определить начальный пикет раскладки балки. Сооружение не построено!!!")
                                Return False
                            End If
                        End If
                        'возвращаем значение в словарь
                        Dim lineBeam As DwgLine = dataBeamI.DWGEntity
                        lineBeam.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
                        lineBeam.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
                        dataBeamI.DWGEntity = lineBeam
                        Dim strGSONBeamI As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeamI)
                        dataBeamI.KeyParameter = strGSONBeamI
                        beamsInRow(i) = dataBeamI
                        prevUserBeam = userBeamI
                    End If
                End If
            Catch ex As System.Exception
                Throw
            End Try
        Next i
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'последовательно достаем балки в обратном направлнии
        For i As Integer = beamsInRow.Count - 1 To 0 Step -1
            Try
                'достаем данные балки
                Dim dataBeamI As StructureElement = beamsInRow.ElementAt(i)
                If IsNothing(dataBeamI) = True Then
                    Return False
                End If
                Dim userBeamI As BeamI = dataBeamI.getBeamI()
                If IsNothing(userBeamI) = True Then
                    Return False
                End If
                'определяем номер пролета выбранной балки
                Dim numberProlet As Integer = userBeamI.numberProlet
                If numberProlet < startPlacementNumberPillar Then
                    Dim monolitBeam2 As Double = userBeamI.startLenghtMonolith 'участки омоноличивания
                    Dim monolitBeam1 As Double = userBeamI.endLenghtMonolith
                    Dim clearence As Double = userBeamI.clearence 'проектный зазор
                    If dictAxisPillar.ContainsKey(numberProlet + 1) = True Then
                        Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet + 1)
                        Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                        Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                        'для левого ряда (меньше или равно 0)
                        clearence = userAxisPillar.Clearence
                        If userBeamI.numberRow > 0 Then
                            If userAxisPillar.RightClearence > 0 Then
                                clearence = userAxisPillar.RightClearence
                            End If
                        End If
                        monolitBeam2 = userAxisPillar.SiteMonolit
                        userBeamI.endLenghtMonolith = monolitBeam2
                    End If
                    'ищем второй участок омоноличивания
                    If dictAxisPillar.ContainsKey(numberProlet) = True Then
                        Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet)
                        Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                        Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                        monolitBeam1 = userAxisPillar.SiteMonolit
                        userBeamI.startLenghtMonolith = monolitBeam1
                    End If
                    'зазор должен быть больше 0
                    If clearence <= 0 Then Return False
                    If clearence > 0 Then
                        'находим траекторию раскладки балок
                        Dim dataTraectoryPlacementBeams As StructureElement = Nothing
                        Dim userTraectoryPlacementBeams As TrajectoryPlacementBeams = Nothing
                        If traectoryPlacementBeams.ContainsKey(userBeamI.numberRow) = True Then
                            dataTraectoryPlacementBeams = traectoryPlacementBeams.Item(userBeamI.numberRow)
                            userTraectoryPlacementBeams = dataTraectoryPlacementBeams.getAxisPlacementBeams()
                        End If
                        If IsNothing(userTraectoryPlacementBeams) = True Then
                            MsgBox("Траектория для раскладки балок ряда: " & userBeamI.numberRow & ", не найдена. Сооружение не построено.")
                            Return False
                        End If
                        ApplyTrajectoryOffsets(userBeamI, userTraectoryPlacementBeams)
                        'формируем ось раскладки балок
                        Dim axisPlacementBeams As Polyline3D = New Polyline3D
                        Dim polylinePlacementBeams As DwgPolyline = dataTraectoryPlacementBeams.DWGEntity
                        polylinePlacementBeams = FuncAlignment.getReverseDwgPolyline(polylinePlacementBeams)
                        polylinePlacementBeams.GetPolyline(axisPlacementBeams)
                        'если это первый пролет в этом циклк
                        If IsNothing(prevUserBeam) = True Then
                            'делаем пересечение осей и находим начальную точку раскладки балок
                            Dim startPointPlacementBeams3d As Vector3D = New Vector3D(0, 0, 0)
                            Dim startPointPlacementBeams2d As Vector3D = New Vector3D(0, 0, 0)
                            Dim startPK As Double = -9999
                            Dim startOff As Double = -999
                            Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlacementBeams, tempLineAxisPillar.StartPoint.Pos, tempLineAxisPillar.EndPoint.Pos)
                            If pointIntersectCollection.Count = 0 Then Return False
                            If pointIntersectCollection.Count > 0 Then
                                'высота начальной точки раскладки по низу балки
                                startPointPlacementBeams2d = pointIntersectCollection(0)
                                Dim boolFindPk As Boolean = axisPlacementBeams.PosToStaOffset(startPointPlacementBeams2d, startPK, startOff)
                                If boolFindPk = False Then
                                    MsgBox("Не удалось найти пересечение оси раскладки балок и оси опоры для балки: Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                    Return False
                                End If
                            End If
                            'если раскладка идет от линии шкафной стенки, то отодвигаем балку на величину участка опирания балки
                            If boolSelectLineCabinetWall = True Then
                                startPK += userBeamI.b
                                startPointPlacementBeams2d = axisPlacementBeams.StaOffsetToPos(startPK, 0)
                            End If
                            'находим высоту начальной точки
                            Dim elevStartPoint As Double = 0
                            Dim boolFindElevation As Boolean = FuncSurface.getElevationToSurface(projectSurface, startPointPlacementBeams2d, elevStartPoint)
                            If boolFindElevation = False Then
                                MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                Return False
                            End If
                            startPointPlacementBeams3d = New Vector3D(startPointPlacementBeams2d, elevStartPoint)
                            'определяем длину балки
                            Dim shortLenghtBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                            If shortLenghtBeam <= 0 Then
                                MsgBox("Проектная балка имеет нулевую длину. " & "Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                Return False
                            End If
                            'вычисляем второй конец раскладки балки
                            Dim endPointPlacementBeams As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPlacementBeams, startPointPlacementBeams3d, shortLenghtBeam, projectSurface, userBeamI.height, userBeamI.offsetSurface)
                            If endPointPlacementBeams.X = -1 And endPointPlacementBeams.Y = -1 And endPointPlacementBeams.Z = -1 Then
                                MsgBox("Ошибка в определении конечной точки опирания балки в пролете №:" & userBeamI.numberProlet & ". Сооржение не построено!!!")
                                Return False
                            End If
                            'создаем новую линию балки
                            axisLineBeamI.StartPoint = startPointPlacementBeams3d
                            axisLineBeamI.EndPoint = endPointPlacementBeams
                            If IsNothing(lineCabinetWall) = True And numberProlet > 1 Then
                                'раскладка от средней оси опоры
                                clearence = clearence / 2
                            ElseIf IsNothing(lineCabinetWall) = True And numberProlet = userBridge.ProletCount Then
                                clearence = 0
                            End If
                            'считаем минимальный зазор
                            For k = 0 To 10
                                Dim listPointStartBeam As List(Of Vector3D) = New List(Of Vector3D)
                                Dim listPointEndBeam As List(Of Vector3D) = New List(Of Vector3D)
                                Dim boolRectoreBeams As Boolean = restoreElementsBeam(axisLineBeamI, userBeamI, listPointStartBeam, listPointEndBeam, rectoreBeam.fullBeam)
                                If listPointStartBeam.Count = 4 And listPointEndBeam.Count = 4 Then
                                    Dim distLineMove As Double = 9999
                                    If clearence > 0 Then
                                        For j As Integer = 0 To 3
                                            'расчет зазора от вертикальной плоскости
                                            Dim tempmove As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(tempLineAxisPillar.StartPoint.Pos, tempLineAxisPillar.EndPoint.Pos, listPointStartBeam.Item(j), listPointEndBeam.Item(j))
                                            If tempmove < distLineMove Then
                                                distLineMove = tempmove
                                            End If
                                        Next j
                                        If distLineMove <> clearence Then
                                            distLineMove = clearence - distLineMove
                                        Else
                                            Exit For
                                        End If
                                        'переносим линию
                                        Dim boolMoveLine As Boolean = BridgeGeometry.moveLine(axisLineBeamI, distLineMove)
                                    End If
                                    'корректируем балку в плане и в высоте
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(axisLineBeamI, axisPlacementBeams, theoryShortLineBeam)
                                    'опускаем балку на нужную высоту
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                                    If Not boolElevBeam Then Return False
                                    If Math.Abs(theoryShortLineBeam - axisLineBeamI.Length) <= 0.0005 Then
                                        Exit For
                                    End If
                                End If
                            Next k
                            userBeamI._elementBridgePoint.StartAxisPoint = axisLineBeamI.StartPoint
                            userBeamI._elementBridgePoint.EndAxisPoint = axisLineBeamI.EndPoint
                        Else
                            'уже есть предыдущая балка, значит нужно найти точку пересечения оси раскладки и оси балки
                            Dim startAlignPoint3d As Vector3D = prevUserBeam._elementBridgePoint.EndAxisPoint
                            Dim pkStartPoint As Double = 0
                            Dim off As Double = 0
                            Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPlacementBeams, startAlignPoint3d.Pos, pkStartPoint, off)
                            If boolStartPoint = True Then
                                pkStartPoint = pkStartPoint + prevUserBeam.b + userBeamI.clearence + userBeamI.b
                                Dim startAlignPoint As Vector2D = axisPlacementBeams.StaOffsetToPos(pkStartPoint, 0)
                                Dim elevationST As Double = 0
                                Dim boolFindElevation As Boolean = FuncSurface.getElevationToSurface(projectSurface, startAlignPoint, elevationST)
                                If boolFindElevation = False Then
                                    MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                                    Return False
                                End If
                                Dim distEndPointPr As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                Dim endAlignPoint3d As Vector3D = CalculationBeams.correctionLenghtBeamToElevation(axisPlacementBeams, startAlignPoint3d, distEndPointPr, projectSurface, userBeamI.height, userBeamI.offsetSurface)
                                Dim LineShortBeam As DwgLine = dataBeamI.DWGEntity
                                LineShortBeam.StartPoint = startAlignPoint3d
                                LineShortBeam.EndPoint = endAlignPoint3d
                                userBeamI._elementBridgePoint.StartAxisPoint = startAlignPoint3d
                                userBeamI._elementBridgePoint.EndAxisPoint = endAlignPoint3d
                                '==============================================================================================================
                                'оформляем балку
                                For k As Integer = 0 To 10
                                    Dim userListClearence As List(Of Double) = New List(Of Double)
                                    Dim minZazor As Double = getClearenceBeams(prevUserBeam, userBeamI, userListClearence)
                                    Dim deltaTrimBeam As Double = Math.Round(clearence - minZazor, 4)
                                    If Math.Abs(deltaTrimBeam) >= 0.001 Then
                                        Dim boolMoveBearm As Boolean = BridgeGeometry.moveLine(LineShortBeam, deltaTrimBeam)
                                    End If
                                    userBeamI._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                    userBeamI._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                    'делаем коррекцию балки в плане
                                    Dim theoryShortLineBeam As Double = userBeamI.lenght - userBeamI.a - userBeamI.b
                                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(LineShortBeam, axisPlacementBeams, theoryShortLineBeam)
                                    'корректируем балку по высоте
                                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(LineShortBeam, userBeamI, projectSurface, userBeamI.offsetSurface)
                                    If Not boolElevBeam Then Return False
                                    userBeamI._elementBridgePoint.StartAxisPoint = LineShortBeam.StartPoint
                                    userBeamI._elementBridgePoint.EndAxisPoint = LineShortBeam.EndPoint
                                    If Math.Abs(theoryShortLineBeam - LineShortBeam.Length) <= 0.001 And Math.Abs(deltaTrimBeam) <= 0.001 Then
                                        Exit For
                                    End If
                                Next k
                            Else
                                MsgBox("Не удалось определить начальный пикет раскладки балки. Сооружение не построено!!!")
                                Return False
                            End If
                        End If
                        'возвращаем значение в словарь
                        Dim lineBeam As DwgLine = dataBeamI.DWGEntity
                        lineBeam.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
                        lineBeam.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
                        dataBeamI.DWGEntity = lineBeam
                        Dim strGSONBeamI As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeamI)
                        dataBeamI.KeyParameter = strGSONBeamI
                        beamsInRow(i) = dataBeamI
                        prevUserBeam = userBeamI
                    End If
                End If
            Catch ex As System.Exception
                Throw
            End Try
        Next i
        '=======================================================================================================================
        'проверяем направление балок
        For i As Integer = 0 To beamsInRow.Count - 1
            'достаем данные балки
            Dim dataBeamI As StructureElement = beamsInRow.ElementAt(i)
            Dim boolrezBir As Boolean = correctionAxisDirectionBeam(dataBeamI, projectAlignment)
        Next
        Return True
    End Function
    Public Shared Function calculatePlacementFloatBeams(ByRef beamsInRow As List(Of StructureElement), dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByRef traectoryPlacementBeams As Dictionary(Of Integer, StructureElement), ByVal projectAlignment As Alignment, ByVal projectSurface As Surface, ByVal userBridge As Bridges) As Boolean
        Dim result As Boolean = False
        If IsNothing(beamsInRow) = True Then
            MsgBox("Список балок пустой. Сооружение не построено!")
            Return False
        End If
        If beamsInRow.Count = 0 Then
            MsgBox("Список балок пустой. Сооружение не построено!")
            Return False
        End If
        'определяем начальную опору
        Dim startPlacementNumberPillar As Integer = Pillar.getDefinitNumberPillar(dictAxisPillar)
        If startPlacementNumberPillar = 0 Then
            MsgBox("Определяющая опора для раскладки балок не выбрана. Сооружение не построено.")
            Return False
        End If
        'роверяем наличие проектной поверхности и ее содержимое
        If IsNothing(projectAlignment) = True Then
            MsgBox("Проектная ось трассы автодороги не найдена. Сооружение не построено!")
            Return False
        End If
        If projectAlignment.Plan.CompoundLine.Length = 0 Then
            MsgBox("Проектная ось трассы автодороги имеет нулевую длину. Сооружение не построено!")
            Return False
        End If
        'роверяем наличие проектной поверхности и ее содержимое
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена. Сооружение не построено!")
            Return False
        End If
        If projectSurface.Triangles.Count = 0 Then
            MsgBox("Проектная поверхность пустая. Сооружение не построено!")
            Return False
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'For i As Integer = 0 To dictAxisPillar.Count - 2
        'последовательно достаем балки в прямом направлнии
        'предыдущая балка
        Dim prevUserBeam As BeamI = Nothing
        For j As Integer = 0 To beamsInRow.Count - 1
            Try
                'достаем данные балки
                Dim dataBeamI As StructureElement = beamsInRow.ElementAt(j)
                If IsNothing(dataBeamI) = True Then
                    Return False
                End If
                Dim userBeamI As BeamI = dataBeamI.getBeamI()
                If IsNothing(userBeamI) = True Then
                    Return False
                End If
                Dim axisLineBeamI As DwgLine = New DwgLine
                Dim numberProlet As Integer = userBeamI.numberProlet
                'находим оси опор
                'первая опра
                Dim listAxisPillar1 As List(Of StructureElement) = dictAxisPillar.Item(numberProlet)
                Dim dataPillar1 As StructureElement = listAxisPillar1.Item(1)
                Dim userAxisPillar1 As Pillar = dataPillar1.getPillar
                Dim axisLinePillar1 As DwgLine = dataPillar1.DWGEntity
                'следующая опора
                Dim listAxisPillar2 As List(Of StructureElement) = dictAxisPillar.Item(numberProlet + 1)
                Dim dataPillar2 As StructureElement = listAxisPillar2.Item(1)
                Dim userAxisPillar2 As Pillar = dataPillar2.getPillar
                Dim axisLinePillar2 As DwgLine = dataPillar2.DWGEntity
                If axisLinePillar1.Length <= 0 Or axisLinePillar2.Length <= 0 Then Return False
                If axisLinePillar1.Length > 0 And axisLinePillar2.Length > 0 Then
                    'определяем номер пролета выбранной балки
                    If numberProlet >= startPlacementNumberPillar Then
                        Dim monolitBeam1 As Double = userBeamI.startLenghtMonolith 'участки омоноличивания
                        Dim firstClearence As Double = userAxisPillar1.Clearence 'проектный зазор на первой опоре
                        If userBeamI.numberRow > 0 Then
                            If userAxisPillar1.RightClearence > 0 Then
                                firstClearence = userAxisPillar1.RightClearence
                            End If
                        End If
                        monolitBeam1 = userAxisPillar1.SiteMonolit
                        userBeamI.startLenghtMonolith = monolitBeam1
                        'ищем второй участок омоноличивания
                        Dim monolitBeam2 As Double = userBeamI.endLenghtMonolith
                        Dim lastClearence As Double = userAxisPillar2.Clearence 'проектный зазор на первой опоре
                        If userBeamI.numberRow > 0 Then
                            If userAxisPillar2.RightClearence > 0 Then
                                lastClearence = userAxisPillar2.RightClearence
                            End If
                        End If
                        monolitBeam2 = userAxisPillar2.SiteMonolit
                        userBeamI.endLenghtMonolith = monolitBeam2
                        'находим траекторию раскладки балок
                        Dim dataTraectoryPlacementBeams As StructureElement = Nothing
                        Dim userTraectoryPlacementBeams As TrajectoryPlacementBeams = Nothing
                        If traectoryPlacementBeams.ContainsKey(userBeamI.numberRow) = True Then
                            dataTraectoryPlacementBeams = traectoryPlacementBeams.Item(userBeamI.numberRow)
                            userTraectoryPlacementBeams = dataTraectoryPlacementBeams.getAxisPlacementBeams()
                        End If
                        If IsNothing(userTraectoryPlacementBeams) = True Then
                            MsgBox("Траектория для раскладки балок ряда: " & userBeamI.numberRow & ", не найдена. Сооружение не построено.")
                            Return False
                        End If
                        ApplyTrajectoryOffsets(userBeamI, userTraectoryPlacementBeams)
                        'формируем ось раскладки балок
                        Dim axisPlacementBeams As Polyline3D = New Polyline3D
                        Dim polylinePlacementBeams As DwgPolyline = dataTraectoryPlacementBeams.DWGEntity
                        polylinePlacementBeams.GetPolyline(axisPlacementBeams)
                        '================================================================================================================
                        'делаем пересечение осей и находим начальную точку раскладки балок
                        Dim startPointPlacementBeams2d As Vector2D = New Vector2D(0, 0)
                        Dim startPointPlacementBeams3d As Vector3D = New Vector3D(0, 0, 0)
                        Dim startPK As Double = -9999
                        Dim startOff As Double = -999
                        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlacementBeams, axisLinePillar1.StartPoint.Pos, axisLinePillar1.EndPoint.Pos)
                        If pointIntersectCollection.Count = 0 Then Return False
                        If pointIntersectCollection.Count > 0 Then
                            'высота начальной точки раскладки по низу балки
                            startPointPlacementBeams2d = pointIntersectCollection(0)
                            Dim boolFindPk As Boolean = axisPlacementBeams.PosToStaOffset(startPointPlacementBeams2d, startPK, startOff)
                            If boolFindPk = False Then
                                MsgBox("Не удалось найти пересечение оси раскладки балок и оси опоры для балки: Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                Return False
                            End If
                        End If
                        'находим высоту начальной точки
                        Dim elevStartPoint As Double = 0
                        Dim boolFindElevation As Boolean = FuncSurface.getElevationToSurface(projectSurface, startPointPlacementBeams2d, elevStartPoint)
                        If boolFindElevation = False Then
                            MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                            Return False
                        End If
                        startPointPlacementBeams3d = New Vector3D(startPointPlacementBeams2d, elevStartPoint)
                        '================================================================================================================
                        'делаем пересечение осей и находим конечную точку раскладки балок
                        Dim endPointPlacementBeams2d As Vector2D = New Vector2D(0, 0)
                        Dim endPointPlacementBeams3d As Vector3D = New Vector3D(0, 0, 0)
                        Dim endPK As Double = -9999
                        Dim endOff As Double = -999
                        pointIntersectCollection = PolylineExtentions.GetIntersections(axisPlacementBeams, axisLinePillar2.StartPoint.Pos, axisLinePillar2.EndPoint.Pos)
                        If pointIntersectCollection.Count = 0 Then Return False
                        If pointIntersectCollection.Count > 0 Then
                            'высота начальной точки раскладки по низу балки
                            endPointPlacementBeams2d = pointIntersectCollection(0)
                            Dim boolFindPk As Boolean = axisPlacementBeams.PosToStaOffset(endPointPlacementBeams2d, endPK, endOff)
                            If boolFindPk = False Then
                                MsgBox("Не удалось найти пересечение оси раскладки балок и оси опоры для балки: Пролет: " & userBeamI.numberProlet & ", Ряд: " & userBeamI.numberRow)
                                Return False
                            End If
                        End If
                        'находим высоту начальной точки
                        Dim elevEndPoint As Double = 0
                        boolFindElevation = FuncSurface.getElevationToSurface(projectSurface, endPointPlacementBeams2d, elevEndPoint)
                        If boolFindElevation = False Then
                            MsgBox("Ошибка в определении высоты начальной точки опирания балки. Сооржение не построено!!!")
                            Return False
                        End If
                        endPointPlacementBeams3d = New Vector3D(endPointPlacementBeams2d, elevEndPoint)
                        'создаем новую линию балки
                        axisLineBeamI.StartPoint = startPointPlacementBeams3d
                        axisLineBeamI.EndPoint = endPointPlacementBeams3d
                        '============================================================================================================
                        'считаем минимальный зазор начальной осью опоры
                        Dim listPointStartBeam As List(Of Vector3D) = New List(Of Vector3D)
                        Dim listPointEndBeam As List(Of Vector3D) = New List(Of Vector3D)
                        If IsNothing(prevUserBeam) = True Then
                            Dim boolRectoreBeams As Boolean = restoreElementsBeam(axisLineBeamI, userBeamI, listPointStartBeam, listPointEndBeam, rectoreBeam.fullBeam)
                            If listPointStartBeam.Count = 4 And listPointEndBeam.Count = 4 Then
                                Dim distLineMove As Double = 9999
                                For k As Integer = 0 To 3
                                    'расчет зазора от вертикальной плоскости
                                    Dim tempmove As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(axisLinePillar1.StartPoint.Pos, axisLinePillar1.EndPoint.Pos, listPointStartBeam.Item(k), listPointEndBeam.Item(k))
                                    If tempmove < distLineMove Then
                                        distLineMove = tempmove
                                    End If
                                Next k
                                If distLineMove <> firstClearence Then
                                    distLineMove = firstClearence - distLineMove
                                End If
                                'делаем обрезку линии вначале
                                Dim boolMoveLine As Boolean = BridgeGeometry.extendLine(axisLineBeamI, -1 * distLineMove, 0)
                                'опускаем балку на нужную высоту
                                Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                                If Not boolElevBeam Then Return False
                            End If
                        Else
                            'делаем обрезку линии вначале
                            Dim boolMoveLine As Boolean = BridgeGeometry.extendLine(axisLineBeamI, -1 * (firstClearence / 2), 0)
                            'опускаем балку на нужную высоту
                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                            If Not boolElevBeam Then Return False
                        End If
                        '============================================================================================================
                        'считаем минимальный зазор с конечной осью опоры
                        listPointStartBeam = New List(Of Vector3D)
                        listPointEndBeam = New List(Of Vector3D)
                        If numberProlet = userBridge.ProletCount Then
                            Dim boolRectoreBeams2 = restoreElementsBeam(axisLineBeamI, userBeamI, listPointStartBeam, listPointEndBeam, rectoreBeam.fullBeam)
                            If listPointStartBeam.Count = 4 And listPointEndBeam.Count = 4 Then
                                Dim distLineMove As Double = 9999
                                If lastClearence > 0 Then
                                    Dim tempClearence As Double = lastClearence / 2
                                    For k As Integer = 0 To 3
                                        'расчет зазора от вертикальной плоскости
                                        Dim tempmove As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(axisLinePillar2.StartPoint.Pos, axisLinePillar2.EndPoint.Pos, listPointEndBeam.Item(k), listPointStartBeam.Item(k))
                                        If tempmove < distLineMove Then
                                            distLineMove = tempmove
                                        End If
                                    Next k
                                    If distLineMove <> lastClearence Then
                                        distLineMove = lastClearence - distLineMove
                                    End If
                                    'делаем обрезку линии вначале
                                    Dim boolMoveLine As Boolean = BridgeGeometry.extendLine(axisLineBeamI, 0, -1 * distLineMove)
                                End If
                                'опускаем балку на нужную высоту
                                Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                                If Not boolElevBeam Then Return False
                            End If
                        Else
                            'делаем обрезку линии вначале
                            Dim boolMoveLine As Boolean = BridgeGeometry.extendLine(axisLineBeamI, 0, -1 * (firstClearence / 2))
                            'опускаем балку на нужную высоту
                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                            If Not boolElevBeam Then Return False
                        End If
                        userBeamI._elementBridgePoint.StartAxisPoint = axisLineBeamI.StartPoint
                        userBeamI._elementBridgePoint.EndAxisPoint = axisLineBeamI.EndPoint
                        userBeamI.lenght = Math.Round(axisLineBeamI.Length + userBeamI.a + userBeamI.b, 3)
                        ''делаем коррекцию по зазору
                        'If IsNothing(prevUserBeam) = False Then
                        '    Dim userListClearence As List(Of Double) = New List(Of Double)
                        '    Dim minZazor As Double = getClearenceBeams(prevUserBeam, userBeamI, userListClearence)
                        '    Dim deltaTrimBeam As Double = Math.Round(firstClearence - minZazor, 4)
                        '    If Math.Abs(deltaTrimBeam) >= 0.001 Then
                        '        Dim boolMoveBearm As Boolean = BridgeGeometry.extendLine(axisLineBeamI, deltaTrimBeam, 0)
                        '    End If
                        '    'корректируем балку по высоте
                        '    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineBeamI, userBeamI, projectSurface, userBeamI.offsetSurface)
                        'End If
                        'возвращаем значение в словарь
                        Dim lineBeam As DwgLine = dataBeamI.DWGEntity
                        lineBeam.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
                        lineBeam.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
                        dataBeamI.DWGEntity = lineBeam
                        Dim strGSONBeamI As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeamI)
                        dataBeamI.KeyParameter = strGSONBeamI
                        beamsInRow(j) = dataBeamI
                        prevUserBeam = userBeamI
                    End If
                End If
            Catch ex As System.Exception
                Throw
            End Try
        Next j
        'End If
        'Next i
        Return True
    End Function
    '==========================================================================================================
    'раскладка балок между двумя осями опирания
    Public Shared Function calculatePositionMiddleBeam(ByRef beamsInRow As List(Of StructureElement), ByVal dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByRef traectoryPlacementBeams As Dictionary(Of Integer, StructureElement), ByVal projectAlignment As Alignment, ByVal projectSurface As Surface, ByVal userBridge As Bridges) As Boolean
        Dim result As Boolean = False
        If IsNothing(beamsInRow) = True Then
            Return result
        End If
        For i As Integer = 0 To beamsInRow.Count - 1
            Dim dataBeam As StructureElement = beamsInRow.ElementAt(i)
            If IsNothing(dataBeam) = True Then
                Return result
            End If
            Dim userBeamI As BeamI = dataBeam.getBeamI()
            If IsNothing(userBeamI) = True Then
                Return result
            End If
            Dim numberProlet As Integer = userBeamI.numberProlet
            Dim listFirstAxisPillar As List(Of StructureElement) = Nothing
            Dim listSecondAxisPillar As List(Of StructureElement) = Nothing
            If dictAxisPillar.ContainsKey(numberProlet) = True Then
                listFirstAxisPillar = dictAxisPillar.Item(numberProlet)
            End If
            If dictAxisPillar.ContainsKey(numberProlet + 1) = True Then
                listSecondAxisPillar = dictAxisPillar.Item(numberProlet + 1)
            End If
            If IsNothing(listFirstAxisPillar) = True Or IsNothing(listSecondAxisPillar) = True Then
                Return result
            End If
            Dim numberRow As Integer = userBeamI.numberRow
            If traectoryPlacementBeams.ContainsKey(numberRow) Then
                Dim dataTraectoryPlacementBeams As StructureElement = traectoryPlacementBeams.Item(numberRow)
                Dim userTraectoryPlacementBeams As TrajectoryPlacementBeams = dataTraectoryPlacementBeams.getAxisPlacementBeams()
                If IsNothing(userTraectoryPlacementBeams) = True Then Return result
                ApplyTrajectoryOffsets(userBeamI, userTraectoryPlacementBeams)
                Dim polylinePlacementBeams As DwgPolyline = dataTraectoryPlacementBeams.DWGEntity
                Dim traectoryBeams As Polyline3D = New Polyline3D
                polylinePlacementBeams.GetPolyline(traectoryBeams)
                Dim firstAxisPillar As StructureElement = Nothing
                Dim secondAxisPillar As StructureElement = Nothing
                If dictAxisPillar.Count = 2 Then
                    firstAxisPillar = listFirstAxisPillar.Item(1)
                    secondAxisPillar = listSecondAxisPillar.Item(1)
                Else
                    If numberProlet = 1 Then
                        firstAxisPillar = listFirstAxisPillar.Item(1)
                        secondAxisPillar = listSecondAxisPillar.Item(0)
                    ElseIf numberProlet = userBridge.ProletCount Then
                        firstAxisPillar = listFirstAxisPillar.Item(2)
                        secondAxisPillar = listSecondAxisPillar.Item(1)
                    Else
                        firstAxisPillar = listFirstAxisPillar.Item(2)
                        secondAxisPillar = listSecondAxisPillar.Item(0)
                    End If
                End If
                'участки омоноличивания
                Dim monolitBeam1 As Double = userBeamI.startLenghtMonolith 'участки омоноличивания
                Dim monolitBeam2 As Double = userBeamI.endLenghtMonolith
                If dictAxisPillar.ContainsKey(numberProlet) = True Then
                    Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet)
                    Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                    Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                    monolitBeam1 = userAxisPillar.SiteMonolit
                    userBeamI.startLenghtMonolith = monolitBeam1
                End If
                'ищем второй участок омоноличивания
                If dictAxisPillar.ContainsKey(numberProlet + 1) = True Then
                    Dim listDataAxisPillar As List(Of StructureElement) = dictAxisPillar.Item(numberProlet + 1)
                    Dim dataAxisPillar As StructureElement = listDataAxisPillar.Item(1)
                    Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                    monolitBeam2 = userAxisPillar.SiteMonolit
                    userBeamI.endLenghtMonolith = monolitBeam2
                End If
                'пересечение
                Dim axisBeamsPillar1 As DwgLine = firstAxisPillar.DWGEntity
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(traectoryBeams, axisBeamsPillar1.StartPoint.Pos, axisBeamsPillar1.EndPoint.Pos)
                If pointIntersectCollection.Count = 0 Then
                    'удлинняем ось чтобы получить гарантированное пересечение
                    Dim newStartPoint As Vector3D = axisBeamsPillar1.StartPoint
                    Dim newEndPoint As Vector3D = axisBeamsPillar1.EndPoint
                    Dim boolExtendLine As Boolean = MathFunction.FuncExtendPos(newStartPoint, newEndPoint, axisBeamsPillar1.Length, axisBeamsPillar1.Length)
                    pointIntersectCollection = PolylineExtentions.GetIntersections(traectoryBeams, newStartPoint.Pos, newEndPoint.Pos)
                End If
                If pointIntersectCollection.Count = 0 Then
                    Return result
                End If
                Dim startPointAxisBeam As Vector3D = New Vector3D(pointIntersectCollection.ElementAt(0), 0)
                '3. ищем пересечение оси трассы со второй осью опирания
                'пересечение
                Dim axisBeamsPillar2 As DwgLine = secondAxisPillar.DWGEntity
                pointIntersectCollection = PolylineExtentions.GetIntersections(traectoryBeams, axisBeamsPillar2.StartPoint.Pos, axisBeamsPillar2.EndPoint.Pos)
                If pointIntersectCollection.Count = 0 Then
                    'удлинняем ось чтобы получить гарантированное пересечение
                    Dim newStartPoint As Vector3D = axisBeamsPillar2.StartPoint
                    Dim newEndPoint As Vector3D = axisBeamsPillar2.EndPoint
                    Dim boolExtendLine As Boolean = MathFunction.FuncExtendPos(newStartPoint, newEndPoint, axisBeamsPillar2.Length, axisBeamsPillar2.Length)
                    pointIntersectCollection = PolylineExtentions.GetIntersections(traectoryBeams, newStartPoint.Pos, newEndPoint.Pos)
                End If
                If pointIntersectCollection.Count = 0 Then
                    Return result
                End If
                Dim endPointAxisBeam As Vector3D = New Vector3D(pointIntersectCollection.ElementAt(0), 0)
                '4 находим высоты точер пересечения
                Dim elvation1 As Double = 0
                Dim boolFindElevation1 As Boolean = FuncSurface.getElevationToSurface(projectSurface, startPointAxisBeam.Pos, elvation1)
                Dim elvation2 As Double = 0
                Dim boolFindElevation2 As Boolean = FuncSurface.getElevationToSurface(projectSurface, endPointAxisBeam.Pos, elvation2)
                If Not boolFindElevation1 OrElse Not boolFindElevation2 Then Return False
                If boolFindElevation1 = True And boolFindElevation2 = True Then
                    startPointAxisBeam = New Vector3D(startPointAxisBeam.Pos, elvation1)
                    endPointAxisBeam = New Vector3D(endPointAxisBeam.Pos, elvation2)
                    Dim axisLineShortBeam As DwgLine = New DwgLine()
                    axisLineShortBeam.StartPoint = startPointAxisBeam
                    axisLineShortBeam.EndPoint = endPointAxisBeam
                    If Not (userBeamI.nameAlbum Like "Индивидуального проектирования") Then
                        For j As Integer = 0 To 10
                            Dim theoryShortLineBeam As Double = Math.Round(userBeamI.lenght - userBeamI.a - userBeamI.b, 3)
                            'делаем коррекцию балки в плане
                            Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(axisLineShortBeam, traectoryBeams, theoryShortLineBeam)
                            'делаем коррекцию по высоте
                            Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineShortBeam, userBeamI, projectSurface, userBeamI.offsetSurface)
                            If Not boolElevBeam Then Return False
                            Dim deltaLenght As Double = Math.Round(theoryShortLineBeam - (userBeamI.lenght - userBeamI.a - userBeamI.b))
                            If Math.Abs(deltaLenght) <= 0.0005 Then
                                Exit For
                            End If
                        Next
                    Else
                        userBeamI.lenght = Math.Round((endPointAxisBeam - startPointAxisBeam).Length, 3)
                    End If
                    'возвращаем значение в словарь
                    userBeamI._elementBridgePoint.StartAxisPoint = axisLineShortBeam.StartPoint
                    userBeamI._elementBridgePoint.EndAxisPoint = axisLineShortBeam.EndPoint
                    Dim lineBeam As DwgLine = dataBeam.DWGEntity
                    lineBeam.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
                    lineBeam.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
                    dataBeam.DWGEntity = lineBeam
                    Dim strGSONBeamI As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeamI)
                    dataBeam.KeyParameter = strGSONBeamI
                End If
            End If
        Next i
        Return True
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

                    If Double.IsNaN(elev1) OrElse Double.IsInfinity(elev1) OrElse
                       Double.IsNaN(elev2) OrElse Double.IsInfinity(elev2) OrElse
                       Double.IsNaN(elev3) OrElse Double.IsInfinity(elev3) OrElse
                       Double.IsNaN(elev4) OrElse Double.IsInfinity(elev4) Then
                        Return False
                    End If

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
        Dim newAxisLineTopBeam As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(axisLineTopBeam, hBeam)
        axisLineTopBeam.Dispose()
        If IsNothing(newAxisLineTopBeam) = True Then Return False
        axisLineTopBeam = newAxisLineTopBeam
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
                    If Math.Abs(deltaLenght) > 0.0005 Then
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
            Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
            Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
            If startAlignPointLeftElevation > startAlignPointRightElevation Then
                startAlignPointLeftElevation = startAlignPointRightElevation
            End If
            startPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            boolFindIntersectPoint = True
        Else
            'нет пересечения пытаемся удлиннить ось
            Dim newAxisPillar As DwgLine = New DwgLine
            newAxisPillar.StartPoint = startAxisPillar.StartPoint
            newAxisPillar.EndPoint = startAxisPillar.EndPoint
            Dim boolExtendLine As Boolean = BridgeGeometry.extendLine(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
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
                Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
                Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
                If startAlignPointLeftElevation > startAlignPointRightElevation Then
                    startAlignPointLeftElevation = startAlignPointRightElevation
                End If
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
                Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
                Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
                If startAlignPointLeftElevation > startAlignPointRightElevation Then
                    startAlignPointLeftElevation = startAlignPointRightElevation
                End If
                endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            Else
                'нет пересечения пытаемся удлиннить ось
                Dim newAxisPillar As DwgLine = New DwgLine
                newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                Dim boolExtendLine As Boolean = BridgeGeometry.extendLine(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                If pointIntersectCollection.Count > 0 Then
                    startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                    'вычисляем верх балки
                    Dim startDistPrBeam As Double = -1
                    Dim off As Double = -1
                    Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                    Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.offsetSurface / 2)
                    Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.offsetSurface / 2)
                    Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
                    Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
                    If startAlignPointLeftElevation > startAlignPointRightElevation Then
                        startAlignPointLeftElevation = startAlignPointRightElevation
                    End If
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
                    zazor = userAxisPillar.RightClearence

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
                        Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
                        Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
                        If startAlignPointLeftElevation > startAlignPointRightElevation Then
                            startAlignPointLeftElevation = startAlignPointRightElevation
                        End If
                        endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                    End If
                Else
                    'нет пересечения пытаемся удлиннить ось
                    Dim newAxisPillar As DwgLine = New DwgLine
                    newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                    newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                    Dim boolExtendLine As Boolean = BridgeGeometry.extendLine(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                    pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                    If pointIntersectCollection.Count > 0 Then
                        startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                        'вычисляем верх балки
                        Dim startDistPrBeam As Double = -1
                        Dim off As Double = -1
                        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                        Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTopPlateLeft)
                        Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTopPlateRight)
                        Dim startAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, ptTopLeft, "Отметка левого края балки") - userBeam.offsetSurface
                        Dim startAlignPointRightElevation As Double = RequireSurfaceElevation(surf, ptTopRight, "Отметка правого края балки") - userBeam.offsetSurface
                        If startAlignPointLeftElevation > startAlignPointRightElevation Then
                            startAlignPointLeftElevation = startAlignPointRightElevation
                        End If
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
                Dim endAlignPointLeftElevation As Double = RequireSurfaceElevation(surf, leftPointBeam, "Отметка левого края балки") - userBeam.offsetSurface
                Dim endAlignPointRightElevation As Double = RequireSurfaceElevation(surf, rightPointBeam, "Отметка правого края балки") - userBeam.offsetSurface
                If endAlignPointLeftElevation > endAlignPointRightElevation Then
                    endAlignPointLeftElevation = endAlignPointRightElevation
                End If
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
    Public Shared Function correctionLenghtBeamToElevation(ByVal axisPline3D As Polyline3D, ByVal startPoint As Vector3D, ByVal radius As Double, ByVal surf As Surface, ByVal heightBeam As Double, ByVal dEarth As Double) As Vector3D
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
    '==========================================================================================================
    'находит балку с минимальной высотой (список - 2 пролета)
    Public Shared Function getBeamToMinElevation(ByVal listBeamsProlet As List(Of Dictionary(Of Integer, StructureElement)), Optional ByVal arraySubFermenters As SubFermenters() = Nothing) As Double
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
                                        'берем отметку начала отрезка
                                        If elevStartBeam < minElevationBottomBeam Then
                                            minElevationBottomBeam = elevStartBeam
                                            'находим нужный подферменник и вычитаем из отметки его просвет с балкой
                                            If IsNothing(arraySubFermenters) = False Then
                                                If arraySubFermenters.Length > 0 Then
                                                    Dim userBeam As BeamI = dataBeam.getBeamI()
                                                    For k As Integer = 0 To arraySubFermenters.Length - 1
                                                        Dim subFermenter As SubFermenters = arraySubFermenters(k)
                                                        If subFermenter.NumberPillar = userBeam.numberProlet Then
                                                            If userBeam.numberProlet = subFermenter.NumberProlet Then
                                                                If userBeam.numberRow = subFermenter.NumberRow Then
                                                                    minElevationBottomBeam = minElevationBottomBeam - subFermenter.DeltaHeightBeam
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                        End If
                                    Else
                                        If elevEndBeam < minElevationBottomBeam Then
                                            minElevationBottomBeam = elevEndBeam
                                            'находим нужный подферменник и вычитаем из отметки его просвет с балкой
                                            If IsNothing(arraySubFermenters) = False Then
                                                If arraySubFermenters.Length > 0 Then
                                                    Dim userBeam As BeamI = dataBeam.getBeamI()
                                                    For k As Integer = 0 To arraySubFermenters.Length - 1
                                                        Dim subFermenter As SubFermenters = arraySubFermenters(k)
                                                        If subFermenter.NumberPillar = userBeam.numberProlet - 1 Then
                                                            If userBeam.numberProlet = subFermenter.NumberProlet Then
                                                                If userBeam.numberRow = subFermenter.NumberRow Then
                                                                    minElevationBottomBeam = minElevationBottomBeam - subFermenter.DeltaHeightBeam
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Next k
                                                End If
                                            End If
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
    '===========================================================================================================
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
    '===========================================================================================================
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
    '=========================================================================================================
    'функция возвращает предыдущую балку
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
    '==========================================================================================================
    'Функция рассчитывает положение оси балки между осями опирания
    Public Shared Function calculatePositionBeam(ByRef userBeam As BeamI, ByVal axisPlacementBeams3D As Polyline3D, ByVal projectSurface As Surface, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Boolean
        Dim result As Boolean = False
        If IsNothing(axisPlacementBeams3D) = True Then
            Return result
        End If
        '1. ищем 2 оси опирания для выбранной балки
        Dim numberProlet As Integer = userBeam.numberProlet
        Dim listAxisBeamsPillar As List(Of StructureElement) = AxisBeamsPillars.getAxisPillarBeamsInProlet(dictionaryObjectsBridge, numberProlet)
        'обязательно должно быть найдено 2 оси опирания
        If listAxisBeamsPillar.Count <> 2 Then Return result
        '2. ищем пересечение оси трассы с первой осью опирания
        Dim dataAxisBeamsPillar1 As StructureElement = listAxisBeamsPillar.Item(0)
        Dim axisBeamsPillar1 As DwgLine = dataAxisBeamsPillar1.DWGEntity
        'пересечение
        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPlacementBeams3D, axisBeamsPillar1.StartPoint.Pos, axisBeamsPillar1.EndPoint.Pos)
        If pointIntersectCollection.Count = 0 Then
            'удлинняем ось чтобы получить гарантированное пересечение
            Dim newStartPoint As Vector3D = axisBeamsPillar1.StartPoint
            Dim newEndPoint As Vector3D = axisBeamsPillar1.EndPoint
            Dim boolExtendLine As Boolean = MathFunction.FuncExtendPos(newStartPoint, newEndPoint, axisBeamsPillar1.Length, axisBeamsPillar1.Length)
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPlacementBeams3D, newStartPoint.Pos, newEndPoint.Pos)
        End If
        If pointIntersectCollection.Count = 0 Then
            Return result
        End If
        Dim startPointAxisBeam As Vector3D = New Vector3D(pointIntersectCollection.ElementAt(0), 0)
        '3. ищем пересечение оси трассы со второй осью опирания
        Dim dataAxisBeamsPillar2 As StructureElement = listAxisBeamsPillar.Item(1)
        Dim axisBeamsPillar2 As DwgLine = dataAxisBeamsPillar2.DWGEntity
        'пересечение
        pointIntersectCollection = PolylineExtentions.GetIntersections(axisPlacementBeams3D, axisBeamsPillar2.StartPoint.Pos, axisBeamsPillar2.EndPoint.Pos)
        If pointIntersectCollection.Count = 0 Then
            'удлинняем ось чтобы получить гарантированное пересечение
            Dim newStartPoint As Vector3D = axisBeamsPillar2.StartPoint
            Dim newEndPoint As Vector3D = axisBeamsPillar2.EndPoint
            Dim boolExtendLine As Boolean = MathFunction.FuncExtendPos(newStartPoint, newEndPoint, axisBeamsPillar2.Length, axisBeamsPillar2.Length)
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPlacementBeams3D, newStartPoint.Pos, newEndPoint.Pos)
        End If
        If pointIntersectCollection.Count = 0 Then
            Return result
        End If
        Dim endPointAxisBeam As Vector3D = New Vector3D(pointIntersectCollection.ElementAt(0), 0)
        '4 находим высоты точер пересечения
        Dim elvation1 As Double = 0
        Dim boolFindElevation1 As Boolean = FuncSurface.getElevationToSurface(projectSurface, startPointAxisBeam.Pos, elvation1)
        Dim elvation2 As Double = 0
        Dim boolFindElevation2 As Boolean = FuncSurface.getElevationToSurface(projectSurface, endPointAxisBeam.Pos, elvation2)
        If Not boolFindElevation1 OrElse Not boolFindElevation2 Then Return False
        If boolFindElevation1 = True And boolFindElevation2 = True Then
            startPointAxisBeam = New Vector3D(startPointAxisBeam.Pos, elvation1)
            endPointAxisBeam = New Vector3D(endPointAxisBeam.Pos, elvation2)
            Dim axisLineShortBeam As DwgLine = New DwgLine()
            axisLineShortBeam.StartPoint = startPointAxisBeam
            axisLineShortBeam.EndPoint = endPointAxisBeam
            If Not (userBeam.nameAlbum Like "Индивидуального проектирования") Then
                For i As Integer = 0 To 10
                    Dim theoryShortLineBeam As Double = Math.Round(userBeam.lenght - userBeam.a - userBeam.b, 3)
                    'делаем коррекцию балки в плане
                    Dim boolCorrBeam As Boolean = CalculationBeams.correctionLenght(axisLineShortBeam, axisPlacementBeams3D, theoryShortLineBeam)
                    'делаем коррекцию по высоте
                    Dim boolElevBeam As Boolean = CalculationBeams.correctElevation(axisLineShortBeam, userBeam, projectSurface, userBeam.offsetSurface)
                    If Not boolElevBeam Then Return False
                    Dim deltaLenght As Double = Math.Round(theoryShortLineBeam - (userBeam.lenght - userBeam.a - userBeam.b))
                    If Math.Abs(deltaLenght) <= 0.001 Then
                        Exit For
                    End If
                Next
            Else
                userBeam.lenght = Math.Round((endPointAxisBeam - startPointAxisBeam).Length, 3)
            End If
            'возвращаем значение в словарь
            userBeam._elementBridgePoint.StartAxisPoint = axisLineShortBeam.StartPoint
            userBeam._elementBridgePoint.EndAxisPoint = axisLineShortBeam.EndPoint
        End If
        Return True
    End Function
    '=========================================================================================================
    'функция возвращает 4 зазора между соседними балками
    Public Shared Function getClearenceBeams(ByVal userPreviousBeamI As BeamI, ByVal userBeamI As BeamI, Optional ByRef listClearence As List(Of Double) = Nothing) As Double
        Dim result As Double = 999999
        If IsNothing(userPreviousBeamI) = True Then Exit Function
        If IsNothing(userBeamI) = True Then Exit Function
        Dim prevLineShortBeam As DwgLine = New DwgLine()
        prevLineShortBeam.StartPoint = userPreviousBeamI._elementBridgePoint.StartAxisPoint
        prevLineShortBeam.EndPoint = userPreviousBeamI._elementBridgePoint.EndAxisPoint
        Dim lineShortBeam As DwgLine = New DwgLine()
        lineShortBeam.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
        lineShortBeam.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
        If prevLineShortBeam.Length = 0 Then Return -1
        If lineShortBeam.Length = 0 Then Return -1
        'получаем сечения предыдущей балки
        Dim startSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionPrevBeam As Boolean = CalculationBeams.restoreElementsBeam(prevLineShortBeam, userPreviousBeamI, startSectionPrevBeam, endSectionPrevBeam, CalculationBeams.rectoreBeam.fullBeam)
        'получаем сечения балки
        Dim startSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionBeam As Boolean = CalculationBeams.restoreElementsBeam(lineShortBeam, userBeamI, startSectionBeam, endSectionBeam, CalculationBeams.rectoreBeam.fullBeam)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'приводим балки к одному уровню по низу
        'верх балки предыдущей
        If boolSectionBeam = True Then
            Dim topLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(0) 'лево
            Dim topRightPrevBeam As Vector3D = endSectionPrevBeam.Item(1) 'право
            'низ балки предыдущей
            Dim downLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(2) 'лево
            Dim downRightPrevBeam As Vector3D = endSectionPrevBeam.Item(3) 'право
            'высота по низу предыдущей балки
            Dim elevPrevBeam As Double = (downLeftPrevBeam.Z + downRightPrevBeam.Z) / 2
            'верх определяемая балка
            Dim topLeftBeam As Vector3D = startSectionBeam.Item(0) 'лево
            Dim topRightBeam As Vector3D = startSectionBeam.Item(1) 'право
            'низ балки определяемой
            Dim downLeftBeam As Vector3D = startSectionBeam.Item(2) 'лево
            Dim downRightBeam As Vector3D = startSectionBeam.Item(3) 'право
            'высота по низу предыдущей балки
            Dim elevBeam As Double = (downLeftBeam.Z + downRightBeam.Z) / 2
            'дельта по высоте
            Dim deltaH As Double = elevPrevBeam - elevBeam
            If Math.Abs(deltaH) > 0.01 Then
                Dim oldLenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length
                If deltaH < 0 Then
                    'предудущая балка ниже чем расчетная (приводим ее уровню расчетной бплки
                    downLeftPrevBeam = MathFunction.FuncCalcPointInLine(downLeftPrevBeam, topLeftPrevBeam, Math.Abs(deltaH), 3)
                    downRightPrevBeam = MathFunction.FuncCalcPointInLine(downRightPrevBeam, topRightPrevBeam, Math.Abs(deltaH), 3)
                    Dim lenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length - oldLenghtDownWidthBeam
                    If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                        Dim dLen As Double = lenghtDownWidthBeam / 2
                        Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftPrevBeam, downRightPrevBeam, -1 * dLen, -1 * dLen, 3)
                    End If
                Else 'расчетная балка ниже чем предыдущая (подтягиваем расчетную балку к предыдущей)
                    downLeftBeam = MathFunction.FuncCalcPointInLine(downLeftBeam, topLeftBeam, deltaH)
                    downRightBeam = MathFunction.FuncCalcPointInLine(downRightBeam, topRightBeam, deltaH)
                    Dim lenghtDownWidthBeam As Double = (downLeftBeam - downRightBeam).Length - oldLenghtDownWidthBeam
                    If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                        Dim dLen As Double = lenghtDownWidthBeam / 2
                        Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftBeam, downRightBeam, -1 * dLen, -1 * dLen, 3)
                    End If
                End If
            End If

            '====================================================================================================================
            'анализируем зазор по верху
            '====================================================================================================================
            'проверяем пересечение верха балки с горизонтальной линией предыдущей балки
            'верх лево
            Dim topLeftZazor As Double = 999999
            Dim intersectPoint As Vector2D = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(0).Pos, topLeftBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topLeftBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - topLeftPrevBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - topLeftPrevBeam.Pos).Length
                        If tempZazor < topLeftZazor Then
                            topLeftZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предыдущей балки и смотрим пересечение
            If topLeftZazor = 999999 Then
                Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
                Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'верх право
            Dim topRightZazor As Double = 999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(1).Pos, topRightBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topRightBeam.Pos).Length
                If tempZazor < topRightZazor Then
                    topRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                    If tempZazor < topRightZazor Then
                        topRightZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - topRightPrevBeam.Pos).Length
                    If tempZazor < topRightZazor Then
                        topRightZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - topRightPrevBeam.Pos).Length
                        If tempZazor < topRightZazor Then
                            topRightZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If topRightZazor = 999999 Then
                Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
                Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topRightZazor = tempZazor
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'низ лево
            Dim downLeftZazor As Double = 999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(2).Pos, downLeftBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downLeftBeam.Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - downLeftPrevBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - downLeftPrevBeam.Pos).Length
                        If tempZazor < downLeftZazor Then
                            downLeftZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If downLeftZazor = 999999 Then
                Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
                Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                End If
            End If

            'низ право
            Dim downRightZazor As Double = 999999
            intersectPoint = New Vector2D(-1, -1)
            'проверяем на предмет явного пересечения
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(3).Pos, downRightBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downRightBeam.Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                End If
            End If
            'пересечения нет, проверяем в обратную сторону
            If boolIntersect = False Then
                boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = -1 * (intersectPoint - downRightPrevBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                Else
                    'пересечение виртуального пересечения
                    boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                    If boolIntersect = True Then
                        Dim tempZazor As Double = (intersectPoint - downRightPrevBeam.Pos).Length
                        If tempZazor < downRightZazor Then
                            downRightZazor = tempZazor
                        End If
                    End If
                End If
            End If
            'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
            If downRightZazor = 999999 Then
                Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
                Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

                Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                End If
            End If
            listClearence = New List(Of Double)
            listClearence.Add(topLeftZazor)
            listClearence.Add(topRightZazor)
            listClearence.Add(downLeftZazor)
            listClearence.Add(downRightZazor)
            result = listClearence.Min
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция возвлащает балки для ряда с заданным номером ряда
    Public Shared Function getBeamsToRow(ByVal dictionaryBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal numberRow As Integer) As List(Of StructureElement)
        Dim result As List(Of StructureElement) = New List(Of StructureElement)
        If IsNothing(dictionaryBeams) = True Then Return result
        If dictionaryBeams.Count = 0 Then Return result
        For Each dictionarySpan As Dictionary(Of Integer, StructureElement) In dictionaryBeams.Values
            If IsNothing(dictionarySpan) = False Then
                For Each dataBeam As StructureElement In dictionarySpan.Values
                    If IsNothing(dataBeam) = False Then
                        Dim userBeam As BeamI = dataBeam.getBeamI()
                        If IsNothing(userBeam) = False Then
                            If userBeam.numberRow = numberRow Then
                                result.Add(dataBeam)
                            End If
                        End If
                    End If
                Next
            End If
        Next
        Return result
    End Function
    'функция проверяет и корректирует балку еcли она против направления пикетажа
    Public Shared Function correctionAxisDirectionBeam(ByRef dataBeamI As StructureElement, ByRef align As Alignment) As Boolean
        If IsNothing(dataBeamI) = True Then
            Return False
        End If
        Dim userBeamI As BeamI = dataBeamI.getBeamI()
        If IsNothing(userBeamI) = True Then
            Return False
        End If
        Dim axisLineBeamI As DwgLine = dataBeamI.DWGEntity
        If IsNothing(axisLineBeamI) = True Then
            axisLineBeamI = New DwgLine
        End If
        If axisLineBeamI.Length = 0 Then
            axisLineBeamI.StartPoint = userBeamI._elementBridgePoint.StartAxisPoint
            axisLineBeamI.EndPoint = userBeamI._elementBridgePoint.EndAxisPoint
        End If
        If axisLineBeamI.Length = 0 Then
            Return False
        End If
        If IsNothing(align) = True Then
            Return False
        End If
        'находим пикеты начала и конца предудущей балки
        Try
            Dim pkStart As Double = 0
            Dim offStart As Double = 0
            Dim boolFindPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisLineBeamI.StartPoint.Pos, pkStart, offStart)
            Dim pkEnd As Double = 0
            Dim offend As Double = 0
            Dim boolFindPk1 As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisLineBeamI.EndPoint.Pos, pkEnd, offend)
            If boolFindPk = True And boolFindPk1 = True Then
                If pkStart > pkEnd Then
                    Dim tempStartPoint As Vector3D = axisLineBeamI.StartPoint
                    Dim tempEndPoint As Vector3D = axisLineBeamI.EndPoint
                    axisLineBeamI.StartPoint = tempEndPoint
                    axisLineBeamI.EndPoint = tempStartPoint
                    userBeamI._elementBridgePoint.StartAxisPoint = axisLineBeamI.StartPoint
                    userBeamI._elementBridgePoint.EndAxisPoint = axisLineBeamI.EndPoint
                    Dim newKeyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeamI)
                    dataBeamI.KeyParameter = newKeyParam
                    Return True
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'функция вычисляет середину между смежными балками
    Public Shared Function calculateMiddlePointBeams(ByVal dataPrevBeam As StructureElement, ByVal dataBeam As StructureElement) As Vector3D
        calculateMiddlePointBeams = Nothing
        If IsNothing(dataBeam) = True Then Exit Function
        If IsNothing(dataPrevBeam) = True Then Exit Function
        Dim lineShortBeam As DwgLine = dataBeam.DWGEntity
        Dim prevLineShortBeam As DwgLine = dataPrevBeam.DWGEntity
        Dim userBeam As BeamI = dataBeam.getBeamI()
        Dim userPrevBeam As BeamI = dataPrevBeam.getBeamI()
        If lineShortBeam.Length = 0 Then
            lineShortBeam.StartPoint = userBeam._elementBridgePoint.StartAxisPoint
            lineShortBeam.EndPoint = userBeam._elementBridgePoint.EndAxisPoint
        End If
        If lineShortBeam.Length = 0 Then Exit Function
        If prevLineShortBeam.Length = 0 Then
            prevLineShortBeam.StartPoint = userPrevBeam._elementBridgePoint.StartAxisPoint
            prevLineShortBeam.EndPoint = userPrevBeam._elementBridgePoint.EndAxisPoint
        End If
        If prevLineShortBeam.Length = 0 Then Exit Function
        Dim startPointPrevBeam As Vector3D = New Vector3D()
        Dim endPointPrevBeam As Vector3D = New Vector3D()
        Dim boolFindPoint As Boolean = MathFunction.FuncVirtualExtendLine(prevLineShortBeam, 0, userPrevBeam.b, startPointPrevBeam, endPointPrevBeam)

        Dim startPointBeam As Vector3D = New Vector3D()
        Dim endPointBeam As Vector3D = New Vector3D()
        boolFindPoint = MathFunction.FuncVirtualExtendLine(lineShortBeam, userBeam.a, 0, startPointBeam, endPointBeam)

        Dim middlePoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(endPointPrevBeam, startPointBeam)
        Return middlePoint
    End Function

    Private Shared Function RequireSurfaceElevation(surface As Surface,
                                                    point As Vector2D,
                                                    valueName As String) As Double
        Dim elevation As Double = 0.0
        Try
            If FuncSurface.getElevationToSurface(surface, point, elevation) Then Return elevation
        Catch ex As Exception
            Throw New BuildStageException(
                "Чтение проектной отметки",
                valueName & " не получена в точке " & point.ToString() & ".",
                "Проверьте назначенную проектную поверхность и её триангуляцию.",
                "CalculationBeams.RequireSurfaceElevation",
                ex)
        End Try
        Throw New BuildStageException(
            "Чтение проектной отметки",
            valueName & " отсутствует в точке " & point.ToString() & ".",
            "Проверьте назначенную проектную поверхность и её триангуляцию.",
            "CalculationBeams.RequireSurfaceElevation")
    End Function
End Class
