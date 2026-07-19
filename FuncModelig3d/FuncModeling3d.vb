Imports System.Windows.Media.Animation
Imports NetTopologySuite
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Triangulate
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg.Entities
Public Class FuncModeling3d
    Public Shared Function FuncCreateBoxByPolyline(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal poly3D As Polyline3D, ByVal offsetElevation As Double) As Shell
        FuncCreateBoxByPolyline = Nothing
        If IsNothing(poly3D) = True Then Return Nothing
        Dim listUpBeam As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
        Dim listDownBeam As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
        If poly3D.Count > 3 Then
            For i As Integer = 0 To poly3D.Count - 1
                Dim vert As Topomatic.Cad.Foundation.Vector3D = poly3D.Item(i).Vertex
                If listUpBeam.Count = 0 Then
                    listUpBeam.Add(vert)
                    listDownBeam.Add(New Topomatic.Cad.Foundation.Vector3D(vert.Pos, vert.Z + offsetElevation))
                Else
                    Dim dist As Double = (vert.Pos - listUpBeam.Last.Pos).Length
                    If dist > 0.001 Then
                        listUpBeam.Add(vert)
                        listDownBeam.Add(New Topomatic.Cad.Foundation.Vector3D(vert.Pos, vert.Z + offsetElevation))
                    End If
                End If
                If listUpBeam.Count = 4 Then
                    Exit For
                End If
            Next
            If listUpBeam.Count > 3 And listDownBeam.Count > 3 Then
                Dim Shell = New Shell()
                Dim Positions As List(Of Vector3D) = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(0))
                Positions.Add(listUpBeam.Item(1))
                Positions.Add(listUpBeam.Item(2))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(2))
                Positions.Add(listUpBeam.Item(3))
                Positions.Add(listUpBeam.Item(0))
                Tools.AddFace(Shell, Positions)
                'низ
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(0))
                Positions.Add(listDownBeam.Item(3))
                Positions.Add(listDownBeam.Item(2))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(2))
                Positions.Add(listDownBeam.Item(1))
                Positions.Add(listDownBeam.Item(0))
                Tools.AddFace(Shell, Positions)

                Positions = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(3))
                Positions.Add(listUpBeam.Item(2))
                Positions.Add(listDownBeam.Item(2))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(2))
                Positions.Add(listDownBeam.Item(3))
                Positions.Add(listUpBeam.Item(3))
                Tools.AddFace(Shell, Positions)

                Positions = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(0))
                Positions.Add(listDownBeam.Item(0))
                Positions.Add(listDownBeam.Item(1))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(1))
                Positions.Add(listUpBeam.Item(1))
                Positions.Add(listUpBeam.Item(0))
                Tools.AddFace(Shell, Positions)

                Positions = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(0))
                Positions.Add(listUpBeam.Item(3))
                Positions.Add(listDownBeam.Item(3))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(3))
                Positions.Add(listDownBeam.Item(0))
                Positions.Add(listUpBeam.Item(0))
                Tools.AddFace(Shell, Positions)

                Positions = New List(Of Vector3D)()
                Positions.Add(listUpBeam.Item(2))
                Positions.Add(listUpBeam.Item(1))
                Positions.Add(listDownBeam.Item(1))
                Tools.AddFace(Shell, Positions)
                Positions = New List(Of Vector3D)()
                Positions.Add(listDownBeam.Item(1))
                Positions.Add(listDownBeam.Item(2))
                Positions.Add(listUpBeam.Item(2))
                Tools.AddFace(Shell, Positions)
                Tools.Flip(Shell)
                Return Shell
            End If
        End If
    End Function
    Public Shared Function FuncCreateBoxByPoints(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal listUpBeam As List(Of Vector3D), ByVal listDownBeam As List(Of Vector3D)) As Shell
        FuncCreateBoxByPoints = Nothing
        If listUpBeam.Count > 3 And listDownBeam.Count > 3 Then
            Dim Shell = New Shell()
            Dim Positions As List(Of Vector3D) = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(0))
            Positions.Add(listUpBeam.Item(1))
            Positions.Add(listUpBeam.Item(2))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(2))
            Positions.Add(listUpBeam.Item(3))
            Positions.Add(listUpBeam.Item(0))
            Tools.AddFace(Shell, Positions)
            'низ
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(0))
            Positions.Add(listDownBeam.Item(3))
            Positions.Add(listDownBeam.Item(2))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(2))
            Positions.Add(listDownBeam.Item(1))
            Positions.Add(listDownBeam.Item(0))
            Tools.AddFace(Shell, Positions)

            Positions = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(3))
            Positions.Add(listUpBeam.Item(2))
            Positions.Add(listDownBeam.Item(2))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(2))
            Positions.Add(listDownBeam.Item(3))
            Positions.Add(listUpBeam.Item(3))
            Tools.AddFace(Shell, Positions)

            Positions = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(0))
            Positions.Add(listDownBeam.Item(0))
            Positions.Add(listDownBeam.Item(1))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(1))
            Positions.Add(listUpBeam.Item(1))
            Positions.Add(listUpBeam.Item(0))
            Tools.AddFace(Shell, Positions)

            Positions = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(0))
            Positions.Add(listUpBeam.Item(3))
            Positions.Add(listDownBeam.Item(3))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(3))
            Positions.Add(listDownBeam.Item(0))
            Positions.Add(listUpBeam.Item(0))
            Tools.AddFace(Shell, Positions)

            Positions = New List(Of Vector3D)()
            Positions.Add(listUpBeam.Item(2))
            Positions.Add(listUpBeam.Item(1))
            Positions.Add(listDownBeam.Item(1))
            Tools.AddFace(Shell, Positions)
            Positions = New List(Of Vector3D)()
            Positions.Add(listDownBeam.Item(1))
            Positions.Add(listDownBeam.Item(2))
            Positions.Add(listUpBeam.Item(2))
            Tools.AddFace(Shell, Positions)
            Tools.Flip(Shell)
            Return Shell
        End If
    End Function
    'функция строит тело между двумя 3д полилиниями (в полилиниях одинаковое количество точек)
    Public Shared Function FuncCreateSolidByTwoPolyline3d(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal polyTopPolyline3d As DwgPolyline3D, ByVal polyBottomPolyline3d As DwgPolyline3D, Optional ByVal centerTopLine As DwgLine = Nothing) As Shell
        'рисуем тело насадки основное
        Dim shellNozzle As Shell = Nothing
        Dim pointBottom As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        Dim pointTop As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        If IsNothing(polyTopPolyline3d) = True Then Return Nothing
        If IsNothing(polyBottomPolyline3d) = True Then Return Nothing
        If polyBottomPolyline3d.Count < 4 Then Return Nothing
        Dim tempPolyTopPolyline3d As DwgPolyline3D = polyTopPolyline3d.Clone
        Dim tempPolyBottomPolyline3d As DwgPolyline3D = polyBottomPolyline3d.Clone
        If IsNothing(centerTopLine) = False Then
            If centerTopLine.Length > 0 Then
                If polyTopPolyline3d.Count > 1 Then
Line1:
                    For i As Integer = 0 To tempPolyTopPolyline3d.Count - 1
                        If (tempPolyTopPolyline3d.Item(i).Pos - centerTopLine.StartPoint.Pos).Length <= 0.001 Then
                            tempPolyTopPolyline3d.Remove(tempPolyTopPolyline3d.Item(i))
                            GoTo Line1
                        ElseIf (tempPolyTopPolyline3d.Item(i).Pos - centerTopLine.EndPoint.Pos).Length <= 0.001 Then
                            tempPolyTopPolyline3d.Remove(tempPolyTopPolyline3d.Item(i))
                            GoTo Line1
                        End If
                    Next i
                End If
                If polyBottomPolyline3d.Count > 1 Then
Line2:
                    For i As Integer = 0 To tempPolyBottomPolyline3d.Count - 1
                        If (tempPolyBottomPolyline3d.Item(i).Pos - centerTopLine.StartPoint.Pos).Length = 0 Then
                            tempPolyBottomPolyline3d.Remove(tempPolyBottomPolyline3d.Item(i))
                            GoTo Line2
                        ElseIf (tempPolyBottomPolyline3d.Item(i).Pos - centerTopLine.EndPoint.Pos).Length = 0 Then
                            tempPolyBottomPolyline3d.Remove(tempPolyBottomPolyline3d.Item(i))
                            GoTo Line2
                        End If
                    Next i
                End If
            End If
        End If
        If tempPolyTopPolyline3d.Count < 4 Then Return Nothing
        If (tempPolyTopPolyline3d.Count Mod 2) <> 0 Then Return Nothing
        If tempPolyBottomPolyline3d.Count < 4 Then Return Nothing
        If (tempPolyBottomPolyline3d.Count Mod 2) <> 0 Then Return Nothing
        If tempPolyTopPolyline3d.Count <> tempPolyBottomPolyline3d.Count Then Return Nothing
        Dim indexVertTop As Integer = 0
        Dim indexVertBottom As Integer = 0
        For i As Integer = 0 To tempPolyTopPolyline3d.Count / 2 - 2
            'забираем 4 точки по верхней 3д полилинии
            Try
                Dim leftTopPt1 As Vector3D = tempPolyTopPolyline3d.Item(i)
                Dim leftTopPt2 As Vector3D = tempPolyTopPolyline3d.Item(i + 1)
                Dim rightTopPt2 As Vector3D = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 2 - i)
                Dim rightTopPt1 As Vector3D = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 1 - i)
                'забираем 4 точки по нижней 3д полилинии
                Dim leftBottomPt1 As Vector3D = tempPolyBottomPolyline3d.Item(i)
                Dim leftBottomPt2 As Vector3D = tempPolyBottomPolyline3d.Item(i + 1)
                Dim rightBottomPt2 As Vector3D = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 2 - i)
                Dim rightBottomPt1 As Vector3D = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 1 - i)
                'ищем пересечение с линией шкафной стенки
                Dim boolFindSect1 As Boolean = False
                Dim boolFindSect2 As Boolean = False
                Dim intersectPoint1 As Vector2D = Nothing
                Dim intersectPoint2 As Vector2D = Nothing
                If IsNothing(centerTopLine) = False Then
                    If centerTopLine.Length > 0 Then
                        intersectPoint1 = MathFunction.FuncFindLineIntersection(leftTopPt1.Pos, rightTopPt1.Pos, centerTopLine.StartPoint.Pos, centerTopLine.EndPoint.Pos, boolFindSect1)
                        intersectPoint2 = MathFunction.FuncFindLineIntersection(leftTopPt2.Pos, rightTopPt2.Pos, centerTopLine.StartPoint.Pos, centerTopLine.EndPoint.Pos, boolFindSect2)
                    End If
                End If
                If boolFindSect1 = True And boolFindSect2 = True Then
                    'есть серединная линия, делим фигуру на 2 части
                    'вычисляем отметки по верху пересечения серединной линии
                    Dim elevTopMiddle1 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint1)
                    Dim elevTopMiddle2 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint2)
                    'вичисляем отметки по низу серединной линии
                    Dim elevBottomMiddle1 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt1, rightBottomPt1, intersectPoint1)
                    Dim elevBottomMiddle2 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt2, rightBottomPt2, intersectPoint2)
                    'составляем точки по верху фигуры
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevTopMiddle2))
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevTopMiddle1))

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevBottomMiddle2))
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevBottomMiddle1))
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                    'строим вторую половину
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevTopMiddle1))
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevTopMiddle2))
                    pointTop.Add(rightTopPt2)
                    pointTop.Add(rightTopPt1)

                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevBottomMiddle1))
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevBottomMiddle2))
                    pointBottom.Add(rightBottomPt2)
                    pointBottom.Add(rightBottomPt1)
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                Else
                    'составляем точки по верху фигуры
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(rightTopPt2)
                    pointTop.Add(rightTopPt1)

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(rightBottomPt2)
                    pointBottom.Add(rightBottomPt1)
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                End If
            Catch ex As System.Exception
                Return Nothing
            End Try
        Next i
        Return shellNozzle
    End Function
    Public Shared Function FuncCreateSolidByTwoNPolyline3d(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal polyTopPolyline3d As DwgPolyline3D, ByVal polyBottomPolyline3d As DwgPolyline3D, Optional ByVal centerTopLine As DwgLine = Nothing) As Shell
        'рисуем тело насадки основное
        Dim shellNozzle As Shell = Nothing
        Dim pointBottom As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        Dim pointTop As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        If IsNothing(polyTopPolyline3d) = True Then Return Nothing
        If IsNothing(polyBottomPolyline3d) = True Then Return Nothing
        If polyBottomPolyline3d.Count < 4 Then Return Nothing
        If polyTopPolyline3d.Count < 4 Then Return Nothing
        If (polyTopPolyline3d.Count Mod 2) <> 0 Then Return Nothing
        Dim indexVertTop As Integer = 0
        Dim indexVertBottom As Integer = 0
        For i As Integer = 0 To polyBottomPolyline3d.Count / 2 - 2
            'забираем 4 точки по верхней 3д полилинии
            Try
                'забираем 4 точки по нижней 3д полилинии
                Dim leftBottomPt1 As Vector3D = polyBottomPolyline3d.Item(i)
                Dim leftBottomPt2 As Vector3D = polyBottomPolyline3d.Item(i + 1)
                Dim rightBottomPt2 As Vector3D = polyBottomPolyline3d.Item(polyBottomPolyline3d.Count - 2 - i)
                Dim rightBottomPt1 As Vector3D = polyBottomPolyline3d.Item(polyBottomPolyline3d.Count - 1 - i)
                If (leftBottomPt1.Pos - leftBottomPt2.Pos).Length = 0 Then
                    Continue For
                End If
                If (rightBottomPt1.Pos - rightBottomPt2.Pos).Length = 0 Then
                    Continue For
                End If

                Dim leftTopPt1 As Vector3D = polyTopPolyline3d.Item(i)
                Dim leftTopPt2 As Vector3D = polyTopPolyline3d.Item(i + 1)
                Dim rightTopPt2 As Vector3D = polyTopPolyline3d.Item(polyTopPolyline3d.Count - 2 - i)
                Dim rightTopPt1 As Vector3D = polyTopPolyline3d.Item(polyTopPolyline3d.Count - 1 - i)
                For j As Integer = 0 To polyTopPolyline3d.Count - 1
                    Dim tempPt As Vector3D = polyTopPolyline3d.Item(j)
                    If (tempPt.Pos - leftBottomPt1.Pos).Length <= 0.001 Then
                        leftTopPt1 = tempPt
                    ElseIf (tempPt.Pos - leftBottomPt2.Pos).Length <= 0.001 Then
                        leftTopPt2 = tempPt
                    ElseIf (tempPt.Pos - rightBottomPt2.Pos).Length <= 0.001 Then
                        rightTopPt2 = tempPt
                    ElseIf (tempPt.Pos - rightBottomPt1.Pos).Length <= 0.001 Then
                        rightTopPt1 = tempPt
                    End If
                Next j
                'ищем пересечение с линией шкафной стенки
                Dim boolFindSect1 As Boolean = False
                Dim boolFindSect2 As Boolean = False
                Dim intersectPoint1 As Vector2D = Nothing
                Dim intersectPoint2 As Vector2D = Nothing
                If IsNothing(centerTopLine) = False Then
                    If centerTopLine.Length > 0 Then
                        intersectPoint1 = MathFunction.FuncFindLineIntersection(leftTopPt1.Pos, rightTopPt1.Pos, centerTopLine.StartPoint.Pos, centerTopLine.EndPoint.Pos, boolFindSect1)
                        intersectPoint2 = MathFunction.FuncFindLineIntersection(leftTopPt2.Pos, rightTopPt2.Pos, centerTopLine.StartPoint.Pos, centerTopLine.EndPoint.Pos, boolFindSect2)
                    End If
                End If
                If boolFindSect1 = True And boolFindSect2 = True Then
                    'есть серединная линия, делим фигуру на 2 части
                    'вычисляем отметки по верху пересечения серединной линии
                    Dim elevTopMiddle1 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint1)
                    Dim elevTopMiddle2 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint2)
                    'вичисляем отметки по низу серединной линии
                    Dim elevBottomMiddle1 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt1, rightBottomPt1, intersectPoint1)
                    Dim elevBottomMiddle2 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt2, rightBottomPt2, intersectPoint2)
                    'составляем точки по верху фигуры
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevTopMiddle2))
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevTopMiddle1))

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevBottomMiddle2))
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevBottomMiddle1))
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                    'строим вторую половину
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevTopMiddle1))
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevTopMiddle2))
                    pointTop.Add(rightTopPt2)
                    pointTop.Add(rightTopPt1)

                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevBottomMiddle1))
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevBottomMiddle2))
                    pointBottom.Add(rightBottomPt2)
                    pointBottom.Add(rightBottomPt1)
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                Else
                    'составляем точки по верху фигуры
                    pointTop = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointBottom = New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(rightTopPt2)
                    pointTop.Add(rightTopPt1)

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(rightBottomPt2)
                    pointBottom.Add(rightBottomPt1)
                    If IsNothing(shellNozzle) = True Then
                        shellNozzle = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                    Else
                        Dim shellNozzle1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(activDoc, pointBottom, pointTop)
                        shellNozzle = Tools.Union(shellNozzle, shellNozzle1)
                    End If
                End If
            Catch ex As System.Exception
                Return Nothing
            End Try
        Next i
        Return shellNozzle
    End Function
    '==============================================================================================================================
    'функция строит политело
    Public Shared Function FuncCreatePolyBoxByPolyline(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal listVertexUpPoint As List(Of Vector3D), ByVal offsetElevation As Double) As Shell
        FuncCreatePolyBoxByPolyline = Nothing
        Dim rezShell As Shell = New Shell
        Dim round As Integer = 3
        If IsNothing(listVertexUpPoint) = True Then Return Nothing
        'делаем новый список для вершин
        Dim polyList2d As List(Of Vector2D) = New List(Of Vector2D)
        For i As Integer = 0 To listVertexUpPoint.Count - 1
            polyList2d.Add(New Vector2D(listVertexUpPoint(i).Pos))
        Next
        'делаем триангуляцию
        'Dim rezTriangle As Dictionary(Of Integer, Triangle.Triangle) = New Dictionary(Of Integer, Triangle.Triangle)
        'Dim TriangleClass As Triangle = New Triangle
        'Dim result As String = TriangleClass.TriangulateClockwiseSimplePolygon(polyList2d, rezTriangle)
        'If result Like "OK" Then
        '    If rezTriangle.Count > 0 Then
        '        For i As Integer = 0 To rezTriangle.Count - 1
        '            Dim vertex3 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).A, 0)
        '            Dim vertex3d As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).A, 0)
        '            'находим высоту верщины
        '            For j As Integer = 0 To listVertexUpPoint.Count - 1
        '                If (listVertexUpPoint.Item(j).Pos - vertex3).Length < 0.002 Then
        '                    vertex3 = New Topomatic.Cad.Foundation.Vector3D(vertex3.Pos, listVertexUpPoint.Item(j).Z)
        '                    vertex3d = New Topomatic.Cad.Foundation.Vector3D(vertex3.Pos, listVertexUpPoint.Item(j).Z + offsetElevation)
        '                    Exit For
        '                End If
        '            Next j
        '            Dim vertex2 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).B, 0)
        '            Dim vertex2d As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).B, 0)
        '            'находим высоту верщины
        '            For j As Integer = 0 To listVertexUpPoint.Count - 1
        '                If (listVertexUpPoint.Item(j).Pos - vertex2).Length < 0.002 Then
        '                    vertex2 = New Topomatic.Cad.Foundation.Vector3D(vertex2.Pos, listVertexUpPoint.Item(j).Z)
        '                    vertex2d = New Topomatic.Cad.Foundation.Vector3D(vertex2.Pos, listVertexUpPoint.Item(j).Z + offsetElevation)
        '                    Exit For
        '                End If
        '            Next j
        '            Dim vertex1 As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).C, 0)
        '            Dim vertex1d As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(rezTriangle.Item(i).C, 0)
        '            'находим высоту верщины
        '            For j As Integer = 0 To listVertexUpPoint.Count - 1
        '                If (listVertexUpPoint.Item(j).Pos - vertex1).Length < 0.002 Then
        '                    vertex1 = New Topomatic.Cad.Foundation.Vector3D(vertex1.Pos, listVertexUpPoint.Item(j).Z)
        '                    vertex1d = New Topomatic.Cad.Foundation.Vector3D(vertex1.Pos, listVertexUpPoint.Item(j).Z + offsetElevation)
        '                    Exit For
        '                End If
        '            Next j
        '            'стоим тело
        '            Dim Shell = New Shell()
        '            'верхний треугольник
        '            Dim Positions As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex1d)
        '            Positions.Add(vertex2d)
        '            Positions.Add(vertex3d)
        '            Tools.AddFace(Shell, Positions)
        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex3)
        '            Positions.Add(vertex2)
        '            Positions.Add(vertex1)
        '            Tools.AddFace(Shell, Positions)

        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex1)
        '            Positions.Add(vertex2)
        '            Positions.Add(vertex2d)
        '            Tools.AddFace(Shell, Positions)
        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex2d)
        '            Positions.Add(vertex1d)
        '            Positions.Add(vertex1)
        '            Tools.AddFace(Shell, Positions)

        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex2)
        '            Positions.Add(vertex3)
        '            Positions.Add(vertex3d)
        '            Tools.AddFace(Shell, Positions)

        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex3d)
        '            Positions.Add(vertex2d)
        '            Positions.Add(vertex2)
        '            Tools.AddFace(Shell, Positions)

        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex3)
        '            Positions.Add(vertex1)
        '            Positions.Add(vertex1d)
        '            Tools.AddFace(Shell, Positions)
        '            Positions = New List(Of Topomatic.Cad.Foundation.Vector3D)()
        '            Positions.Add(vertex1d)
        '            Positions.Add(vertex3d)
        '            Positions.Add(vertex3)
        '            Tools.AddFace(Shell, Positions)

        '            If i = 0 Then
        '                rezShell = Shell
        '            Else
        '                rezShell = Tools.Union(rezShell, Shell)
        '            End If
        '        Next i
        '    End If
        'End If
        Return rezShell
    End Function
    Public Function TriangulateInContour(contourPoints As List(Of Point)) As Dictionary(Of Integer, List(Of Point))
        If contourPoints Is Nothing OrElse contourPoints.Count < 3 Then
            Throw New ArgumentException("Контур должен содержать как минимум 3 точки")
        End If

        ' Создаем полигон из точек контура
        Dim contourPolygon = CreatePolygonFromPoints(contourPoints)

        ' Создаем триангуляцию с ограничениями
        Dim triangulationBuilder As New ConformingDelaunayTriangulationBuilder()



        ' Устанавливаем ограничения (контур)
        Dim constraints As New GeometryCollection(New Geometry() {contourPolygon})
        triangulationBuilder.Constraints = constraints

        ' Выполняем триангуляцию
        Dim geometryFactory As New GeometryFactory()
        Dim triangulation = triangulationBuilder.GetTriangles(geometryFactory)

        ' Фильтруем треугольники, оставляя только внутри контура
        Return FilterTrianglesInContour(triangulation, contourPolygon)
    End Function

    Private Function CreatePolygonFromPoints(points As List(Of Point)) As Polygon
        Dim coordinates As New List(Of Coordinate)()

        ' Добавляем все точки контура
        For Each point In points
            coordinates.Add(New Coordinate(point.X, point.Y))
        Next

        ' Замыкаем полигон (добавляем первую точку в конец)
        If points.Count > 0 Then
            coordinates.Add(New Coordinate(points(0).X, points(0).Y))
        End If

        Dim geometryFactory As New GeometryFactory()
        Return geometryFactory.CreatePolygon(coordinates.ToArray())
    End Function

    ''' <summary>
    ''' Фильтрует треугольники, оставляя только те, что внутри контура
    ''' </summary>
    Private Function FilterTrianglesInContour(triangulation As GeometryCollection,
                                            contourPolygon As Polygon) As Dictionary(Of Integer, List(Of Point))

        Dim result As New Dictionary(Of Integer, List(Of Point))()
        Dim triangleIndex As Integer = 0

        For i As Integer = 0 To triangulation.NumGeometries - 1
            Dim triangle = DirectCast(triangulation.GetGeometryN(i), Polygon)

            ' Проверяем, что треугольник полностью внутри контура
            If contourPolygon.Contains(triangle) Then
                Dim vertices As New List(Of Point)()
                Dim coordinates = triangle.Coordinates

                ' Берем первые 3 координаты (вершины треугольника)
                For j As Integer = 0 To Math.Min(2, coordinates.Length - 1)
                    Dim coord = coordinates(j)
                    vertices.Add(New Point(coord.X, coord.Y))
                Next

                result.Add(triangleIndex, vertices)
                triangleIndex += 1
            End If
        Next

        Return result
    End Function
End Class
