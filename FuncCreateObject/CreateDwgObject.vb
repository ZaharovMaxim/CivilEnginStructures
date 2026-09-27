Imports System.IO
Imports System.Windows.Shapes
Imports NetTopologySuite.Geometries.Utilities
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Planchet.Entities
Imports Topomatic.Sfc
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Public Class CreateDwgObject
    Private _drawDocument As Drawing
    Public Sub New(DrawDocument As Drawing)
        _drawDocument = DrawDocument
    End Sub
    Public Property DrawDocument As Drawing
        Get
            Return _drawDocument
        End Get
        Set(value As Drawing)
            _drawDocument = value
        End Set
    End Property
    'РИСОВАНИЕ ПРИМИТИВОВ
    'создание полилинии из массива (0-координата X, 1- координата Y)
    Public Function createPolylineToArrayCoordinates(ByRef ArrayVertex2d As Double(,), Optional ByVal boolClosed As Boolean = False, Optional boolReverse As Boolean = False) As DwgPolyline
        Dim result As DwgPolyline = Nothing
        If IsNothing(DrawDocument) = False Then
            If IsArray(ArrayVertex2d) = True Then
                If ArrayVertex2d.GetUpperBound(1) > 0 Then
                    Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
                    result = New Topomatic.Dwg.Entities.DwgPolyline()
                    If boolReverse = False Then
                        For i As Integer = 0 To ArrayVertex2d.GetUpperBound(1)
                            rp.X = ArrayVertex2d(0, i)
                            rp.Y = ArrayVertex2d(1, i)
                            rp.Z = 0
                            Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                            Dim bulgeSegment As Single = 0
                            If ArrayVertex2d.GetUpperBound(0) > 1 Then
                                bulgeSegment = ArrayVertex2d(2, i)
                            End If
                            result.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
                        Next i
                    Else
                        For i As Integer = ArrayVertex2d.GetUpperBound(1) To 0 Step -1
                            rp.X = ArrayVertex2d(0, i)
                            rp.Y = ArrayVertex2d(1, i)
                            rp.Z = 0
                            Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                            Dim bulgeSegment As Single = 0
                            If ArrayVertex2d.GetUpperBound(0) > 1 Then
                                If i <> 0 Then
                                    bulgeSegment = Val(ArrayVertex2d(2, i - 1)) * -1
                                Else
                                    bulgeSegment = 0
                                End If
                            End If
                            result.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
                        Next i
                    End If
                    If IsNothing(result) = False Then
                        If result.Count > 1 Then
                            DrawDocument.ActiveSpace.Entities.Add(result)
                            If boolClosed = True Then
                                If result.Area > 0 Then
                                    result.Closed = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return result
    End Function
    'создание полилинии из списка координат точек
    Public Function createPolylineBeListPoint(ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As DwgPolyline
        Dim polyline As DwgPolyline = Nothing
        If IsNothing(DrawDocument) = False Then
            If points.Count > 1 Then
                Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
                polyline = New Topomatic.Dwg.Entities.DwgPolyline()
                For i As Integer = 0 To points.Count - 1
                    rp.X = points.Item(i).X
                    rp.Y = points.Item(i).Y
                    rp.Z = 0
                    Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                    polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                Next i
                DrawDocument.ActiveSpace.Entities.Add(polyline)
                If boolClosed = True Then
                    If polyline.Area > 0 Then
                        polyline.Closed = True
                    End If
                End If
            End If
        End If
        Return polyline
    End Function
    'создание полилинии из списка
    Public Shared Function createPolyline3dFromListCoord(ByVal ActivDocument As Drawing, ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As DwgPolyline3D
        createPolyline3dFromListCoord = Nothing
        If points.Count > 1 Then
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            Dim polyline As DwgPolyline3D = New Topomatic.Dwg.Entities.DwgPolyline3D()
            For i As Integer = 0 To points.Count - 1
                polyline.Add(points.Item(i))
            Next i

            If IsNothing(ActivDocument) = False Then
                ActivDocument.ActiveSpace.Add(polyline)
                If boolClosed = True Then
                    polyline.Closed = True
                End If
                Return polyline
            End If
        End If
    End Function
    'перерисовка полилинии из списка
    Public Shared Function FuncReDrawPolylineToListCoord(ByVal ActivDocument As Drawing, ByRef points As List(Of Vector3D), ByRef polyline As DwgPolyline, Optional boolClosed As Boolean = False) As Boolean
        FuncReDrawPolylineToListCoord = Nothing
        If points.Count > 1 Then
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            If ActivDocument.ActiveSpace.Entities.Contains(polyline) = False Then
                polyline = New Topomatic.Dwg.Entities.DwgPolyline()
            End If
            For i As Integer = 0 To points.Count - 1
                rp.X = points.Item(i).X
                rp.Y = points.Item(i).Y
                rp.Z = 0
                Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                If polyline.Count >= i + 1 Then
                    polyline.Item(i) = New Topomatic.Cad.Foundation.BugleVector2D(pos, 0)
                Else
                    polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
                End If
            Next i
            If boolClosed = True Then
                polyline.Closed = True
            End If
            If ActivDocument.ActiveSpace.Entities.Contains(polyline) = False Then
                ActivDocument.ActiveSpace.Entities.Add(polyline)
            End If
            Return True
        End If
    End Function
    'создание полилинии из Polyline2DCurve
    Public Shared Function FuncDrawPolylineToPolylineCurve(ByVal ActivDocument As Drawing, ByRef points As Polyline2DCurve, Optional ByVal boolClosed As Boolean = False) As DwgPolyline
        FuncDrawPolylineToPolylineCurve = Nothing
        If points.Count > 1 Then
            Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
            For i As Integer = 0 To points.Count - 1
                polyline.Add(points.Item(i))
            Next i

            If IsNothing(ActivDocument) = False Then
                ActivDocument.ActiveSpace.Add(polyline)
                If boolClosed = True Then
                    polyline.Closed = True
                End If
                Return polyline
            End If

        End If
    End Function
    'создание полилинии параллельной заданной
    Public Shared Function FuncDrawOffsetPolyline(ByVal ActivDocument As Drawing, ByVal polyline As DwgPolyline, ByVal offset As Double) As DwgPolyline
        FuncDrawOffsetPolyline = Nothing
        If IsNothing(polyline) = False Then
            If polyline.Length > 0 Then
                Dim acList As List(Of DwgEntity) = New List(Of DwgEntity)
                polyline.Offset(acList, offset)
                If acList.Count > 0 Then
                    For i As Integer = 0 To acList.Count - 1
                        Dim acEnt As DwgEntity = acList.ElementAt(i)
                        If TypeOf (acEnt) Is DwgPolyline Then
                            Dim acPoly As DwgPolyline = acEnt
                            If acPoly.Length > 0 Then
                                ActivDocument.ActiveSpace.Add(acPoly)
                                Return acPoly
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Function
    'создание 3d полилинии из массива
    Public Shared Function FuncDrawPolyline3DToArrayCoord(ByVal ActivDocument As Drawing, ByRef ArrayVertex3d As Double(,), Optional ByVal boolClosed As Boolean = False) As DwgPolyline3D
        FuncDrawPolyline3DToArrayCoord = Nothing
        If IsArray(ArrayVertex3d) = True Then
            If ArrayVertex3d.GetUpperBound(1) > 0 Then
                Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
                Dim polyline As DwgPolyline3D = New Topomatic.Dwg.Entities.DwgPolyline3D()
                For i As Integer = 0 To ArrayVertex3d.GetUpperBound(1)
                    rp.X = ArrayVertex3d(0, i)
                    rp.Y = ArrayVertex3d(1, i)
                    rp.Z = ArrayVertex3d(2, i)
                    Dim pos As Vector3D = New Topomatic.Cad.Foundation.Vector3D(rp.X, rp.Y, rp.Z)
                    polyline.Add(pos)
                Next i

                If IsNothing(ActivDocument) = False Then
                    ActivDocument.ActiveSpace.Add(polyline)
                    If boolClosed = True Then
                        polyline.Closed = True
                    End If
                    Return polyline
                End If
            End If
        End If
    End Function
    'создание 3d полилинии из списка
    Public Function createPolyline3D(ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As DwgPolyline3D
        Dim result As DwgPolyline3D = Nothing
        Dim userDrawing As Drawing = DrawDocument
        If IsNothing(userDrawing) = False Then
            If points.Count > 1 Then
                result = New Topomatic.Dwg.Entities.DwgPolyline3D()
                Dim prevPoint As Vector3D = New Vector3D(-1, -1, -1)
                For Each pos As Vector3D In points
                    If result.Count > 0 Then
                        If (pos - prevPoint).Length = 0 Then
                            Continue For
                        End If
                    End If
                    result.Add(pos)
                    prevPoint = pos
                Next
                If result.Count > 1 Then
                    userDrawing.ActiveSpace.Entities.Add(result)
                    If boolClosed = True And result.Count > 2 Then
                        result.Closed = True
                    End If
                Else
                    result = Nothing
                End If
            End If
        End If
        Return result
    End Function
    'обновление 3d полилинии из списка
    Public Function reDrawPolyline3D(ByRef polyline3d As DwgPolyline3D, ByRef points As List(Of Vector3D)) As Boolean
        If IsNothing(polyline3d) = True Then Return False
        If IsNothing(points) = True Then Return False
        If points.Count < 2 Then
            Return False
        End If
        If points.Count > 1 Then
            polyline3d.Clear()
            Dim prevPoint As Vector3D = New Vector3D(-1, -1, -1)
            For Each pos As Vector3D In points
                If polyline3d.Count > 0 Then
                    If (pos - prevPoint).Length = 0 Then
                        Continue For
                    End If
                End If
                polyline3d.Add(pos)
                prevPoint = pos
            Next
        Else
            Return False
        End If
        Return True
    End Function
    'создание структурной линии на ЦММ
    Public Shared Function FuncCreateStructureLine(ByVal acSurface As Surface, ByRef ArrayVertex3d As Double(,), Optional codeLine As Integer = 0, Optional ByVal boolClosed As Boolean = False, Optional isSituation As Boolean = True) As StructureLine
        FuncCreateStructureLine = Nothing
        Try
            Dim editor As Topomatic.Sfc.PointEditor = New Topomatic.Sfc.PointEditor(acSurface)
            Dim tempNewStructureLine As Topomatic.Sfc.StructureLine = New Topomatic.Sfc.StructureLine
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            If IsNothing(ArrayVertex3d) = False Then
                If ArrayVertex3d.GetUpperBound(1) > 0 Then
                    For i As Integer = 0 To ArrayVertex3d.GetUpperBound(1)
                        rp.X = ArrayVertex3d(0, i)
                        rp.Y = ArrayVertex3d(1, i)
                        rp.Z = ArrayVertex3d(2, i)
                        Dim point As Topomatic.Sfc.SurfacePoint = New Topomatic.Sfc.SurfacePoint(rp)
                        If isSituation = True Then
                            point.IsSituation = True
                        End If
                        Dim point_index As Integer = editor.Add(point)
                        tempNewStructureLine.Add(point_index)
                    Next i
                    acSurface.StructureLines.Add(tempNewStructureLine)
                    If boolClosed = True Then
                        tempNewStructureLine.IsClosed = True
                    End If
                    If isSituation = True Then
                        tempNewStructureLine.IsSituation = True
                    Else
                        tempNewStructureLine.IsSituation = False
                    End If
                    tempNewStructureLine.LinearCode = codeLine
                End If
            End If
            Return tempNewStructureLine
        Catch ex As Exception
        End Try
    End Function
    'создание структурной линии на ЦММ из полилинии
    Public Shared Function FuncCreateStructureLineByPolyline(ByVal acSurface As Surface, ByVal polyline As DwgPolyline, Optional ByVal code As Integer = 0, Optional ByVal desk As String = "", Optional ByVal offsetElevation As Double = 0) As StructureLine
        FuncCreateStructureLineByPolyline = Nothing
        Try
            Dim editor As Topomatic.Sfc.PointEditor = New Topomatic.Sfc.PointEditor(acSurface)
            Dim tempNewStructureLine As Topomatic.Sfc.StructureLine = New Topomatic.Sfc.StructureLine
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            Dim elev As Double = 0
            If IsNothing(polyline) = False Then
                If polyline.Count > 1 Then
                    For i As Integer = 0 To polyline.Count - 1
                        rp.X = polyline.Item(i).Vertex.X
                        rp.Y = polyline.Item(i).Vertex.Y
                        Try
                            rp.Z = acSurface.GetElevation(polyline.Item(i).Vertex)
                            elev = rp.Z + offsetElevation
                        Catch ex As System.ArgumentOutOfRangeException
                            rp.Z = elev
                        End Try
                        Dim point As Topomatic.Sfc.SurfacePoint = New Topomatic.Sfc.SurfacePoint(rp)
                        Dim point_index As Integer = editor.Add(point)
                        tempNewStructureLine.Add(point_index)
                    Next i
                    acSurface.StructureLines.Add(tempNewStructureLine)
                    tempNewStructureLine.LinearCode = code
                    tempNewStructureLine.Description = desk
                    If offsetElevation = 0 Then
                        tempNewStructureLine.IsLimitation = True
                    Else
                        tempNewStructureLine.IsLimitation = False
                    End If
                End If
            End If
            Return tempNewStructureLine
        Catch ex As Exception
        End Try
    End Function
    'обновление структурной линии на ЦММ
    Public Shared Function FuncReDrawStructuresLine(ByVal acStructLine As StructureLine, ByVal arrayVertex As Double(,), Optional offsetElevation As Double = 0, Optional isSituation As Boolean = True) As Boolean
        FuncReDrawStructuresLine = False
        Try
            Dim acSurface As Surface = acStructLine.Surface
            Dim editor As Topomatic.Sfc.PointEditor = New Topomatic.Sfc.PointEditor(acSurface)
            Dim countVertex As Integer = arrayVertex.GetUpperBound(1) + 1
            If acStructLine.Count > countVertex Then
                countVertex = acStructLine.Count
            End If
            For i As Integer = 0 To countVertex - 1
                If i <= arrayVertex.GetUpperBound(1) Then
                    Dim oldIndVertex As Integer = acStructLine.ElementAt(i).Index
                    Dim sPoint = acSurface.Points.Item(oldIndVertex)
                    Dim newVertex As Vector3D = New Vector3D(arrayVertex(0, i), arrayVertex(1, i), arrayVertex(2, i) + offsetElevation)
                    sPoint.Vertex = newVertex
                    editor.SetVertex(oldIndVertex, newVertex)
                Else
                    acStructLine.RemoveAt(i)
                End If
            Next i
            If isSituation = True Then
                acStructLine.IsSituation = True
            Else
                acStructLine.IsSituation = False
                acSurface.BeginUpdate()
                If (acSurface.Style.Dynamic) Then
                    SurfaceTools.UpdateLimitations(acStructLine)
                    SurfaceTools.CheckStructureLinesCross(acSurface, acSurface.StructureLines.IndexOf(acStructLine))
                End If
                acSurface.EndUpdate()
            End If
            Return True
        Catch ex As Exception
        End Try
    End Function
    'создать откос
    Public Shared Function FuncDrawSlope(ByVal ActivDocument As Drawing, ByVal IPolySourse As Polyline3D, ByVal IPolyDest As Polyline3D) As DwgSlope
        FuncDrawSlope = Nothing
        Try
            Dim tempSlope As DwgSlope = New DwgSlope()
            If IsNothing(IPolySourse) = True Then Exit Function
            If IsNothing(IPolyDest) = True Then Exit Function
            tempSlope.AssignSourcePolyline(IPolySourse)
            tempSlope.AssignDestPolyline(IPolyDest)
            ActivDocument.ActiveSpace.Add(tempSlope)
            Return tempSlope
        Catch ex As System.Exception
        End Try
    End Function
    'функция вставляет прямоугольник
    Public Shared Function CreateRotatedRectangle(ByVal centerPoint As Topomatic.Cad.Foundation.Vector2D, ByVal Length As Double, width As Double, rotation As Double) As List(Of Topomatic.Cad.Foundation.Vector2D)
        Dim points As New List(Of Topomatic.Cad.Foundation.Vector2D)
        ' Половины размеров
        Dim halfLength As Double = Length / 2.0F
        Dim halfWidth As Double = width / 2.0F
        ' Исходные углы прямоугольника (до поворота, относительно центра)
        Dim corners As Topomatic.Cad.Foundation.Vector2D() = {
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, -halfWidth),  ' Верхний левый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, -halfWidth),   ' Верхний правый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, halfWidth),    ' Нижний правый
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, halfWidth)    ' Нижний левый
        }
        ' Поворачиваем и смещаем каждую точку
        For Each corner As Topomatic.Cad.Foundation.Vector2D In corners
            ' Поворот точки
            Dim rotatedX As Double = corner.X * Math.Cos(rotation) - corner.Y * Math.Sin(rotation)
            Dim rotatedY As Double = corner.X * Math.Sin(rotation) + corner.Y * Math.Cos(rotation)
            ' Смещение к центру
            Dim finalPoint As New Topomatic.Cad.Foundation.Vector2D(centerPoint.X + CSng(rotatedX), centerPoint.Y + CSng(rotatedY))
            points.Add(finalPoint)
        Next
        Return points
    End Function
    'функция рисует окружность
    Public Function CreateCircle(ByVal centerPoint As Topomatic.Cad.Foundation.Vector3D, ByVal diameter As Double) As DwgCircle
        Dim result As DwgCircle = Nothing
        If diameter > 0 Then
            result = New DwgCircle()
            result.Center = centerPoint
            result.Diametr = diameter
            If IsNothing(DrawDocument) = False Then
                DrawDocument.ActiveSpace.Entities.Add(result)
            End If
        End If
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'РИСОВАНИЕ 3Д ТЕЛ
    'функция строит тело между двумя 3д полилиниями (в полилиниях одинаковое количество точек)
    Public Function createSolid3DByTwoPolylines3d(ByVal polyTopPolyline3d As DwgPolyline3D, ByVal polyBottomPolyline3d As DwgPolyline3D, Optional ByVal centerTopLine As DwgLine = Nothing) As Shell
        'рисуем тело насадки основное
        Dim shellObject As Shell = Nothing
        If IsNothing(polyTopPolyline3d) = True Then Return Nothing
        If IsNothing(polyBottomPolyline3d) = True Then Return Nothing
        If polyTopPolyline3d.Length = 0 Then Return Nothing
        If polyBottomPolyline3d.Length = 0 Then Return Nothing
        If polyTopPolyline3d.Count < 4 Then Return Nothing
        If polyBottomPolyline3d.Count < 4 Then Return Nothing
        'прогоняем точки полилинии и удаляем moddlePT
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
        'берем точки попарно слева и справа
        Dim indexVertTop As Integer = 0
        Dim indexVertBottom As Integer = 0
        For i As Integer = 0 To tempPolyTopPolyline3d.Count / 2 - 2
            'забираем 4 точки по верхней 3д полилинии
            Try
                'точки по верху
                Dim leftTopPt1 As Vector3D = tempPolyTopPolyline3d.Item(indexVertTop) 'первая левая точка
                Dim leftTopPt2 As Vector3D = tempPolyTopPolyline3d.Item(indexVertTop + 1) 'вторая левая точка
                Dim rightTopPt2 As Vector3D = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 2 - indexVertTop) 'предпоследняя правая точка
                Dim rightTopPt1 As Vector3D = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 1 - indexVertTop) 'последняя правая точка
                If (leftTopPt1.Pos - leftTopPt2.Pos).Length <= 0.001 Then
                    If (rightTopPt2.Pos - rightTopPt1.Pos).Length <= 0.001 Then
                        indexVertTop += 1
                        leftTopPt1 = tempPolyTopPolyline3d.Item(indexVertTop) 'первая левая точка
                        leftTopPt2 = tempPolyTopPolyline3d.Item(indexVertTop + 1) 'вторая левая точка
                        rightTopPt2 = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 2 - indexVertTop) 'предпоследняя правая точка
                        rightTopPt1 = tempPolyTopPolyline3d.Item(tempPolyTopPolyline3d.Count - 1 - indexVertTop) 'последняя правая точка
                    End If
                End If
                'забираем 4 точки по нижней 3д полилинии
                Dim leftBottomPt1 As Vector3D = tempPolyBottomPolyline3d.Item(indexVertBottom) 'первая левая точка
                Dim leftBottomPt2 As Vector3D = tempPolyBottomPolyline3d.Item(indexVertBottom + 1) 'вторая левая точка
                Dim rightBottomPt2 As Vector3D = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 2 - indexVertBottom) 'предпоследняя правая точка
                Dim rightBottomPt1 As Vector3D = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 1 - indexVertBottom) 'последняя правая точка
                If (leftBottomPt1.Pos - leftBottomPt2.Pos).Length <= 0.001 Then
                    If (rightBottomPt2.Pos - rightBottomPt1.Pos).Length <= 0.001 Then
                        indexVertBottom += 1
                        leftBottomPt1 = tempPolyBottomPolyline3d.Item(indexVertBottom) 'первая левая точка
                        leftBottomPt2 = tempPolyBottomPolyline3d.Item(indexVertBottom + 1) 'вторая левая точка
                        rightBottomPt2 = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 2 - indexVertBottom) 'предпоследняя правая точка
                        rightBottomPt1 = tempPolyBottomPolyline3d.Item(tempPolyBottomPolyline3d.Count - 1 - indexVertBottom) 'последняя правая точка
                    End If
                End If
                'ищем пересечение с осевой линией
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
                'пересечение есть
                If boolFindSect1 = True And boolFindSect2 = True Then
                    'есть серединная линия, делим фигуру на 2 части
                    'вычисляем отметки по верху пересечения серединной линии
                    Dim elevTopMiddle1 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint1)
                    Dim elevTopMiddle2 As Double = MathFunction.FuncCalcElevationByLine(centerTopLine.StartPoint, centerTopLine.EndPoint, intersectPoint2)
                    'вичисляем отметки по низу серединной линии
                    Dim elevBottomMiddle1 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt1, rightBottomPt1, intersectPoint1)
                    Dim elevBottomMiddle2 As Double = MathFunction.FuncCalcElevationByLine(leftBottomPt2, rightBottomPt2, intersectPoint2)
                    'составляем точки по верху фигуры
                    Dim pointTop As New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    Dim pointBottom As New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevTopMiddle2))
                    pointTop.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevTopMiddle1))

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint2, elevBottomMiddle2))
                    pointBottom.Add(New Cad.Foundation.Vector3D(intersectPoint1, elevBottomMiddle1))
                    If IsNothing(shellObject) = True Then
                        shellObject = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                    Else
                        Dim shellObject1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                        shellObject = Tools.Union(shellObject, shellObject1)
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
                    If IsNothing(shellObject) = True Then
                        shellObject = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                    Else
                        Dim shellObject1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                        shellObject = Tools.Union(shellObject, shellObject1)
                    End If
                Else
                    'составляем точки по верху фигуры
                    Dim pointTop As New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    Dim pointBottom As New List(Of Topomatic.Cad.Foundation.Vector3D)()
                    pointTop.Add(leftTopPt1)
                    pointTop.Add(leftTopPt2)
                    pointTop.Add(rightTopPt2)
                    pointTop.Add(rightTopPt1)

                    pointBottom.Add(leftBottomPt1)
                    pointBottom.Add(leftBottomPt2)
                    pointBottom.Add(rightBottomPt2)
                    pointBottom.Add(rightBottomPt1)
                    If IsNothing(shellObject) = True Then
                        shellObject = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                    Else
                        Dim shellObject1 As Shell = FuncModeling3d.FuncCreateBoxByPoints(DrawDocument, pointBottom, pointTop)
                        shellObject = Tools.Union(shellObject, shellObject1)
                    End If
                End If
                indexVertTop += 1
                indexVertBottom += 1
            Catch ex As System.Exception
                Return Nothing
            End Try
        Next i
        Return shellObject
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'аттнотации
    '================================================================================================
    'создание многострочного текста
    Public Shared Function FuncDrawMText(ByVal ActivDocument As Drawing, ByVal Str As String, ByVal InsertVertex As Vector3D, Optional ByVal atPoint As AttachmentPoint = AttachmentPoint.BottomLeft) As DwgMText
        FuncDrawMText = Nothing
        Try
            FuncDrawMText = New Topomatic.Dwg.Entities.DwgMText
            FuncDrawMText.Content = Str
            FuncDrawMText.Position = InsertVertex
            FuncDrawMText.AttachmentPoint = atPoint
            ActivDocument.ActiveSpace.Add(FuncDrawMText)
        Catch ex As Exception
        End Try
    End Function
    '================================================================================================
    'создание выноски с текстом
    Public Shared Function FuncDrawLeader(ByVal ActivDocument As Drawing, ByVal Str As String, ByVal InsertFirstVertex As Topomatic.Cad.Foundation.Vector2D, ByVal InserSecondVertex As Topomatic.Cad.Foundation.Vector2D, Optional annoScale As Double = 1, Optional hText As Double = 2, Optional textStyle As DwgStyle = Nothing, Optional nameArrow As String = "", Optional arrowSize As Double = 2) As DwgLeader
        FuncDrawLeader = Nothing
        Try
            FuncDrawLeader = New Topomatic.Dwg.Entities.DwgLeader
            FuncDrawLeader.Add(InsertFirstVertex)
            FuncDrawLeader.Position = InserSecondVertex
            FuncDrawLeader.Content = Str
            If IsNothing(textStyle) = False Then
                FuncDrawLeader.Style = textStyle
            End If
            FuncDrawLeader.Height = hText
            FuncDrawLeader.Annotative = True
            'FuncDrawLeader.AnnotationScale = annoScale
            If nameArrow Like "" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.None
            ElseIf nameArrow Like "_ArchTick" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.ArchTick
            ElseIf nameArrow Like "_BoxBlank" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.BoxBlank
            ElseIf nameArrow Like "_BoxFilled" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.BoxFilled
            ElseIf nameArrow Like "_Closed" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Closed
            ElseIf nameArrow Like "_ClosedBlank" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.ClosedBlank
            ElseIf nameArrow Like "_DatumBlank" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.DatumBlank
            ElseIf nameArrow Like "_DatumFilled" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.DatumFilled
            ElseIf nameArrow Like "_Default" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Default
            ElseIf nameArrow Like "_Dot" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Dot
            ElseIf nameArrow Like "_DotBlank" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.DotBlank
            ElseIf nameArrow Like "_DotSmall" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.DotSmall
            ElseIf nameArrow Like "_Integral" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Integral
            ElseIf nameArrow Like "_None" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.None
            ElseIf nameArrow Like "_Oblique" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Oblique
            ElseIf nameArrow Like "_Open" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Open
            ElseIf nameArrow Like "_Open30" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Open30
            ElseIf nameArrow Like "_Open90" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Open90
            ElseIf nameArrow Like "_Origin" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Origin
            ElseIf nameArrow Like "_Origin2" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Origin2
            ElseIf nameArrow Like "_Small" Then
                FuncDrawLeader.ArrowheadType = AcDimArrowheadType.Small
            Else
                Dim FindBlock As DwgBlock = FuncReadBlockByName(ActivDocument, nameArrow)
                If IsNothing(FindBlock) = False Then
                    FuncDrawLeader.ArrowheadType = AcDimArrowheadType.UserDefined
                    FuncDrawLeader.ArrowheadName = nameArrow
                Else
                    FuncDrawLeader.ArrowheadType = AcDimArrowheadType.None
                End If
            End If
            FuncDrawLeader.ArrowheadSize = arrowSize * annoScale
            ActivDocument.ActiveSpace.Add(FuncDrawLeader)
        Catch ex As Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'учассткм м штриховки
    '================================================================================================
    'создание штриховки
    Public Function createHatchToPolyline(ByVal polyline As Topomatic.Dwg.Entities.DwgPolyline, ByVal nameHatch As String) As DwgHatch
        Dim result As DwgHatch = Nothing
        Dim userDrawing As Drawing = DrawDocument
        If IsNothing(polyline) = True Then
            If polyline.Area > 0 Then
                Try
                    Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
                    For Each vertex As BugleVector2D In polyline
                        BoundaryPath.Add(vertex)
                    Next
                    result = New Topomatic.Dwg.Entities.DwgHatch()
                    result.BoundaryPath.Add(BoundaryPath)
                    result.PatternName = nameHatch
                    result.Elevation = polyline.Elevation
                    userDrawing.ActiveSpace.Entities.Add(result)
                Catch ex As System.Exception
                    result = Nothing
                End Try
            End If
        End If
        Return result
    End Function
    Public Function createHatchToPolyline3d(ByVal polyline3d As Topomatic.Dwg.Entities.DwgPolyline3D, ByVal nameHatch As String) As DwgHatch
        Dim result As DwgHatch = Nothing
        Dim userDrawing As Drawing = DrawDocument
        If IsNothing(polyline3d) = False Then
            If polyline3d.Area > 0 Then
                Try
                    Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
                    Dim maxPoint As Double = -9999999
                    Dim minPoint As Double = 9999999
                    For i As Integer = 0 To polyline3d.Count - 1
                        Dim vertex As Topomatic.Cad.Foundation.Vector3D = polyline3d.Item(i)
                        BoundaryPath.Add(New BugleVector2D(vertex.Pos, 0))
                        If vertex.Z < minPoint Then minPoint = vertex.Z
                        If vertex.Z > maxPoint Then maxPoint = vertex.Z
                    Next i
                    Dim elevationHatch As Double = (maxPoint + minPoint) / 2
                    If BoundaryPath.Count > 2 Then
                        result = New DwgHatch
                        result.BoundaryPath.Add(BoundaryPath)
                        result.PatternName = nameHatch
                        result.Elevation = elevationHatch
                        userDrawing.ActiveSpace.Entities.Add(result)
                    End If
                Catch ex As System.Exception
                    result = Nothing
                End Try
            End If
        End If
        Return result
    End Function
    '================================================================================================
    'обновление штриховки
    Public Function reDrawHatchByPolyline(ByVal polyline As Topomatic.Dwg.Entities.DwgPolyline, ByRef acHatch As DwgHatch) As Boolean
        Try
            If IsNothing(polyline) = True Then Return False
            If polyline.Count < 3 Then Return False
            If IsNothing(acHatch) = True Then Return False
            acHatch.BoundaryPath.Clear()
            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
            For Each vertex As BugleVector2D In polyline
                BoundaryPath.Add(vertex)
            Next
            acHatch.BoundaryPath.Add(BoundaryPath)
            acHatch.Elevation = polyline.Elevation
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Function reDrawHatchByPolyline3d(ByVal polyline3d As Topomatic.Dwg.Entities.DwgPolyline3D, ByRef acHatch As DwgHatch) As Boolean
        Try
            If IsNothing(polyline3d) = True Then Return False
            If polyline3d.Count < 3 Then Return False
            If IsNothing(acHatch) = True Then Return False
            acHatch.BoundaryPath.Clear()
            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
            Dim maxPoint As Double = -9999999
            Dim minPoint As Double = 9999999
            For i As Integer = 0 To polyline3d.Count - 1
                Dim vertex As Topomatic.Cad.Foundation.Vector3D = polyline3d.Item(i)
                BoundaryPath.Add(New BugleVector2D(vertex.Pos, 0))
                If vertex.Z < minPoint Then minPoint = vertex.Z
                If vertex.Z > maxPoint Then maxPoint = vertex.Z
            Next
            Dim elevationHatch As Double = (maxPoint + minPoint) / 2
            acHatch.BoundaryPath.Add(BoundaryPath)
            acHatch.Elevation = elevationHatch
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    '================================================================================================
    'преобразование Участка Робур в штриховку

    'СОЗДАНИЕ ОБЪЕКТОВ (слой, тип линии, стиль и т.п
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать слой
    Public Shared Function FuncAddLayer(ByVal ActivDocument As Drawing, ByVal Name As String, Optional ByVal Indcolor As Integer = 7, Optional ByVal NameTypeLine As String = "", Optional ValLineWight As Integer = 20, Optional ByVal boolVisible As Boolean = True) As DwgLayer
        FuncAddLayer = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Layers.Count - 1
                If ActivDocument.Layers.Item(i).Name Like Name.Trim Then
                    FuncAddLayer = ActivDocument.Layers.Item(i)
                    Exit For
                End If
            Next i

            FuncAddLayer = ActivDocument.Layers.Add(Name.Trim)
            'назначаем цвет
            Indcolor = Math.Abs(Indcolor)
            If Indcolor > 249 Then
                Indcolor = 256
            End If
            Dim CColor As CadColor = New CadColor(Indcolor)
            FuncAddLayer.Color = CColor
            'назначаем тип линии слою
            Dim LineTypes As DwgLinetypes = ActivDocument.Linetypes
            Dim LineType As DwgLinetype = Nothing
            For Each Ltype As DwgLinetype In LineTypes
                Dim NameLineType As String = Ltype.Name
                If NameLineType Like NameTypeLine Then
                    FuncAddLayer.Linetype = Ltype
                    Exit For
                End If
            Next
            'назначаем вес линиям
            FuncAddLayer.Lineweight = ValLineWight
            'назначаем видимость слою
            If boolVisible = True Then
                FuncAddLayer.Visible = True
            Else
                FuncAddLayer.Visible = False
            End If
            Return FuncAddLayer
        Catch ex As System.ArgumentOutOfRangeException
        End Try
    End Function

    'СВОЙСТВА ОБЪЕКТА (слой, тип линии, стиль и т.п
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Поменять цвет объекта
    Public Shared Function FuncSetColorObject(ByVal ActivDocument As Drawing, ByVal acEntity As Topomatic.Dwg.Entities.DwgEntity, ByVal color1 As Integer, ByVal color2 As Integer, ByVal color3 As Integer) As Boolean
        FuncSetColorObject = False
        Try
            If color1 < 0 Then
                color1 = 0
            ElseIf color1 > 255 Then
                color1 = 255
            End If
            If color2 < 0 Then
                color2 = 0
            ElseIf color2 > 255 Then
                color2 = 255
            End If
            If color3 < 0 Then
                color3 = 0
            ElseIf color3 > 255 Then
                color3 = 255
            End If
            Dim colorUse As Topomatic.Cad.Foundation.CadColor = New Topomatic.Cad.Foundation.CadColor(System.Drawing.Color.FromArgb(color1, color2, color3))
            acEntity.Color = colorUse
            FuncSetColorObject = True
        Catch ex As Exception
        End Try
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'блоки
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает имена всех определений блока
    Public Shared Function FuncReadNameInsertBlocks(ByVal ActivDocument As Drawing, ByRef ArrayRezult As String()) As Boolean
        FuncReadNameInsertBlocks = False
        Dim countArrayRezult As Integer = 0
        If IsArray(ArrayRezult) = True Then
            countArrayRezult = ArrayRezult.Length
        End If
        Dim Dwgblks As DwgBlocks = ActivDocument.Blocks
        If Dwgblks.Count > 1 Then
            For Each dwgblk As DwgBlock In Dwgblks
                If dwgblk.IsHiden = False Then
                    ReDim Preserve ArrayRezult(countArrayRezult)
                    ArrayRezult(countArrayRezult) = dwgblk.Name
                    countArrayRezult += 1
                End If

            Next
        End If
        FuncReadNameInsertBlocks = True
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает имена всех определений блока
    Public Shared Function FuncReadInsertBlock(ByVal ActivDocument As Drawing, ByRef ArrayRezult As String(,)) As Boolean
        FuncReadInsertBlock = False
        Dim countArrayRezult As Integer = 0
        If IsArray(ArrayRezult) = True Then
            countArrayRezult = ArrayRezult.GetUpperBound(1)
        End If
        Dim Dwgblks As DwgBlocks = ActivDocument.Blocks
        If Dwgblks.Count > 1 Then
            For Each dwgblk As DwgBlock In Dwgblks
                ReDim Preserve ArrayRezult(1, countArrayRezult)
                ArrayRezult(0, countArrayRezult) = dwgblk.Name
                ArrayRezult(1, countArrayRezult) = dwgblk.Description
                countArrayRezult += 1
            Next
        End If
        FuncReadInsertBlock = True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция проверяет наличие определения блока в чертеже
    Public Shared Function FuncReadBlockByName(ByVal ActivDocument As Drawing, ByVal nameBlock As String) As DwgBlock
        FuncReadBlockByName = Nothing
        Try
            If IsNothing(nameBlock) = False And IsNothing(ActivDocument) = False Then
                If nameBlock.Trim.Length > 0 Then
                    Dim Dwgblks As DwgBlocks = ActivDocument.Blocks
                    If Dwgblks.Count > 1 Then
                        For Each dwgblk As DwgBlock In Dwgblks
                            If nameBlock.Trim Like dwgblk.Name Then
                                Return dwgblk
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция возвращает все блоки на чертеже с заданным именем
    Public Shared Function FuncReadBlk(ByVal ActivDocument As Drawing, ByVal NameBlk As String, ByRef ArrayRezult As DwgInsert()) As Integer
        FuncReadBlk = 0
        Dim dwgEnts As DwgEntities = ActivDocument.ActiveSpace.Entities
        If dwgEnts.Count > 0 Then
            For Each dwgEnt As DwgEntity In dwgEnts
                If TypeOf dwgEnt Is DwgInsert Then
                    Dim blk As DwgInsert = dwgEnt
                    If blk.EntityName Like NameBlk Then
                        ReDim Preserve ArrayRezult(FuncReadBlk)
                        ArrayRezult(FuncReadBlk) = blk
                        FuncReadBlk += 1
                    End If
                End If
            Next
        End If
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция читиет атрибуты блока, возвращает массив (Tag,Prompt,Content)
    Public Shared Function FuncReadAttributeBlk(ByVal ActivDocument As Drawing, ByVal dwgBlk As DwgInsert, ByRef ArrayRezult As String(,)) As Boolean
        FuncReadAttributeBlk = False
        Dim CountArrayRezult As Integer = 0
        If IsArray(ArrayRezult) = True Then
            CountArrayRezult = ArrayRezult.GetUpperBound(1)
        End If
        Dim attribute As IEnumerable(Of IAttrib) = dwgBlk.Attribs
        If attribute.Count > 0 Then
            For i As Integer = 0 To attribute.Count - 1
                Dim attr As IAttrib = attribute.ElementAt(i)
                ReDim Preserve ArrayRezult(3, CountArrayRezult)
                ArrayRezult(0, CountArrayRezult) = attr.TagString
                ArrayRezult(1, CountArrayRezult) = attr.PromtString
                ArrayRezult(2, CountArrayRezult) = attr.Content
                ArrayRezult(3, CountArrayRezult) = attr.Style.Name
                CountArrayRezult += 1
            Next
        End If
        FuncReadAttributeBlk = True
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция читиет атрибуты блока, возвращает массив (Tag,Prompt,Content)
    Public Shared Function FuncReadNameAttribBlk(ByVal ActivDocument As Drawing, ByVal dwgBlk As DwgBlock, ByRef ArrayRezult As String()) As Boolean
        FuncReadNameAttribBlk = False
        Dim CountArrayRezult As Integer = 0
        If IsArray(ArrayRezult) = True Then
            CountArrayRezult = ArrayRezult.GetUpperBound(1)
        End If
        If IsNothing(dwgBlk) = False Then
            If dwgBlk.Count > 0 Then
                For i As Integer = 0 To dwgBlk.Count - 1
                    Dim dwgEnt As DwgEntity = dwgBlk.Item(i)
                    If TypeOf dwgEnt Is DwgAttdef Then
                        Dim attribDef As DwgAttdef = dwgEnt
                        ReDim Preserve ArrayRezult(CountArrayRezult)
                        ArrayRezult(CountArrayRezult) = attribDef.Content
                        CountArrayRezult += 1
                    End If
                Next i
            End If
        End If
        FuncReadNameAttribBlk = True
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция перезаписыват значения атрибутов блока, входной массив (Tag,Content)
    Public Shared Function FuncOverwriteAttributeBlk(ByVal ActivDocument As Drawing, ByVal dwgBlk As DwgInsert, ByRef ArrayRezult As String(,)) As Boolean
        FuncOverwriteAttributeBlk = False
        If IsArray(ArrayRezult) = False Then
            Return False
        End If
        Dim attribute As IEnumerable(Of IAttrib) = dwgBlk.Attribs
        If attribute.Count > 0 Then
            For i As Integer = 0 To attribute.Count - 1
                Dim attr As IAttrib = attribute.ElementAt(i)
                Dim tagAttrib As String = attr.TagString
                For j As Integer = 0 To ArrayRezult.GetUpperBound(1)
                    Dim tagUser As String = ArrayRezult(0, j)
                    If IsNothing(tagUser) = False Then
                        If tagUser.Trim Like tagAttrib Then
                            attr.Content = ArrayRezult(1, j)
                            Exit For
                        End If
                    End If
                Next
            Next
        End If
        FuncOverwriteAttributeBlk = True
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вставка определения блока из внешнего файла
    Public Shared Function FuncInsertBlockByDWGFile(ByVal ActivDocument As Drawing, ByVal putchTemplateDwg As String, ByVal nameBlock As String) As DwgBlock
        FuncInsertBlockByDWGFile = Nothing
        If File.Exists(putchTemplateDwg) = False Then Exit Function
        If IsNothing(nameBlock) = True Then Exit Function
        If nameBlock.Trim.Length = 0 Then Exit Function
        If IsNothing(ActivDocument) = True Then Exit Function
        Try
            If IsNothing(ActivDocument.Blocks.Item(nameBlock)) = True Then
                Dim sourceDraw As Drawing = New Drawing()
                Topomatic.Acax.Import.Dxf.AcaxImporter.Import(sourceDraw, putchTemplateDwg)
                Dim acBlk As DwgBlock = sourceDraw.Blocks.Item(nameBlock)
                If IsNothing(acBlk) = False Then
                    ActivDocument.BeginUpdate()
                    Try
                        Dim newBlk As DwgBlock = ActivDocument.Blocks.Add(acBlk.Name)
                        newBlk.Entities.CopyFrom(acBlk, New ReferencesContext(sourceDraw))
                        For i As Integer = 0 To newBlk.Count - 1
                            Dim dwgEnt As DwgEntity = newBlk.ElementAt(i)
                            If TypeOf dwgEnt Is DwgAttdef Then
                                Dim userAttrDef As DwgAttdef = dwgEnt
                                If userAttrDef.Content = "####" Then
                                    'ActivDocument.ActiveSpace.Entities.Remove(userAttrDef)
                                End If
                            End If
                        Next i
                        ActivDocument.EndUpdate()
                        Return newBlk
                    Catch ex As Exception
                    End Try
                End If
            Else
                Return ActivDocument.Blocks.Item(nameBlock)
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция вставляет блок в чертеж (height - высота, positionX, PositionY,flag,visible
    Public Shared Function FuncInsertBlk(ByVal ActivDocument As Drawing, ByVal nameBlk As String, ByVal position As Vector3D, ByVal attributeProperties As Dictionary(Of String, String(,)), Optional rotationBlk As Double = 0, Optional scaleBlk As Double = 1, Optional boolBackground As Boolean = False, Optional boolMAttrib As Boolean = False) As DwgInsert
        FuncInsertBlk = Nothing
        If IsNothing(nameBlk) = True Then Exit Function
        Dim blk As DwgBlock = ActivDocument.Blocks.Item(nameBlk)
        If IsNothing(blk) = False Then
            Try
                ActivDocument.BeginUpdate()
                Dim insertBlk As DwgInsert = New DwgInsert()
                insertBlk.Block = blk
                insertBlk.Position = position
                insertBlk.Scale = Vector3D.One
                insertBlk.Rotation = rotationBlk
                insertBlk.Scale = New Topomatic.Cad.Foundation.Vector3D(scaleBlk, scaleBlk, scaleBlk)
                '=======================================================================================================
                If insertBlk.HasAttribs = True Then
                    Dim collectionAttributes As IEnumerable(Of IAttrib) = insertBlk.Attribs
                    'массив для считывания атрибутов блока
                    If collectionAttributes.Count > 0 And attributeProperties.Count > 0 Then
                        Dim arrayProperties As String(,) = Nothing
                        ReDim Preserve arrayProperties(20, collectionAttributes.Count - 1)
                        For i As Integer = 0 To collectionAttributes.Count - 1
                            Dim userIAttrib As IAttrib = collectionAttributes.ElementAt(i)
                            Dim userAttrib As DwgAttrib = userIAttrib
                            arrayProperties(0, i) = userAttrib.TagString
                            arrayProperties(1, i) = userAttrib.PromtString
                            arrayProperties(2, i) = userAttrib.Content
                            arrayProperties(3, i) = userAttrib.Position.X
                            arrayProperties(4, i) = userAttrib.Position.Y
                            arrayProperties(5, i) = userAttrib.Style.Name
                            arrayProperties(6, i) = userAttrib.Height
                            arrayProperties(7, i) = userAttrib.Rotation
                            arrayProperties(8, i) = userAttrib.Flags
                            arrayProperties(9, i) = userAttrib.Annotative
                            arrayProperties(10, i) = userAttrib.Oblique
                            arrayProperties(11, i) = userAttrib.Bounds.Min.X
                            arrayProperties(12, i) = userAttrib.Bounds.Min.Y
                            arrayProperties(13, i) = userAttrib.Bounds.Max.X
                            arrayProperties(14, i) = userAttrib.Bounds.Max.Y
                            arrayProperties(15, i) = userAttrib.Color.ColorIndex
                            arrayProperties(16, i) = userAttrib.Fixed
                            arrayProperties(17, i) = userAttrib.Invisible
                            arrayProperties(18, i) = userAttrib.Justify
                            arrayProperties(19, i) = userAttrib.Layer.Name
                            arrayProperties(19, i) = userAttrib.TextAlignmentPoint.X
                            arrayProperties(20, i) = userAttrib.TextAlignmentPoint.Y
                            insertBlk.RemoveAttrib(userIAttrib)
                        Next i
                        For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                            Try
                                Dim nameAttr As String = arrayProperties(0, i)
                                Dim promptAttr As String = arrayProperties(1, i)
                                If nameAttr.Contains("####") Then
                                    nameAttr = "ISNAME"
                                End If
                                'вставляем новый атрибут
                                If boolMAttrib = False Then
                                    '==================================================================================
                                    'однострочный атрибут
                                    Dim userAttrib As DwgAttrib = insertBlk.AddAttrib(nameAttr, boolMAttrib)
                                    userAttrib.PromtString = arrayProperties(1, i)
                                    userAttrib.Content = ""
                                    Dim posAttr As Topomatic.Cad.Foundation.Vector2D = New Topomatic.Cad.Foundation.Vector2D(Val(arrayProperties(3, i)), Val(arrayProperties(4, i)))
                                    userAttrib.Position = posAttr
                                    Dim styleAttr As DwgStyle = ActivDocument.Styles.Item(arrayProperties(5, i))
                                    userAttrib.Style = styleAttr
                                    userAttrib.Height = Val(arrayProperties(6, i))
                                    userAttrib.Rotation = Val(arrayProperties(7, i))
                                    userAttrib.Flags = Val(arrayProperties(8, i))
                                    userAttrib.Annotative = CBool(arrayProperties(9, i))
                                    userAttrib.Oblique = arrayProperties(10, i)
                                    Dim userBound As BoundingBox2D = New BoundingBox2D
                                    userBound.Min.X = arrayProperties(11, i)
                                    userBound.Min.Y = arrayProperties(12, i)
                                    userBound.Max.X = arrayProperties(13, i)
                                    userBound.Max.Y = arrayProperties(14, i)
                                    Dim userColor As CadColor = New CadColor(Val(arrayProperties(15, i)))
                                    userAttrib.Color = userColor
                                    userAttrib.Fixed = False
                                    userAttrib.Flags = AttributeFlags.VerificationRequired
                                    userAttrib.Invisible = CBool(arrayProperties(17, i))
                                    userAttrib.Justify = Val(arrayProperties(18, i))
                                    Dim userLayer As DwgLayer = ActivDocument.Layers.Item(arrayProperties(19, i))
                                    userAttrib.Layer = userLayer
                                    Dim posPoint As Vector3D = New Vector3D(Val(arrayProperties(19, i)), Val(arrayProperties(20, i)), insertBlk.Position.Z)
                                    userAttrib.TextAlignmentPoint = posPoint
                                    userAttrib.Backward = False
                                    If IsNothing(attributeProperties) = False Then
                                        If attributeProperties.Count > 0 Then
                                            For j As Integer = 0 To attributeProperties.Count - 1
                                                Dim attrDict As KeyValuePair(Of String, String(,)) = attributeProperties.ElementAt(j)
                                                Dim KeyAttr As String = attrDict.Key
                                                If KeyAttr Like nameAttr Then
                                                    Dim userArrayProperties As String(,) = attrDict.Value
                                                    If IsArray(userArrayProperties) = True Then
                                                        Dim posAttribX As Double = 0
                                                        Dim posAttribY As Double = 0
                                                        Dim boolValue As Boolean = False
                                                        For k As Integer = 0 To userArrayProperties.GetUpperBound(1)
                                                            Dim nameF As String = userArrayProperties(0, k)
                                                            If IsNothing(nameF) = False Then
                                                                If nameF Like "height" Then
                                                                    userAttrib.Height = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "positionX" Then
                                                                    posAttribX = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "positionY" Then
                                                                    posAttribY = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "flag" Then
                                                                    userAttrib.Flags = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "visible" Then
                                                                    userAttrib.Invisible = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "value" Then
                                                                    userAttrib.Content = userArrayProperties(1, k)
                                                                    boolValue = True
                                                                ElseIf nameF Like "Justify" Then
                                                                    userAttrib.Justify = Val(userArrayProperties(1, k))
                                                                End If
                                                                If boolValue = False Then
                                                                    If nameAttr Like "CENTER_LEVEL" Or nameAttr Like "ZERO_LEVEL" Or nameAttr Like "BORDER_UP_LEVEL" Then
                                                                        userAttrib.Content = Math.Round(position.Z, 2)
                                                                    End If
                                                                End If
                                                                If posAttribX <> 0 OrElse posAttribY <> 0 Then
                                                                    Dim posAttr1 As Vector3D = New Vector3D(posAttribX, posAttribY, insertBlk.Position.Z)
                                                                    userAttrib.Position = posAttr1
                                                                End If
                                                            End If
                                                        Next k
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                Else
                                    '=======================================================================================
                                    'многострочный атрибут
                                    Dim userAttrib As DwgMAttrib = insertBlk.AddAttrib(nameAttr, boolMAttrib)
                                    userAttrib.PromtString = arrayProperties(1, i)
                                    userAttrib.Content = ""
                                    Dim posAttr As Vector3D = New Vector3D(Val(arrayProperties(3, i)), Val(arrayProperties(4, i)), insertBlk.Position.Z)
                                    userAttrib.Position = posAttr
                                    Dim styleAttr As DwgStyle = ActivDocument.Styles.Item(arrayProperties(5, i))
                                    userAttrib.Style = styleAttr
                                    userAttrib.Height = Val(arrayProperties(6, i))
                                    userAttrib.Rotation = Val(arrayProperties(7, i))
                                    userAttrib.Flags = Val(arrayProperties(8, i))
                                    userAttrib.Annotative = CBool(arrayProperties(9, i))
                                    Dim userBound As BoundingBox2D = New BoundingBox2D
                                    userBound.Min.X = arrayProperties(11, i)
                                    userBound.Min.Y = arrayProperties(12, i)
                                    userBound.Max.X = arrayProperties(13, i)
                                    userBound.Max.Y = arrayProperties(14, i)
                                    Dim userColor As CadColor = New CadColor(Val(arrayProperties(15, i)))
                                    userAttrib.Color = userColor
                                    userAttrib.Fixed = False
                                    userAttrib.Flags = AttributeFlags.VerificationRequired
                                    userAttrib.Invisible = CBool(arrayProperties(17, i))
                                    userAttrib.AttachmentPoint = AttachmentPoint.BottomLeft
                                    userAttrib.BackgroundColor = New CadColor(255)
                                    userAttrib.BackgroundFillType = BackgroundFillType.UseFillColor
                                    userAttrib.RectangleHeight = 0
                                    userAttrib.RectangleHeight = 0
                                    userAttrib.FillBoxScale = 1.1
                                    Dim userLayer As DwgLayer = ActivDocument.Layers.Item(arrayProperties(19, i))
                                    userAttrib.Layer = userLayer
                                    Dim posPoint As Vector3D = New Vector3D(Val(arrayProperties(19, i)), Val(arrayProperties(20, i)), 0)
                                    If IsNothing(attributeProperties) = False Then
                                        If attributeProperties.Count > 0 Then
                                            For j As Integer = 0 To attributeProperties.Count - 1
                                                Dim attrDict As KeyValuePair(Of String, String(,)) = attributeProperties.ElementAt(j)
                                                Dim KeyAttr As String = attrDict.Key
                                                If KeyAttr Like nameAttr Then
                                                    Dim userArrayProperties As String(,) = attrDict.Value
                                                    If IsArray(userArrayProperties) = True Then
                                                        Dim posAttribX As Double = 0
                                                        Dim posAttribY As Double = 0
                                                        Dim boolValue As Boolean = False
                                                        For k As Integer = 0 To userArrayProperties.GetUpperBound(1)
                                                            Dim nameF As String = userArrayProperties(0, k)
                                                            If IsNothing(nameF) = False Then
                                                                If nameF Like "height" Then
                                                                    userAttrib.Height = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "positionX" Then
                                                                    posAttribX = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "positionY" Then
                                                                    posAttribY = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "flag" Then
                                                                    userAttrib.Flags = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "visible" Then
                                                                    userAttrib.Invisible = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "value" Then
                                                                    userAttrib.Content = userArrayProperties(1, k)
                                                                    boolValue = True
                                                                ElseIf nameF Like "AttachmentPoint" Then
                                                                    userAttrib.AttachmentPoint = Val(userArrayProperties(1, k))
                                                                ElseIf nameF Like "Background" Then
                                                                    userAttrib.BackgroundFillType = BackgroundFillType.UseFillColor
                                                                    userAttrib.BackgroundColor = New CadColor(255)
                                                                End If
                                                                If boolValue = False Then
                                                                    If nameAttr Like "CENTER_LEVEL" Or nameAttr Like "ZERO_LEVEL" Or nameAttr Like "BORDER_UP_LEVEL" Then
                                                                        userAttrib.Content = Math.Round(position.Z, 2)
                                                                    End If
                                                                End If
                                                                If posAttribX <> 0 OrElse posAttribY <> 0 Then
                                                                    Dim posAttr1 As Vector3D = New Vector3D(posAttribX, posAttribY, insertBlk.Position.Z)
                                                                    userAttrib.Position = posAttr1
                                                                End If
                                                            End If
                                                        Next k
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                End If
                            Catch ex As System.Exception
                            End Try
                        Next i
                    End If
                End If
                ActivDocument.ActiveSpace.Add(insertBlk)
                ActivDocument.EndUpdate()
                Return insertBlk
            Catch ex1 As System.Exception
            End Try
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вставка чертежа из внешнего файла как блок
    Public Shared Function FuncInsertOutDWG(ByVal ActivDocument As Drawing, ByVal putchOutDwg As String) As DwgInsert
        FuncInsertOutDWG = Nothing
        If File.Exists(putchOutDwg) = False Then Exit Function
        Try
            Dim sourceDraw As Drawing = New Drawing()
            Topomatic.Acax.Import.Dxf.AcaxImporter.Import(sourceDraw, putchOutDwg)
            ActivDocument.BeginUpdate()
            Dim blk As DwgBlock = ActivDocument.Blocks.Add(ActivDocument.Blocks.GenerateUnicalName())
            blk.Entities.CopyFrom(sourceDraw.ActiveSpace, New ReferencesContext(sourceDraw))
            Dim insertBlk As DwgInsert = New DwgInsert
            insertBlk.Block = blk
            insertBlk.Position = Vector3D.Empty
            insertBlk.Scale = Vector3D.One
            ActivDocument.ActiveSpace.Add(insertBlk)
            ActivDocument.EndUpdate()
            Return insertBlk
        Catch ex As System.Exception
        End Try
    End Function



    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ТИПЫ ЛИНИЙ
    '================================================================================================
    'создание линии
    Public Shared Function FuncDrawLine(ByVal ActivDocument As Drawing, ByVal StartPoint As Vector3D, ByVal EndPoint As Vector3D) As DwgLine
        FuncDrawLine = Nothing
        Try
            FuncDrawLine = New Topomatic.Dwg.Entities.DwgLine
            FuncDrawLine.StartPoint = StartPoint
            FuncDrawLine.EndPoint = EndPoint
            ActivDocument.ActiveSpace.Add(FuncDrawLine)
        Catch ex As Exception
        End Try
    End Function


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нумеровать полигон
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncNomberPolygon(ByVal ActivDocument As Drawing, ByVal acPoly As DwgPolyline, ByVal blockInsert As DwgBlock, ByVal layerBlk As DwgLayer, ByVal colorBlk As CadColor, Optional startNumber As Integer = 0, Optional prefNumber As String = "", Optional arrayPropertiesBlk As String(,) = Nothing) As Boolean
        FuncNomberPolygon = False
        If IsArray(acPoly) = False Then Return False
        If IsArray(blockInsert) = False Then Return False
        If IsArray(layerBlk) = False Then layerBlk = ActivDocument.ActiveLayer
        'атрибуты 0-имя атрибута
        For Each vertPoly As BugleVector2D In acPoly

        Next


    End Function



    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '3D OBJECT
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция изменяет координаты вставки 3д объекта
    Public Shared Sub LineDynamicRender(ByRef userCadView As CadView, ByRef arrayvertex As Double(,))
        If userCadView IsNot Nothing Then
            Dim positions As New List(Of Topomatic.Cad.Foundation.Vector2D)()
            Dim dynamic_draw As DrawCursorEvent = Sub(pen As CadPen, vertex As Topomatic.Cad.Foundation.Vector2D)
                                                      If positions.Count > 0 Then
                                                          pen.Color = System.Drawing.Color.Lime
                                                          pen.BeginDraw()
                                                          Try
                                                              For i As Integer = 1 To positions.Count - 1
                                                                  'Используем рефлексию для вызова метода
                                                                  Dim method = pen.GetType().GetMethod("DrawLine",
                                New Type() {GetType(Topomatic.Cad.Foundation.Vector2D), GetType(Topomatic.Cad.Foundation.Vector2D)})
                                                                  method.Invoke(pen, New Object() {
                                New Topomatic.Cad.Foundation.Vector2D(positions(i - 1)),
                                New Topomatic.Cad.Foundation.Vector2D(positions(i))
                            })
                                                              Next
                                                              Dim method2 = pen.GetType().GetMethod("DrawLine",
                            New Type() {GetType(Topomatic.Cad.Foundation.Vector2D), GetType(Topomatic.Cad.Foundation.Vector2D)})
                                                              method2.Invoke(pen, New Object() {
                            New Topomatic.Cad.Foundation.Vector2D(positions(positions.Count - 1)),
                            vertex
                        })
                                                          Finally
                                                              pen.EndDraw()
                                                          End Try
                                                      End If
                                                  End Sub
            AddHandler userCadView.DynamicDraw, dynamic_draw
            Try
                Dim pos As Vector3D
                While CadCursors.GetPoint(userCadView, pos, "Укажите точку")
                    positions.Add(pos.Pos)
                    'vertexResult.Add(pos.Pos)
                End While
            Finally
                RemoveHandler userCadView.DynamicDraw, dynamic_draw
            End Try
        End If
    End Sub
End Class
