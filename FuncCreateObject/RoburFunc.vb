Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Shapes
Imports NetTopologySuite.Algorithm
Imports NetTopologySuite.Mathematics
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Plugins
'Imports Topomatic.Alg.CogoController
'Imports Topomatic.Sfc
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Controls.Dialogs
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
'Imports Topomatic.Srv
Imports Topomatic.Dwg.Layer
Imports Topomatic.Planchet.Entities
Imports Topomatic.Sfc
Imports Topomatic.Sfc.Layer
Imports Topomatic.Visualization.Geometry
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D

Public Class RoburFunc
    'РИСОВАНИЕ ПРИМИТИВОВ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать линию
    Public Shared Function FuncDrawLineToPoint(ByVal ActivDocument As Drawing, ByVal point1 As Vector3D, ByVal point2 As Vector3D) As DwgLine
        FuncDrawLineToPoint = Nothing
        If IsNothing(point1) = True Then Exit Function
        If IsNothing(point2) = True Then Exit Function
        If point1 <> point2 Then
            Dim line As DwgLine = New Topomatic.Dwg.Entities.DwgLine()
            line.StartPoint = point1
            line.EndPoint = point2
            ActivDocument.ActiveSpace.Add(line)
            Return line
        End If
    End Function

    'создание полилинии из массива
    Public Shared Function FuncDrawPolylineToArrayCoord(ByVal ActivDocument As Drawing, ByRef ArrayVertex2d As Double(,), Optional ByVal boolClosed As Boolean = False, Optional boolReverse As Boolean = False) As DwgPolyline
        FuncDrawPolylineToArrayCoord = Nothing
        If IsArray(ArrayVertex2d) = True Then
            If ArrayVertex2d.GetUpperBound(1) > 0 Then
                Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
                Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
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
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
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
                        polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, bulgeSegment))
                    Next i
                End If
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

    'создание полилинии из списка
    Public Shared Function FuncDrawPolylineToListCoord(ByVal ActivDocument As Drawing, ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As DwgPolyline
        FuncDrawPolylineToListCoord = Nothing
        If points.Count > 1 Then
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            Dim polyline As DwgPolyline = New Topomatic.Dwg.Entities.DwgPolyline()
            For i As Integer = 0 To points.Count - 1
                rp.X = points.Item(i).X
                rp.Y = points.Item(i).Y
                rp.Z = 0
                Dim pos = New Topomatic.Cad.Foundation.Vector2D(rp.X, rp.Y)
                polyline.Add(New Topomatic.Cad.Foundation.BugleVector2D(pos, 0))
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

    '=================================================================================================
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

    '=================================================================================================
    'создание 3d полилинии из списка
    Public Shared Function FuncDrawPolyline3DToListCoord(ByVal ActivDocument As Drawing, ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As DwgPolyline3D
        FuncDrawPolyline3DToListCoord = Nothing
        If points.Count > 1 Then
            Dim polyline As DwgPolyline3D = New Topomatic.Dwg.Entities.DwgPolyline3D()
            For Each pos As Vector3D In points
                polyline.Add(pos)
            Next
            If IsNothing(ActivDocument) = False Then
                ActivDocument.ActiveSpace.Add(polyline)
                If boolClosed = True Then
                    polyline.Closed = True
                End If
                Return polyline
            End If
        End If
    End Function

    '=================================================================================================
    'обновление 3d полилинии из списка
    Public Shared Function FuncReDrawPolyline3DToListCoord(ByRef polyline3d As DwgPolyline3D, ByRef points As List(Of Vector3D), Optional ByVal boolClosed As Boolean = False) As Boolean
        FuncReDrawPolyline3DToListCoord = False
        If IsNothing(polyline3d) = True Then Return False
        If IsNothing(points) = True Then Return False
        If points.Count < 2 Then
            Return False
        End If
        If points.Count > 1 Then
            polyline3d.Clear()
            For Each pos As Vector3D In points
                polyline3d.Add(pos)
            Next
            If polyline3d.Count > 2 And boolClosed = True Then
                polyline3d.Closed = True
            End If
            Return True
        End If
    End Function

    '================================================================================================
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
    '================================================================================================
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
    '================================================================================================
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

    '================================================================================================
    'создание параллельного размера
    Public Shared Function FuncDrawDimPar(ByVal ActivDocument As Drawing, ByVal Str As String, ByVal InsertFirstVertex As Topomatic.Cad.Foundation.Vector2D, ByVal InserSecondVertex As Topomatic.Cad.Foundation.Vector2D, Optional nameArrow As String = "", Optional hText As Double = 2) As DwgDimension
        FuncDrawDimPar = Nothing
        Try

        Catch ex As Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'учасикм м штриховки
    '================================================================================================
    'создание штриховки
    Public Shared Function FuncDrawHatch(ByVal ActivDocument As Drawing, ByVal acPline As Topomatic.Dwg.Entities.DwgPolyline, ByVal nameHatch As String) As DwgHatch
        FuncDrawHatch = Nothing
        Try
            FuncDrawHatch = New Topomatic.Dwg.Entities.DwgHatch
            Dim points As List(Of Topomatic.Cad.Foundation.Vector2D) = New List(Of Topomatic.Cad.Foundation.Vector2D)
            acPline.ConvertToPosArray(points)
            FuncDrawHatch = ActivDocument.ActiveSpace.AddHatch(AcPatternType.PreDefined, nameHatch, points)

        Catch ex As Exception
        End Try
    End Function
    '================================================================================================
    'создание штриховки
    Public Shared Function FuncDrawHatch2(ByVal ActivDocument As Drawing, ByVal acPline As Topomatic.Dwg.Entities.DwgPolyline, ByVal nameHatch As String) As DwgHatch
        FuncDrawHatch2 = Nothing
        Try
            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
            For Each vertex As BugleVector2D In acPline
                BoundaryPath.Add(vertex)
            Next
            FuncDrawHatch2 = New Topomatic.Dwg.Entities.DwgHatch()
            FuncDrawHatch2.BoundaryPath.Add(BoundaryPath)
            FuncDrawHatch2.PatternName = nameHatch
            ActivDocument.ActiveSpace.Add(FuncDrawHatch2)
        Catch ex As Exception
        End Try
    End Function
    Public Shared Function FuncDrawHatchByPolyline3d(ByVal ActivDocument As Drawing, ByVal acPline3d As Topomatic.Dwg.Entities.DwgPolyline3D, ByVal nameHatch As String) As DwgHatch
        FuncDrawHatchByPolyline3d = Nothing
        If IsNothing(acPline3d) = False Then
            If acPline3d.Count > 2 Then
                If acPline3d.Area > 0 Then

                    Try
                        Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
                        For i As Integer = 0 To acPline3d.Count - 1
                            Dim vertex As Topomatic.Cad.Foundation.Vector2D = acPline3d.Item(i).Pos
                            BoundaryPath.Add(New BugleVector2D(vertex, 0))
                        Next
                        If BoundaryPath.Count > 2 Then
                            Dim newHatch As DwgHatch = New DwgHatch
                            newHatch.BoundaryPath.Add(BoundaryPath)
                            newHatch.PatternName = nameHatch
                            ActivDocument.ActiveSpace.Add(newHatch)
                            Return newHatch
                        End If
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If
    End Function
    '================================================================================================
    'обновление штриховки
    Public Shared Function FuncReDrawHatchByPolyline(ByVal ActivDocument As Drawing, ByVal acPline As Topomatic.Dwg.Entities.DwgPolyline, ByRef acHatch As DwgHatch) As Boolean
        Try
            If IsNothing(acPline) = True Then Return False
            If acPline.Count < 3 Then Return False
            If IsNothing(acHatch) = True Then Return False
            acHatch.BoundaryPath.Clear()
            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
            For Each vertex As BugleVector2D In acPline
                BoundaryPath.Add(vertex)
            Next
            acHatch.BoundaryPath.Add(BoundaryPath)
        Catch ex As Exception
        End Try
    End Function
    Public Shared Function FuncReDrawHatchByPolyline3d(ByVal ActivDocument As Drawing, ByVal ac3dPline As Topomatic.Dwg.Entities.DwgPolyline3D, ByRef acHatch As DwgHatch) As Boolean
        Try
            If IsNothing(ac3dPline) = True Then Return False
            If ac3dPline.Count < 3 Then Return False
            If IsNothing(acHatch) = True Then Return False
            acHatch.BoundaryPath.Clear()
            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
            For i As Integer = 0 To ac3dPline.Count - 1
                Dim vertex As Topomatic.Cad.Foundation.Vector2D = ac3dPline.Item(i).Pos
                BoundaryPath.Add(New BugleVector2D(vertex, 0))
            Next
            acHatch.BoundaryPath.Add(BoundaryPath)
        Catch ex As Exception
        End Try
    End Function
    '================================================================================================
    'преобразование Участка Робур в штриховку
    '==================================================================================================================================================================
    'функция возвращает координаты участка Робур
    Public Shared Function FuncDrawHatchByParcels(ByVal ActivDocument As Drawing, ByVal acParcels As Topomatic.Sfc.Layer.Wrappers.SurfacePatchWrapper, ByVal namePattern As String, ByVal layerObj As DwgLayer, ByVal colorObj As CadColor, ByVal ScaleLineObj As Double, ByVal WidthLineObj As Double, ByVal ScaleObj As Double, ByVal angleObj As Double, Optional ByVal arrayRec As String(,) = Nothing, Optional ByVal nameTablePs As String = "IS_SYSTEM") As Boolean
        FuncDrawHatchByParcels = False
        Dim userSurf As Surface = acParcels.Surface
        Dim listTrg As List(Of Integer) = acParcels.List
        Dim ArrayTrianglRezult As Double(,)
        Dim countArrayTrianglRezult As Integer = 0
        '====================================================================================
        'забираем все грани треугольников
        If listTrg.Count > 0 Then
            For Each i As Integer In listTrg
                Dim triangl As SurfaceTriangle = acParcels.Surface.Triangles.Item(i)
                Dim a As Integer = triangl.A
                Dim b As Integer = triangl.B
                Dim c As Integer = triangl.C
                Dim surfPointA As SurfacePoint = userSurf.Points.Item(a)
                Dim surfPointB As SurfacePoint = userSurf.Points.Item(b)
                Dim surfPointC As SurfacePoint = userSurf.Points.Item(c)
                ReDim Preserve ArrayTrianglRezult(9, countArrayTrianglRezult)
                ArrayTrianglRezult(0, countArrayTrianglRezult) = Math.Round(surfPointA.Vertex.X, 4)
                ArrayTrianglRezult(1, countArrayTrianglRezult) = Math.Round(surfPointA.Vertex.Y, 4)
                ArrayTrianglRezult(2, countArrayTrianglRezult) = Math.Round(surfPointA.Vertex.Z, 4)

                ArrayTrianglRezult(3, countArrayTrianglRezult) = Math.Round(surfPointB.Vertex.X, 4)
                ArrayTrianglRezult(4, countArrayTrianglRezult) = Math.Round(surfPointB.Vertex.Y, 4)
                ArrayTrianglRezult(5, countArrayTrianglRezult) = Math.Round(surfPointB.Vertex.Z, 4)

                ArrayTrianglRezult(6, countArrayTrianglRezult) = Math.Round(surfPointC.Vertex.X, 4)
                ArrayTrianglRezult(7, countArrayTrianglRezult) = Math.Round(surfPointC.Vertex.Y, 4)
                ArrayTrianglRezult(8, countArrayTrianglRezult) = Math.Round(surfPointC.Vertex.Z, 4)

                ArrayTrianglRezult(9, countArrayTrianglRezult) = 0
                countArrayTrianglRezult += 1
            Next

            If IsArray(ArrayTrianglRezult) = True Then
                Dim points As List(Of Vector3D) = New List(Of Vector3D)
                Do
                    Dim flagprogs As Boolean = False
                    For i As Integer = 0 To ArrayTrianglRezult.GetUpperBound(1)
                        flagprogs = False
                        If ArrayTrianglRezult(9, i) = 0 Then
                            Dim ax As Double = ArrayTrianglRezult(0, i)
                            Dim ay As Double = ArrayTrianglRezult(1, i)
                            Dim az As Double = ArrayTrianglRezult(2, i)

                            Dim bx As Double = ArrayTrianglRezult(3, i)
                            Dim by As Double = ArrayTrianglRezult(4, i)
                            Dim bz As Double = ArrayTrianglRezult(5, i)

                            Dim cx As Double = ArrayTrianglRezult(6, i)
                            Dim cy As Double = ArrayTrianglRezult(7, i)
                            Dim cz As Double = ArrayTrianglRezult(8, i)
                            If points.Count = 0 Then
                                Dim pos1 = New Topomatic.Cad.Foundation.Vector3D(ax, ay, az)
                                points.Add(pos1)

                                Dim pos2 = New Topomatic.Cad.Foundation.Vector3D(bx, by, bz)
                                points.Add(pos2)

                                Dim pos3 = New Topomatic.Cad.Foundation.Vector3D(cx, cy, cz)
                                points.Add(pos3)
                                flagprogs = True
                                ArrayTrianglRezult(9, i) = 1
                                Exit For
                            Else
                                For j As Integer = 0 To points.Count - 1
                                    Dim pt1 As Vector3D = New Vector3D(0, 0, 0)
                                    Dim pt2 As Vector3D = New Vector3D(0, 0, 0)
                                    Dim pt3 As Vector3D = New Vector3D(0, 0, 0)
                                    If j = points.Count - 2 Then
                                        pt1 = points.Item(j)
                                        pt2 = points.Item(j + 1)
                                        pt3 = points.Item(0)
                                    ElseIf j = points.Count - 1 Then
                                        pt1 = points.Item(j)
                                        pt2 = points.Item(0)
                                        pt3 = points.Item(1)
                                    Else
                                        pt1 = points.Item(j)
                                        pt2 = points.Item(j + 1)
                                        pt3 = points.Item(j + 2)
                                    End If
                                    'все вершины совпадают
                                    If pt1.X = ax And pt1.Y = ay And pt2.X = bx And pt2.Y = by And pt3.X = cx And pt3.Y = cy Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = bx And pt1.Y = by And pt2.X = cx And pt2.Y = cy And pt3.X = ax And pt3.Y = ay Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = cx And pt1.Y = cy And pt2.X = ax And pt2.Y = ay And pt3.X = bx And pt3.Y = by Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt1.X = ax And pt1.Y = ay And pt2.X = cx And pt2.Y = cy And pt3.X = bx And pt3.Y = by Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = bx And pt1.Y = by And pt2.X = ax And pt2.Y = ay And pt3.X = cx And pt3.Y = cy Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = cx And pt1.Y = cy And pt2.X = bx And pt2.Y = by And pt3.X = ax And pt1.Y = ay Then
                                        points.Remove(pt2)
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                        'совпадают только 2 вершины, добавляем вершину в массив

                                    ElseIf pt1.X = ax And pt1.Y = ay And pt2.X = bx And pt2.Y = by Then
                                        points.Insert(j + 1, New Vector3D(cx, cy, cz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = ax And pt1.Y = ay And pt2.X = cx And pt2.Y = cy Then
                                        points.Insert(j + 1, New Vector3D(bx, by, bz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt1.X = bx And pt1.Y = by And pt2.X = cx And pt2.Y = cy Then
                                        points.Insert(j + 1, New Vector3D(ax, ay, az))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = bx And pt1.Y = by And pt2.X = ax And pt2.Y = ay Then
                                        points.Insert(j + 1, New Vector3D(cx, cy, cz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt1.X = cx And pt1.Y = cy And pt2.X = ax And pt2.Y = ay Then
                                        points.Insert(j + 1, New Vector3D(bx, by, bz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt1.X = cx And pt1.Y = cy And pt2.X = bx And pt2.Y = by Then
                                        points.Insert(j + 1, New Vector3D(ax, ay, az))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt2.X = ax And pt2.Y = ay And pt1.X = bx And pt1.Y = by Then
                                        points.Insert(j + 1, New Vector3D(cx, cy, cz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt2.X = ax And pt2.Y = ay And pt1.X = cx And pt1.Y = cy Then
                                        points.Insert(j + 1, New Vector3D(bx, by, bz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt2.X = bx And pt2.Y = by And pt1.X = cx And pt1.Y = cy Then
                                        points.Insert(j + 1, New Vector3D(ax, ay, az))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt2.X = bx And pt2.Y = by And pt1.X = ax And pt1.Y = ay Then
                                        points.Insert(j + 1, New Vector3D(cx, cy, cz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For

                                    ElseIf pt2.X = cx And pt2.Y = cy And pt1.X = ax And pt1.Y = ay Then
                                        points.Insert(j + 1, New Vector3D(bx, by, bz))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    ElseIf pt2.X = cx And pt2.Y = cy And pt1.X = bx And pt1.Y = by Then
                                        points.Insert(j + 1, New Vector3D(ax, ay, az))
                                        ArrayTrianglRezult(9, i) = 1
                                        flagprogs = True
                                        Exit For
                                    End If
                                Next j
                                If flagprogs = True Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next i
                    If flagprogs = False Then
                        If points.Count > 2 Then
                            Dim acHatch As DwgHatch = New Topomatic.Dwg.Entities.DwgHatch()
                            Dim BoundaryPath As PolylineBoundaryPath = New PolylineBoundaryPath()
                            For Each vertex As Vector3D In points
                                Dim tempV As Topomatic.Cad.Foundation.Vector2D = New Topomatic.Cad.Foundation.Vector2D(vertex.X, vertex.Y)
                                Dim vb As BugleVector2D = New BugleVector2D(tempV, 0)
                                BoundaryPath.Add(vb)
                            Next
                            acHatch.BoundaryPath.Add(BoundaryPath)
                            If IsNothing(layerObj) = False Then
                                acHatch.Layer = layerObj
                            Else
                                acHatch.Layer = ActivDocument.ActiveLayer
                            End If
                            Try
                                acHatch.PatternName = namePattern
                            Catch ex As System.ArgumentNullException
                                acHatch.PatternName = "SOLID"
                            End Try
                            If IsNothing(layerObj) = False Then
                                acHatch.Color = colorObj
                            End If
                            acHatch.Lineweight = WidthLineObj
                            acHatch.PatternAngle = angleObj
                            acHatch.PatternScale = ScaleObj
                            acHatch.LinetypeScale = ScaleLineObj

                            If IsArray(arrayRec) = True Then
                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(acHatch, nameTablePs, arrayRec)
                            End If

                            ActivDocument.ActiveSpace.Add(acHatch)
                            points = New List(Of Vector3D)
                        Else
                            Exit Do
                        End If
                    End If
                Loop
            End If
        End If
    End Function
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


    'чтение объектов автокад
    '////////////////////////////////////////////////////////////////////////////////////////////////////

    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция возвращает имена всех слоев на чертеже
    Public Shared Function FuncReadLayers(ByVal ActivDocument As Drawing, ByRef ArrayRezult As String()) As Boolean
        FuncReadLayers = False
        Dim countArrayRezult As Integer = 0
        If IsArray(ArrayRezult) = True Then
            countArrayRezult = ArrayRezult.Length
        End If
        Dim DwgLayers As DwgLayers = ActivDocument.Layers
        If DwgLayers.Count > 0 Then
            For Each dwglayer As DwgLayer In DwgLayers
                ReDim Preserve ArrayRezult(countArrayRezult)
                ArrayRezult(countArrayRezult) = dwglayer.Name
                countArrayRezult += 1
            Next
        End If
        FuncReadLayers = True
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция проверяет наличие слоя на чертеже
    Public Shared Function FuncFindLayerDwg(ByVal ActivDocument As Drawing, ByVal NameLayer As String) As DwgLayer
        FuncFindLayerDwg = Nothing
        If IsNothing(NameLayer) = False Then
            If NameLayer.Trim.Length > 0 Then
                If IsNothing(ActivDocument) = False Then
                    Try
                        Dim DwgLayers As DwgLayers = ActivDocument.Layers
                        For Each dwglayer As DwgLayer In DwgLayers
                            If dwglayer.Name Like NameLayer.Trim Then
                                Return dwglayer
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
            End If
        End If
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
