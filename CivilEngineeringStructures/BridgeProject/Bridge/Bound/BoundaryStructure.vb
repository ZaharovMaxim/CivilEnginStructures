Imports System.ComponentModel
Imports CivilEnginStructures.BeamI
Imports CivilEnginStructures.Bridges
Imports Microsoft.Office.Interop.Excel
Imports NetTopologySuite.Operation.Buffer
Imports Newtonsoft.Json.Linq
Imports Topomatic.Alg
Imports Topomatic.Alg.Road.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Dwg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Public Class BoundaryStructure
    ' Поля класса
    Private _name As String
    ' Конструктор по умолчанию
    Public Sub New()
        _name = ""
    End Sub
    'создать новую структуру для оси опоры
    Public Shared Function createBoundry(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.OtherElements
        elementAxis.ClassObject = StructureElement.classStructure.OtherObject
        elementAxis.Name = StructureElement.typeObject.boundaresBridge
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.boundaresBridge)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgPolyline3D
        Return elementAxis
    End Function
    'ищет существующую границу сооружения
    Public Shared Function getBoundaryBridge(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional removeDictionary As Boolean = False) As StructureElement
        Dim dataAxisPillar As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.boundaresBridge) = True Then
            Dim listBoundary = dictionaryObjectsBridge.Item(StructureElement.typeObject.boundaresBridge)
            If IsNothing(listBoundary) = False Then
                If listBoundary.Count > 0 Then
                    For k As Integer = 0 To listBoundary.Count - 1
                        Dim tempData As StructureElement = listBoundary.Item(k)
                        If IsNothing(tempData.DWGEntity) = False Then
                            Dim poly3d As DwgPolyline = tempData.DWGEntity
                            If poly3d.Area > 0 Then
                                dataAxisPillar = tempData
                                If removeDictionary = True Then
                                    listBoundary.RemoveAt(k)
                                End If
                                Exit For
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataAxisPillar
    End Function
    '=========================================================================================================
    'функция оформляет и создает границу мостового сооружения и ось трассы автодороги
    Public Shared Function drawAxisAndBoundaryBridge(ByRef drawingDocument As Topomatic.Dwg.Drawing, ByRef dataBridge As StructureElement, ByVal dictionaryBridgeBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal projectAlignment As Alignment, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleBoundary As ProjectCivilStructuresStyle = Nothing) As Boolean
        drawAxisAndBoundaryBridge = False
        If IsNothing(drawingDocument) = True Then Return False
        If IsNothing(dataBridge) = True Then Return False
        If IsNothing(dictionaryBridgeBeams) = True Then Return False
        If dictionaryBridgeBeams.Count = 0 Then Return False
        If IsNothing(projectAlignment) = True Then Return False
        Dim boundStructure As DwgPolyline = New DwgPolyline()
        Dim dataBoundares As StructureElement = getBoundaryBridge(dictionaryBridgeElements, True)
        If IsNothing(dataBoundares) = False Then
            boundStructure = dataBoundares.DWGEntity
        Else
            dataBoundares = createBoundry(dataBridge.IdStructure)
        End If
        Dim userBridge As Bridges = dataBridge.getBridge
        If IsNothing(styleBoundary) = True Then
            styleBoundary = New ProjectCivilStructuresStyle(drawingDocument)
        End If
        'граница габарита моста слева
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, -1 * userBridge.LeftStructureWidth + userBridge.TransverseOffset)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)
        'граница габарита моста справа
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, userBridge.RightStructureWidth + userBridge.TransverseOffset)
        axisPlineReverse = FuncAlignment.getReverseDwgPolyline(axisPlineReverse)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)
        Dim axisCentrePline As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, 0)
        Dim axisCentrePline3D As IPolyline3D = New Polyline3D()
        axisCentrePline.GetPolyline(axisCentrePline3D)
        Try
            Dim firstProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.First.Value
            Dim startFirtPoint As Vector2D = Nothing
            Dim secondFirtPoint As Vector2D = Nothing
            Dim pkStartPoint As Double = 9999999999
            Dim pkEndPoint As Double = -9999999999
            Dim off As Double = 0
            Dim topFirstBeam As StructureElement = firstProlet.First.Value
            If IsNothing(topFirstBeam) = True Then Return False
            Dim userBeam1 As BeamI = topFirstBeam.getBeamI
            'восстанавливаем 1 балку
            Dim startSectionPoint1 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim endSectionPoint1 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
            Dim restoreBeam1 As Boolean = CalculationBeams.restoreElementsBeam(topFirstBeam.DWGEntity, userBeam1, startSectionPoint1, endSectionPoint1, CalculationBeams.rectoreBeam.fullBeam)
            If firstProlet.Count > 1 Then
                Dim topLastBeam As StructureElement = firstProlet.Last.Value
                If IsNothing(topLastBeam) = True Then Return False
                Dim userBeam2 As BeamI = topLastBeam.getBeamI
                'восстанавливаем 2 балку
                Dim startSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam2 As Boolean = CalculationBeams.restoreElementsBeam(topLastBeam.DWGEntity, userBeam2, startSectionPoint2, endSectionPoint2, CalculationBeams.rectoreBeam.fullBeam)
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(0).Pos, startSectionPoint2.Item(0).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(0).Pos
                        secondFirtPoint = startSectionPoint2.Item(0).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(1).Pos, startSectionPoint2.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(1).Pos
                        secondFirtPoint = startSectionPoint2.Item(1).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(2).Pos, startSectionPoint2.Item(2).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(2).Pos
                        secondFirtPoint = startSectionPoint2.Item(2).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(3).Pos, startSectionPoint2.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(3).Pos
                        secondFirtPoint = startSectionPoint2.Item(3).Pos
                        pkStartPoint = pk
                    End If
                End If
            ElseIf firstProlet.Count = 1 Then
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(0).Pos, startSectionPoint1.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(0).Pos
                        secondFirtPoint = startSectionPoint1.Item(1).Pos
                        pkStartPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, startSectionPoint1.Item(2).Pos, startSectionPoint1.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk < pkStartPoint Then
                        startFirtPoint = startSectionPoint1.Item(2).Pos
                        secondFirtPoint = startSectionPoint1.Item(3).Pos
                        pkStartPoint = pk
                    End If
                End If
            End If
            Dim LastProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.Last.Value
            Dim startLastPoint As Vector2D = Nothing
            Dim secondLastPoint As Vector2D = Nothing
            Dim downLastBeam As StructureElement = LastProlet.First.Value
            If IsNothing(downLastBeam) = True Then Return False
            userBeam1 = downLastBeam.getBeamI
            'восстанавливаем 1 балку
            startSectionPoint1 = New List(Of Topomatic.Cad.Foundation.Vector3D)
            endSectionPoint1 = New List(Of Topomatic.Cad.Foundation.Vector3D)
            restoreBeam1 = CalculationBeams.restoreElementsBeam(downLastBeam.DWGEntity, userBeam1, startSectionPoint1, endSectionPoint1, CalculationBeams.rectoreBeam.fullBeam)
            If LastProlet.Count > 1 Then
                downLastBeam = LastProlet.Last.Value
                If IsNothing(downLastBeam) = True Then Return False
                Dim userBeam2 As BeamI = downLastBeam.getBeamI
                'восстанавливаем 2 балку
                Dim startSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint2 As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam2 As Boolean = CalculationBeams.restoreElementsBeam(downLastBeam.DWGEntity, userBeam2, startSectionPoint2, endSectionPoint2, CalculationBeams.rectoreBeam.fullBeam)
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(0).Pos, endSectionPoint2.Item(0).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(0).Pos
                        secondLastPoint = endSectionPoint2.Item(0).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(1).Pos, endSectionPoint2.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(1).Pos
                        secondLastPoint = endSectionPoint2.Item(1).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(2).Pos, endSectionPoint2.Item(2).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(2).Pos
                        secondLastPoint = endSectionPoint2.Item(2).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(3).Pos, endSectionPoint2.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(3).Pos
                        secondLastPoint = endSectionPoint2.Item(3).Pos
                        pkEndPoint = pk
                    End If
                End If
            ElseIf firstProlet.Count = 1 Then
                Dim startAlignPoint As Vector2D = New Vector2D(0, 0)
                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(0).Pos, endSectionPoint1.Item(1).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(0).Pos
                        secondLastPoint = endSectionPoint1.Item(1).Pos
                        pkEndPoint = pk
                    End If
                End If

                pointIntersectCollection = PolylineExtentions.GetIntersections(axisCentrePline3D, endSectionPoint1.Item(2).Pos, endSectionPoint1.Item(3).Pos)
                If pointIntersectCollection.Count > 0 Then
                    startAlignPoint = pointIntersectCollection.ElementAt(0)
                    Dim pk As Double = 0
                    Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisCentrePline3D, startAlignPoint, pk, off)
                    If pk > pkEndPoint Then
                        startLastPoint = endSectionPoint1.Item(2).Pos
                        secondLastPoint = endSectionPoint1.Item(3).Pos
                        pkEndPoint = pk
                    End If
                End If
            End If

            Dim pt1 As Vector3D = New Vector3D(startFirtPoint, 0)
            Dim pt2 As Vector3D = New Vector3D(secondFirtPoint, 0)
            Dim boolExtendLine1 As Boolean = MathFunction.FuncExtendPos(pt1, pt2, 100, 100)

            Dim pt3 As Vector3D = New Vector3D(startLastPoint, 0)
            Dim pt4 As Vector3D = New Vector3D(secondLastPoint, 0)
            Dim boolExtendLine2 As Boolean = MathFunction.FuncExtendPos(pt3, pt4, 100, 100)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем пересечение с левыс и правым габаритом
            Dim pkStartDirect As Double = -999999999
            Dim pkEndDirect As Double = -999999999
            Dim pkStartReverse As Double = -999999999
            Dim pkEndReverse As Double = -999999999
            Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, pt1.Pos, pt2.Pos)
            If pointIntersectCollection1.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection1.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkStartDirect, off)
            End If
            Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, pt1.Pos, pt2.Pos)
            If pointIntersectCollection2.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection2.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkStartReverse, off)
            End If
            Dim pointIntersectCollection3 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, pt3.Pos, pt4.Pos)
            If pointIntersectCollection3.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection3.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkEndDirect, off)
            End If
            Dim pointIntersectCollection4 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, pt3.Pos, pt4.Pos)
            If pointIntersectCollection4.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection4.ElementAt(0) '
                Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkEndReverse, off)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'обрезаем полилинии
            If pkStartDirect <> -999999999 And pkStartReverse <> -999999999 And pkEndDirect <> -999999999 And pkEndReverse <> -999999999 Then
                Dim pline2dCurv1 As Polyline2DCurve = New Polyline2DCurve()
                'pline2dCurv1.Add(New BugleVector2D(pointIntersectCollection1.ElementAt(0), 0))
                For i As Double = pkStartDirect To pkEndDirect Step 1
                    Dim pt As Vector2D = axisPline3DDirect.StaOffsetToPos(i, 0)
                    pline2dCurv1.Add(New BugleVector2D(pt, 0))
                Next
                pline2dCurv1.Add(New BugleVector2D(pointIntersectCollection3.ElementAt(0), 0))

                Dim pline2dCurv2 As Polyline2DCurve = New Polyline2DCurve()
                For i As Double = pkEndReverse To pkStartReverse Step 1
                    Dim pt As Vector2D = axisPline3DReverse.StaOffsetToPos(i, 0)
                    pline2dCurv2.Add(New BugleVector2D(pt, 0))
                Next
                pline2dCurv2.Add(New BugleVector2D(pointIntersectCollection2.ElementAt(0), 0))

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ось трассы
                Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
                For i As Integer = 0 To axisCentrePline.Count - 1
                    pline2dCurv.Add(axisCentrePline.Item(i))
                Next i
                Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(pkStartPoint, pkEndPoint)
                Dim pline2dCurv3 As Polyline2DCurve = New Polyline2DCurve()
                If IsArray(arrayPlineCurve) = True Then
                    For i As Integer = 0 To arrayPlineCurve.Length - 1
                        Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                        If Math.Abs(plineCurv.Length - (pkEndPoint - pkStartPoint)) <= 0.01 Then
                            pline2dCurv3 = plineCurv
                        End If
                    Next i
                End If

                If boundStructure.Count > 0 Then
                    boundStructure.Clear()
                End If
                If pline2dCurv2.Length > 0 And pline2dCurv1.Length > 0 Then
                    For i As Integer = 0 To pline2dCurv1.Count - 1
                        boundStructure.Add(pline2dCurv1.Item(i))
                    Next i
                    Dim count As Integer = pline2dCurv2.Count - 1
                    For i As Integer = 0 To pline2dCurv2.Count - 1
                        Dim bulg As Double = 0
                        If i <> 0 Then
                            bulg = pline2dCurv2.Item(i - 1).Bugle
                        End If
                        boundStructure.Add(pline2dCurv2.Item(i))
                    Next i
                    boundStructure.Closed = True
                End If
                If boundStructure.Count > 3 Then
                    drawingDocument.ActiveSpace.Entities.Add(boundStructure)
                    styleBoundary.setObjectStyle(boundStructure)
                End If
                dataBoundares.DWGEntity = boundStructure
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(boundStructure, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBoundares)
                Return True
            Else
                Return False
            End If
        Catch ex As System.Exception
        End Try
    End Function
End Class
