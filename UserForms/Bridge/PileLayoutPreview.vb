Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Linq

Friend NotInheritable Class PileLayoutGeometry
    Private Const StationTolerance As Double = 0.0005R

    Private ReadOnly localCorners As Double(,)
    Private ReadOnly localPiles As Double(,)
    Private ReadOnly localEdges As Double(,)
    Private ReadOnly horizontalSegments As Double()
    Private ReadOnly verticalSegments As Double()
    Private ReadOnly horizontalStations As Double()
    Private ReadOnly verticalStations As Double()
    Private ReadOnly pileColumnIds As Double()

    Private Sub New(corners As Double(,), piles As Double(,), edges As Double(,), ready As Boolean,
                    horizontal As Double(), vertical As Double(),
                    horizontalAxes As Double(), verticalAxes As Double(),
                    columnIds As Double())
        localCorners = corners
        localPiles = piles
        localEdges = edges
        IsReady = ready
        horizontalSegments = horizontal
        verticalSegments = vertical
        horizontalStations = horizontalAxes
        verticalStations = verticalAxes
        pileColumnIds = columnIds
    End Sub

    Friend Shared Function Create(corners As Double(,), piles As Double(,), edges As Double(,)) As PileLayoutGeometry
        Return CreateCore(corners, piles, edges, Nothing, Nothing)
    End Function

    Friend Shared Function Create(corners As Double(,), piles As Double(,), edges As Double(,),
                                  axisData As Double(,), nominalSize As Double()) As PileLayoutGeometry
        Return CreateCore(corners, piles, edges, axisData, nominalSize)
    End Function

    Private Shared Function CreateCore(corners As Double(,), piles As Double(,), edges As Double(,),
                                       axisData As Double(,), nominalSize As Double()) As PileLayoutGeometry
        If Not HasShape(corners, 4, 3) OrElse Not HasPileShape(piles) OrElse Not HasEdgeShape(edges) Then
            Return EmptyGeometry()
        End If
        If Not AllFinite(corners) OrElse Not AllFinite(piles) OrElse Not AllFinite(edges) Then
            Return EmptyGeometry()
        End If
        Dim validPileRows As Integer() = Enumerable.Range(0, piles.GetLength(0)).Where(
            Function(row)
                Dim dx As Double = piles(row, 3) - piles(row, 0)
                Dim dy As Double = piles(row, 4) - piles(row, 1)
                Dim dz As Double = piles(row, 5) - piles(row, 2)
                Return dx * dx + dy * dy + dz * dz > 0.000000000001R
            End Function).ToArray()
        If validPileRows.Length = 0 Then Return EmptyGeometry()

        Dim originX As Double = corners(0, 0)
        Dim originY As Double = corners(0, 1)
        Dim edgeX As Double = corners(1, 0) - originX
        Dim edgeY As Double = corners(1, 1) - originY
        Dim edgeLength As Double = Math.Sqrt(edgeX * edgeX + edgeY * edgeY)
        If edgeLength <= Double.Epsilon Then Return EmptyGeometry()

        Dim ux As Double = edgeX / edgeLength
        Dim uy As Double = edgeY / edgeLength
        Dim vx As Double = -uy
        Dim vy As Double = ux
        If (corners(3, 0) - originX) * vx + (corners(3, 1) - originY) * vy < 0.0R Then
            vx = -vx
            vy = -vy
        End If

        Dim normalizedCorners As Double(,) = CreateMatrix(4, 3)
        For row As Integer = 0 To 3
            ProjectXY(corners(row, 0), corners(row, 1), originX, originY, ux, uy, vx, vy,
                      normalizedCorners(row, 0), normalizedCorners(row, 1))
            normalizedCorners(row, 2) = corners(row, 2)
        Next
        If Math.Abs(PolygonArea(normalizedCorners)) <= 0.000000001R Then Return EmptyGeometry()

        Dim normalizedPiles As Double(,) = CreateMatrix(validPileRows.Length, 8)
        For row As Integer = 0 To validPileRows.Length - 1
            Dim sourceRow As Integer = validPileRows(row)
            ProjectXY(piles(sourceRow, 0), piles(sourceRow, 1), originX, originY, ux, uy, vx, vy,
                      normalizedPiles(row, 0), normalizedPiles(row, 1))
            normalizedPiles(row, 2) = piles(sourceRow, 2)
            ProjectXY(piles(sourceRow, 3), piles(sourceRow, 4), originX, originY, ux, uy, vx, vy,
                      normalizedPiles(row, 3), normalizedPiles(row, 4))
            normalizedPiles(row, 5) = piles(sourceRow, 5)
            normalizedPiles(row, 6) = Math.Max(0.0R, piles(sourceRow, 6))
            normalizedPiles(row, 7) = Math.Max(0.0R, piles(sourceRow, 7))
        Next

        Dim normalizedEdges As Double(,) = CreateMatrix(If(edges Is Nothing, 0, edges.GetLength(0)), 6)
        If edges IsNot Nothing Then
            For row As Integer = 0 To edges.GetLength(0) - 1
                ProjectXY(edges(row, 0), edges(row, 1), originX, originY, ux, uy, vx, vy,
                          normalizedEdges(row, 0), normalizedEdges(row, 1))
                normalizedEdges(row, 2) = edges(row, 2)
                ProjectXY(edges(row, 3), edges(row, 4), originX, originY, ux, uy, vx, vy,
                          normalizedEdges(row, 3), normalizedEdges(row, 4))
                normalizedEdges(row, 5) = edges(row, 5)
            Next
        End If

        Dim horizontalAxes As Double() = Nothing
        Dim verticalAxes As Double() = Nothing
        Dim columnIds As Double() = Nothing
        If TryCreateNominalStations(axisData, nominalSize, piles.GetLength(0), validPileRows, normalizedCorners,
                                    normalizedPiles, horizontalAxes, verticalAxes, columnIds) Then
            ' Nominal scheme data is used only for dimension chains and column numbers.
        Else
            horizontalAxes = CreateDimensionStations(
                Enumerable.Range(0, 4).Select(Function(row) normalizedCorners(row, 0)),
                Enumerable.Range(0, normalizedPiles.GetLength(0)).Select(Function(row) normalizedPiles(row, 0)))
            verticalAxes = CreateDimensionStations(
                Enumerable.Range(0, 4).Select(Function(row) normalizedCorners(row, 1)),
                Enumerable.Range(0, normalizedPiles.GetLength(0)).Select(Function(row) normalizedPiles(row, 1)))
            columnIds = CreateSequentialColumnIds(normalizedPiles)
        End If

        Return New PileLayoutGeometry(normalizedCorners, normalizedPiles, normalizedEdges, True,
                                      CreateDimensionSegments(horizontalAxes),
                                      CreateDimensionSegments(verticalAxes),
                                      horizontalAxes, verticalAxes, columnIds)
    End Function

    Friend ReadOnly Property IsReady As Boolean

    Friend ReadOnly Property PileCount As Integer
        Get
            Return localPiles.GetLength(0)
        End Get
    End Property

    Friend ReadOnly Property Corners As Double(,)
        Get
            Return DirectCast(localCorners.Clone(), Double(,))
        End Get
    End Property

    Friend ReadOnly Property Piles As Double(,)
        Get
            Return DirectCast(localPiles.Clone(), Double(,))
        End Get
    End Property

    Friend ReadOnly Property Edges As Double(,)
        Get
            Return DirectCast(localEdges.Clone(), Double(,))
        End Get
    End Property

    Friend ReadOnly Property HorizontalSegmentsMillimetres As Double()
        Get
            Return DirectCast(horizontalSegments.Clone(), Double())
        End Get
    End Property

    Friend ReadOnly Property VerticalSegmentsMillimetres As Double()
        Get
            Return DirectCast(verticalSegments.Clone(), Double())
        End Get
    End Property

    Friend Function HorizontalStationValues() As Double()
        Return DirectCast(horizontalStations.Clone(), Double())
    End Function

    Friend Function VerticalStationValues() As Double()
        Return DirectCast(verticalStations.Clone(), Double())
    End Function

    Friend Function PileColumnIdValues() As Double()
        Return DirectCast(pileColumnIds.Clone(), Double())
    End Function

    Private Shared Function EmptyGeometry() As PileLayoutGeometry
        Return New PileLayoutGeometry(CreateMatrix(0, 3), CreateMatrix(0, 8), CreateMatrix(0, 6), False,
                                      Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                      Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                      Array.Empty(Of Double)())
    End Function

    Private Shared Function TryCreateNominalStations(axisData As Double(,), nominalSize As Double(),
                                                     sourcePileCount As Integer,
                                                     validPileRows As Integer(),
                                                     normalizedCorners As Double(,), normalizedPiles As Double(,),
                                                     ByRef horizontalAxes As Double(),
                                                     ByRef verticalAxes As Double(),
                                                     ByRef columnIds As Double()) As Boolean
        Const coordinateTolerance As Double = 0.005R
        Const groupTolerance As Double = 0.0000001R
        If axisData Is Nothing OrElse axisData.Rank <> 2 OrElse axisData.GetLength(1) <> 4 OrElse
           axisData.GetLength(0) <> sourcePileCount OrElse nominalSize Is Nothing OrElse nominalSize.Length <> 2 OrElse
           Not IsFinite(nominalSize(0)) OrElse Not IsFinite(nominalSize(1)) OrElse
           nominalSize(0) <= 0.0R OrElse nominalSize(1) <= 0.0R Then Return False

        Dim selected(validPileRows.Length - 1, 3) As Double
        For row As Integer = 0 To validPileRows.Length - 1
            Dim sourceRow As Integer = validPileRows(row)
            For column As Integer = 0 To 3
                Dim value As Double = axisData(sourceRow, column)
                If Not IsFinite(value) Then Return False
                selected(row, column) = value
            Next
            If Math.Abs(selected(row, 0) - Math.Round(selected(row, 0))) > groupTolerance OrElse
               Math.Abs(selected(row, 1) - Math.Round(selected(row, 1))) > groupTolerance OrElse
               Math.Abs(selected(row, 2) - normalizedPiles(row, 0)) > coordinateTolerance OrElse
               Math.Abs(selected(row, 3) - normalizedPiles(row, 1)) > coordinateTolerance Then Return False
        Next

        Dim minCornerX As Double = Enumerable.Range(0, 4).Min(Function(row) normalizedCorners(row, 0))
        Dim maxCornerX As Double = Enumerable.Range(0, 4).Max(Function(row) normalizedCorners(row, 0))
        Dim minCornerY As Double = Enumerable.Range(0, 4).Min(Function(row) normalizedCorners(row, 1))
        Dim maxCornerY As Double = Enumerable.Range(0, 4).Max(Function(row) normalizedCorners(row, 1))
        If Math.Abs(minCornerX) > coordinateTolerance OrElse Math.Abs(minCornerY) > coordinateTolerance OrElse
           Math.Abs(maxCornerX - nominalSize(0)) > coordinateTolerance OrElse
           Math.Abs(maxCornerY - nominalSize(1)) > coordinateTolerance Then Return False

        For Each group As IGrouping(Of Double, Integer) In
            Enumerable.Range(0, selected.GetLength(0)).GroupBy(Function(row) selected(row, 1))
            Dim firstValue As Double = selected(group.First(), 2)
            If group.Any(Function(row) Math.Abs(selected(row, 2) - firstValue) > groupTolerance) Then Return False
        Next
        For Each group As IGrouping(Of Double, Integer) In
            Enumerable.Range(0, selected.GetLength(0)).GroupBy(Function(row) selected(row, 0))
            Dim firstValue As Double = selected(group.First(), 3)
            If group.Any(Function(row) Math.Abs(selected(row, 3) - firstValue) > groupTolerance) Then Return False
        Next

        horizontalAxes = CreateDimensionStations(New Double() {0.0R, nominalSize(0)},
            Enumerable.Range(0, selected.GetLength(0)).Select(Function(row) selected(row, 2)))
        verticalAxes = CreateDimensionStations(New Double() {0.0R, nominalSize(1)},
            Enumerable.Range(0, selected.GetLength(0)).Select(Function(row) selected(row, 3)))
        columnIds = Enumerable.Range(0, selected.GetLength(0)).Select(Function(row) selected(row, 1)).ToArray()
        Return True
    End Function

    Private Shared Function CreateSequentialColumnIds(piles As Double(,)) As Double()
        Dim stations As Double() = CreateDimensionStations(Array.Empty(Of Double)(),
            Enumerable.Range(0, piles.GetLength(0)).Select(Function(row) piles(row, 0)))
        Dim result(piles.GetLength(0) - 1) As Double
        For row As Integer = 0 To piles.GetLength(0) - 1
            Dim pileX As Double = piles(row, 0)
            Dim stationIndex As Integer = Array.FindIndex(stations,
                Function(value) Math.Abs(value - pileX) <= StationTolerance)
            result(row) = stationIndex + 1
        Next
        Return result
    End Function

    Private Shared Function HasShape(source As Double(,), rows As Integer, columns As Integer) As Boolean
        Return source IsNot Nothing AndAlso source.Rank = 2 AndAlso
               source.GetLength(0) = rows AndAlso source.GetLength(1) = columns
    End Function

    Private Shared Function HasPileShape(source As Double(,)) As Boolean
        Return source IsNot Nothing AndAlso source.Rank = 2 AndAlso
               source.GetLength(0) > 0 AndAlso source.GetLength(1) = 8
    End Function

    Private Shared Function HasEdgeShape(source As Double(,)) As Boolean
        Return source Is Nothing OrElse (source.Rank = 2 AndAlso source.GetLength(1) = 6)
    End Function

    Private Shared Function AllFinite(source As Double(,)) As Boolean
        If source Is Nothing Then Return True
        For row As Integer = 0 To source.GetLength(0) - 1
            For column As Integer = 0 To source.GetLength(1) - 1
                If Not IsFinite(source(row, column)) Then Return False
            Next
        Next
        Return True
    End Function

    Private Shared Function CreateMatrix(rows As Integer, columns As Integer) As Double(,)
        Return DirectCast(Array.CreateInstance(GetType(Double), rows, columns), Double(,))
    End Function

    Private Shared Sub ProjectXY(x As Double, y As Double, originX As Double, originY As Double,
                                 ux As Double, uy As Double, vx As Double, vy As Double,
                                 ByRef resultX As Double, ByRef resultY As Double)
        Dim dx As Double = x - originX
        Dim dy As Double = y - originY
        resultX = dx * ux + dy * uy
        resultY = dx * vx + dy * vy
    End Sub

    Private Shared Function PolygonArea(points As Double(,)) As Double
        Dim result As Double = 0.0R
        For index As Integer = 0 To 3
            Dim nextIndex As Integer = (index + 1) Mod 4
            result += points(index, 0) * points(nextIndex, 1) -
                      points(nextIndex, 0) * points(index, 1)
        Next
        Return result / 2.0R
    End Function

    Private Shared Function CreateDimensionStations(edges As IEnumerable(Of Double),
                                                     centres As IEnumerable(Of Double)) As Double()
        Dim edgeValues As Double() = edges.Where(AddressOf IsFinite).ToArray()
        Dim stations As New List(Of Double)()
        If edgeValues.Length > 0 Then
            stations.Add(edgeValues.Min())
            stations.Add(edgeValues.Max())
        End If
        stations.AddRange(centres.Where(AddressOf IsFinite))
        stations.Sort()

        Dim unique As New List(Of Double)()
        For Each station As Double In stations
            If unique.Count = 0 OrElse Math.Abs(station - unique(unique.Count - 1)) > StationTolerance Then
                unique.Add(station)
            End If
        Next
        Return unique.ToArray()
    End Function

    Private Shared Function CreateDimensionSegments(stations As Double()) As Double()
        Dim result As New List(Of Double)()
        For index As Integer = 1 To stations.Length - 1
            Dim interval As Double = stations(index) - stations(index - 1)
            If interval > StationTolerance Then result.Add(interval * 1000.0R)
        Next
        Return result.ToArray()
    End Function

    Private Shared Function IsFinite(value As Double) As Boolean
        Return Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value)
    End Function
End Class

Friend NotInheritable Class PileLayoutResult
    Private ReadOnly frontHeads As PointF()
    Private ReadOnly frontToes As PointF()
    Private ReadOnly planHeads As PointF()
    Private ReadOnly frontCorners As PointF()
    Private ReadOnly planCorners As PointF()
    Private ReadOnly labelBounds As RectangleF()
    Private ReadOnly pileNumbers As String()
    Private ReadOnly pileNumberBounds As RectangleF()

    Friend Sub New(scale As Double, heads As PointF(), toes As PointF(), plan As PointF(),
                   frontGrillage As PointF(), planGrillage As PointF(), labels As RectangleF(),
                   xMinimum As Double, yMaximum As Double, zMinimum As Double,
                   xOffset As Double, frontOffsetY As Double, planOffsetY As Double,
                   planBottom As Single, annotations As List(Of PileLayoutPreview.DimensionAnnotation),
                   hasBreak As Boolean, breakTop As Double, breakBottom As Double,
                   numbers As String(), numberBounds As RectangleF())
        PixelsPerMetre = scale
        frontHeads = heads
        frontToes = toes
        planHeads = plan
        frontCorners = frontGrillage
        planCorners = planGrillage
        labelBounds = labels
        Me.XMinimum = xMinimum
        Me.YMaximum = yMaximum
        Me.ZMinimum = zMinimum
        Me.XOffset = xOffset
        Me.FrontOffsetY = frontOffsetY
        Me.PlanOffsetY = planOffsetY
        Me.PlanBottom = planBottom
        Me.Annotations = annotations
        HasPileBreak = hasBreak
        BreakTopZ = breakTop
        BreakBottomZ = breakBottom
        pileNumbers = numbers
        pileNumberBounds = numberBounds
    End Sub

    Friend ReadOnly Property PixelsPerMetre As Double
    Friend ReadOnly Property XMinimum As Double
    Friend ReadOnly Property YMaximum As Double
    Friend ReadOnly Property ZMinimum As Double
    Friend ReadOnly Property XOffset As Double
    Friend ReadOnly Property FrontOffsetY As Double
    Friend ReadOnly Property PlanOffsetY As Double
    Friend ReadOnly Property PlanBottom As Single
    Friend ReadOnly Property Annotations As List(Of PileLayoutPreview.DimensionAnnotation)
    Friend ReadOnly Property HasPileBreak As Boolean
    Friend ReadOnly Property BreakTopZ As Double
    Friend ReadOnly Property BreakBottomZ As Double

    Friend ReadOnly Property FrontPileHeads As PointF()
        Get
            Return DirectCast(frontHeads.Clone(), PointF())
        End Get
    End Property

    Friend ReadOnly Property FrontPileToes As PointF()
        Get
            Return DirectCast(frontToes.Clone(), PointF())
        End Get
    End Property

    Friend ReadOnly Property PlanPileHeads As PointF()
        Get
            Return DirectCast(planHeads.Clone(), PointF())
        End Get
    End Property

    Friend ReadOnly Property FrontGrillageCorners As PointF()
        Get
            Return DirectCast(frontCorners.Clone(), PointF())
        End Get
    End Property

    Friend ReadOnly Property PlanGrillageCorners As PointF()
        Get
            Return DirectCast(planCorners.Clone(), PointF())
        End Get
    End Property

    Friend ReadOnly Property DimensionLabelBounds As RectangleF()
        Get
            Return DirectCast(labelBounds.Clone(), RectangleF())
        End Get
    End Property

    Friend ReadOnly Property FrontPileNumbers As String()
        Get
            Return DirectCast(pileNumbers.Clone(), String())
        End Get
    End Property

    Friend ReadOnly Property FrontPileNumberBounds As RectangleF()
        Get
            Return DirectCast(pileNumberBounds.Clone(), RectangleF())
        End Get
    End Property
End Class

Friend NotInheritable Class PileLayoutPreview
    Friend NotInheritable Class DimensionAnnotation
        Friend ReadOnly Text As String
        Friend ReadOnly Bounds As RectangleF
        Friend ReadOnly StartPoint As PointF
        Friend ReadOnly EndPoint As PointF
        Friend ReadOnly Horizontal As Boolean

        Friend Sub New(text As String, bounds As RectangleF, startPoint As PointF,
                       endPoint As PointF, horizontal As Boolean)
            Me.Text = text
            Me.Bounds = bounds
            Me.StartPoint = startPoint
            Me.EndPoint = endPoint
            Me.Horizontal = horizontal
        End Sub
    End Class

    Private Const LabelHeight As Single = 15.0!
    Private Const LabelGap As Single = 2.0!
    Private Const BreakDisplayHeight As Double = 0.4R

    Private Sub New()
    End Sub

    Friend Shared Function CreateLayout(geometry As PileLayoutGeometry,
                                        requestedSize As Size) As PileLayoutResult
        If geometry Is Nothing OrElse Not geometry.IsReady Then Return EmptyLayout()

        Dim width As Integer = Math.Max(1, If(requestedSize.Width > 0, requestedSize.Width, 640))
        Dim height As Integer = Math.Max(1, If(requestedSize.Height > 0, requestedSize.Height, 480))
        Dim corners As Double(,) = geometry.Corners
        Dim piles As Double(,) = geometry.Piles
        Dim edges As Double(,) = geometry.Edges

        Dim minX As Double = corners(0, 0)
        Dim maxX As Double = corners(0, 0)
        Dim minY As Double = corners(0, 1)
        Dim maxY As Double = corners(0, 1)
        Dim minZ As Double = corners(0, 2)
        Dim maxZ As Double = corners(0, 2)
        For row As Integer = 0 To corners.GetLength(0) - 1
            Include(corners(row, 0), minX, maxX)
            Include(corners(row, 1), minY, maxY)
            Include(corners(row, 2), minZ, maxZ)
        Next
        For row As Integer = 0 To piles.GetLength(0) - 1
            Dim radius As Double = Math.Max(piles(row, 6), piles(row, 7)) / 2.0R
            Include(piles(row, 0) - radius, minX, maxX)
            Include(piles(row, 0) + radius, minX, maxX)
            Include(piles(row, 3) - radius, minX, maxX)
            Include(piles(row, 3) + radius, minX, maxX)
            Include(piles(row, 1) - radius, minY, maxY)
            Include(piles(row, 1) + radius, minY, maxY)
            Include(piles(row, 2) - radius, minZ, maxZ)
            Include(piles(row, 2) + radius, minZ, maxZ)
            Include(piles(row, 5) - radius, minZ, maxZ)
            Include(piles(row, 5) + radius, minZ, maxZ)
        Next
        For row As Integer = 0 To edges.GetLength(0) - 1
            Include(edges(row, 0), minX, maxX)
            Include(edges(row, 3), minX, maxX)
            Include(edges(row, 1), minY, maxY)
            Include(edges(row, 4), minY, maxY)
            Include(edges(row, 2), minZ, maxZ)
            Include(edges(row, 5), minZ, maxZ)
        Next

        Dim xSpan As Double = Math.Max(0.001R, maxX - minX)
        Dim ySpan As Double = Math.Max(0.001R, maxY - minY)
        Dim zSpan As Double = Math.Max(0.001R, maxZ - minZ)
        Dim breakTop As Double = 0.0R
        Dim breakBottom As Double = 0.0R
        Dim hasBreak As Boolean = TryFindPileBreak(corners, piles, edges, breakTop, breakBottom)
        Dim visibleZSpan As Double = zSpan
        If hasBreak Then visibleZSpan -= breakBottom - breakTop - BreakDisplayHeight
        Dim leftMargin As Double = Math.Min(92.0R, Math.Max(54.0R, width * 0.16R))
        Dim rightMargin As Double = Math.Min(18.0R, Math.Max(6.0R, width * 0.025R))
        Dim topMargin As Double = Math.Min(24.0R, Math.Max(18.0R, height * 0.045R))
        Dim sectionGap As Double = 36.0R
        Dim bottomReserve As Double = 54.0R
        Dim scale As Double = 0.0R
        Dim xOffset As Double = 0.0R
        Dim horizontalLanes As Integer = 1

        For pass As Integer = 0 To 5
            Dim plotWidth As Double = Math.Max(1.0R, width - leftMargin - rightMargin)
            Dim plotHeight As Double = Math.Max(1.0R, height - topMargin - sectionGap - bottomReserve)
            scale = Math.Max(0.000001R, Math.Min(plotWidth / xSpan, plotHeight / (ySpan + visibleZSpan)))
            xOffset = leftMargin + (plotWidth - xSpan * scale) / 2.0R
            horizontalLanes = CountHorizontalLanes(geometry, minX, scale, xOffset, width)
            Dim requiredReserve As Double = 16.0R + horizontalLanes * (LabelHeight + LabelGap)
            Dim numberLanes As Integer = CountPileNumberLanes(geometry, minX, scale, xOffset)
            Dim requiredSectionGap As Double = 20.0R + numberLanes * (LabelHeight + 1.0R)
            If Math.Abs(requiredReserve - bottomReserve) < 0.5R AndAlso
               Math.Abs(requiredSectionGap - sectionGap) < 0.5R Then Exit For
            bottomReserve = requiredReserve
            sectionGap = requiredSectionGap
        Next

        Dim frontOffsetY As Double = topMargin
        Dim frontBottom As Double = frontOffsetY + visibleZSpan * scale
        Dim planOffsetY As Double = frontBottom + sectionGap
        Dim planBottom As Single = CSng(planOffsetY + ySpan * scale)
        Dim frontHeads(piles.GetLength(0) - 1) As PointF
        Dim frontToes(piles.GetLength(0) - 1) As PointF
        Dim planHeads(piles.GetLength(0) - 1) As PointF
        For row As Integer = 0 To piles.GetLength(0) - 1
            frontHeads(row) = MapFront(piles(row, 0), piles(row, 2), minX, minZ,
                                       scale, xOffset, frontOffsetY, hasBreak, breakTop, breakBottom)
            frontToes(row) = MapFront(piles(row, 3), piles(row, 5), minX, minZ,
                                      scale, xOffset, frontOffsetY, hasBreak, breakTop, breakBottom)
            planHeads(row) = MapPlan(piles(row, 0), piles(row, 1), minX, maxY,
                                     scale, xOffset, planOffsetY)
        Next

        Dim frontCorners(3) As PointF
        Dim planCorners(3) As PointF
        For row As Integer = 0 To 3
            frontCorners(row) = MapFront(corners(row, 0), corners(row, 2), minX, minZ,
                                         scale, xOffset, frontOffsetY, hasBreak, breakTop, breakBottom)
            planCorners(row) = MapPlan(corners(row, 0), corners(row, 1), minX, maxY,
                                       scale, xOffset, planOffsetY)
        Next

        Dim annotations As New List(Of DimensionAnnotation)()
        AddHorizontalAnnotations(annotations, geometry, minX, scale, xOffset,
                                 planBottom, width, height)
        AddVerticalAnnotations(annotations, geometry, minY, maxY, scale, xOffset,
                               planOffsetY, width, height)
        Dim bounds As RectangleF() = annotations.Select(Function(item) item.Bounds).ToArray()
        Dim numbers As String() = Nothing
        Dim numberBounds As RectangleF() = Nothing
        CreatePileNumbers(geometry, frontHeads, frontToes, planOffsetY, numbers, numberBounds)

        Return New PileLayoutResult(scale, frontHeads, frontToes, planHeads,
                                    frontCorners, planCorners, bounds,
                                    minX, maxY, minZ, xOffset, frontOffsetY,
                                    planOffsetY, planBottom, annotations, hasBreak,
                                    breakTop, breakBottom, numbers, numberBounds)
    End Function

    Friend Shared Function Render(geometry As PileLayoutGeometry, requestedSize As Size) As Bitmap
        Dim width As Integer = Math.Max(1, If(requestedSize.Width > 0, requestedSize.Width, 640))
        Dim height As Integer = Math.Max(1, If(requestedSize.Height > 0, requestedSize.Height, 480))
        Dim bitmap As New Bitmap(width, height)
        Using graphics As Graphics = Graphics.FromImage(bitmap)
            graphics.Clear(Color.White)
            If geometry Is Nothing OrElse Not geometry.IsReady Then Return bitmap

            Dim layout As PileLayoutResult = CreateLayout(geometry, requestedSize)
            If layout.PixelsPerMetre <= 0.0R Then Return bitmap
            Dim corners As Double(,) = geometry.Corners
            Dim piles As Double(,) = geometry.Piles
            Dim edges As Double(,) = geometry.Edges
            Dim frontHeads As PointF() = layout.FrontPileHeads
            Dim frontToes As PointF() = layout.FrontPileToes
            Dim planHeads As PointF() = layout.PlanPileHeads
            Dim frontCorners As PointF() = layout.FrontGrillageCorners
            Dim planCorners As PointF() = layout.PlanGrillageCorners

            graphics.SmoothingMode = SmoothingMode.AntiAlias
            graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
            Using bodyPen As New Pen(Color.FromArgb(93, 109, 133), 1.0!),
                  grillagePen As New Pen(Color.FromArgb(32, 50, 77), 1.25!),
                  pilePen As New Pen(Color.FromArgb(37, 99, 235), 1.35!),
                  axisPen As New Pen(Color.FromArgb(150, 160, 176), 0.9!),
                  dimensionPen As New Pen(Color.FromArgb(75, 85, 99), 0.9!),
                  grillageBrush As New SolidBrush(Color.FromArgb(239, 242, 247)),
                  pileBrush As New SolidBrush(Color.FromArgb(232, 240, 254)),
                  textBrush As New SolidBrush(Color.FromArgb(32, 50, 77)),
                  textFont As New Font("Segoe UI", If(width < 520, 8.0!, 8.5!), FontStyle.Regular, GraphicsUnit.Point)

                graphics.FillPolygon(grillageBrush, planCorners)
                graphics.DrawPolygon(grillagePen, planCorners)
                DrawBodyEdges(graphics, bodyPen, edges, layout)
                DrawFrontGrillage(graphics, grillagePen, frontCorners)

                For row As Integer = 0 To piles.GetLength(0) - 1
                    DrawFrontPile(graphics, piles, row, layout, pilePen, pileBrush)
                    DrawPileSymbol(graphics, planHeads(row), piles(row, 6), piles(row, 7),
                                   layout.PixelsPerMetre, pilePen, axisPen, pileBrush)
                Next
                graphics.DrawString(If(layout.HasPileBreak, "Вид спереди · с разрывом", "Вид спереди"),
                                    textFont, textBrush, New PointF(4.0!, 2.0!))
                graphics.DrawString("План свай · размеры, мм", textFont, textBrush,
                                    New PointF(4.0!, CSng(Math.Max(2.0R, layout.PlanOffsetY - 14.0R))))
                DrawPileNumbers(graphics, layout, textFont, textBrush)
                DrawDimensions(graphics, layout, dimensionPen, textFont, textBrush)
            End Using
        End Using
        Return bitmap
    End Function

    Private Shared Function EmptyLayout() As PileLayoutResult
        Return New PileLayoutResult(0.0R, Array.Empty(Of PointF)(), Array.Empty(Of PointF)(),
                                    Array.Empty(Of PointF)(), Array.Empty(Of PointF)(),
                                    Array.Empty(Of PointF)(), Array.Empty(Of RectangleF)(),
                                    0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.0!,
                                    New List(Of DimensionAnnotation)(), False, 0.0R, 0.0R,
                                    Array.Empty(Of String)(), Array.Empty(Of RectangleF)())
    End Function

    Private Shared Sub Include(value As Double, ByRef minimum As Double, ByRef maximum As Double)
        minimum = Math.Min(minimum, value)
        maximum = Math.Max(maximum, value)
    End Sub

    Private Shared Function MapFront(x As Double, z As Double, minX As Double, minZ As Double,
                                     scale As Double, xOffset As Double, yOffset As Double,
                                     hasBreak As Boolean, breakTop As Double,
                                     breakBottom As Double) As PointF
        Dim visibleZ As Double = z
        If hasBreak AndAlso z >= breakBottom Then
            visibleZ -= breakBottom - breakTop - BreakDisplayHeight
        ElseIf hasBreak AndAlso z > breakTop Then
            visibleZ = breakTop + BreakDisplayHeight * (z - breakTop) / (breakBottom - breakTop)
        End If
        Return New PointF(CSng(xOffset + (x - minX) * scale),
                          CSng(yOffset + (visibleZ - minZ) * scale))
    End Function

    Private Shared Function MapPlan(x As Double, y As Double, minX As Double, maxY As Double,
                                    scale As Double, xOffset As Double, yOffset As Double) As PointF
        Return New PointF(CSng(xOffset + (x - minX) * scale),
                          CSng(yOffset + (maxY - y) * scale))
    End Function

    Private Shared Function TryFindPileBreak(corners As Double(,), piles As Double(,), edges As Double(,),
                                             ByRef breakTop As Double,
                                             ByRef breakBottom As Double) As Boolean
        Dim upperLimit As Double = Double.MinValue
        For row As Integer = 0 To corners.GetLength(0) - 1
            upperLimit = Math.Max(upperLimit, corners(row, 2))
        Next
        For row As Integer = 0 To edges.GetLength(0) - 1
            upperLimit = Math.Max(upperLimit, Math.Max(edges(row, 2), edges(row, 5)))
        Next

        Dim lowerLimit As Double = Double.MaxValue
        For row As Integer = 0 To piles.GetLength(0) - 1
            Dim radius As Double = Math.Max(piles(row, 6), piles(row, 7)) / 2.0R
            upperLimit = Math.Max(upperLimit, piles(row, 2) + radius)
            lowerLimit = Math.Min(lowerLimit, piles(row, 5) - radius)
        Next
        If upperLimit = Double.MinValue OrElse lowerLimit = Double.MaxValue Then Return False

        breakTop = upperLimit + 0.55R
        breakBottom = lowerLimit - 0.55R
        If breakBottom - breakTop <= 4.0R Then
            breakTop = 0.0R
            breakBottom = 0.0R
            Return False
        End If
        Return True
    End Function

    Private Shared Sub CreatePileNumbers(geometry As PileLayoutGeometry,
                                         frontHeads As PointF(), frontToes As PointF(),
                                         planOffsetY As Double,
                                         ByRef numbers As String(),
                                         ByRef bounds As RectangleF())
        Dim columnIds As Double() = geometry.PileColumnIdValues()
        Dim piles As Double(,) = geometry.Piles
        Dim groups = Enumerable.Range(0, columnIds.Length).
            GroupBy(Function(row) columnIds(row)).
            OrderBy(Function(group) group.Average(Function(row) piles(row, 0))).ToArray()
        ReDim numbers(groups.Length - 1)
        ReDim bounds(groups.Length - 1)
        Dim numberTop As Single = CSng(frontToes.Max(Function(point) point.Y) + 2.0!)
        Dim occupied As New List(Of List(Of RectangleF))()
        For index As Integer = 0 To groups.Length - 1
            Dim centreX As Single = CSng(groups(index).Average(Function(row) CDbl(frontHeads(row).X)))
            numbers(index) = FormatColumnId(groups(index).Key)
            Dim labelWidth As Single = MeasureNumberWidth(numbers(index))
            Dim probe As New RectangleF(centreX - labelWidth / 2.0!, 0.0!, labelWidth, LabelHeight)
            Dim lane As Integer = AssignLane(probe, occupied)
            Dim top As Single = numberTop + lane * (LabelHeight + 1.0!)
            top = Math.Min(top, CSng(planOffsetY - LabelHeight))
            bounds(index) = New RectangleF(centreX - labelWidth / 2.0!, top,
                                            labelWidth, LabelHeight)
        Next
    End Sub

    Private Shared Function CountPileNumberLanes(geometry As PileLayoutGeometry,
                                                 minX As Double, scale As Double,
                                                 xOffset As Double) As Integer
        Dim columnIds As Double() = geometry.PileColumnIdValues()
        Dim piles As Double(,) = geometry.Piles
        Dim groups = Enumerable.Range(0, columnIds.Length).
            GroupBy(Function(row) columnIds(row)).
            OrderBy(Function(group) group.Average(Function(row) piles(row, 0))).ToArray()
        Dim occupied As New List(Of List(Of RectangleF))()
        For Each group In groups
            Dim centreX As Single = CSng(xOffset + (group.Average(Function(row) piles(row, 0)) - minX) * scale)
            Dim text As String = FormatColumnId(group.Key)
            Dim probe As New RectangleF(centreX - MeasureNumberWidth(text) / 2.0!, 0.0!,
                                        MeasureNumberWidth(text), LabelHeight)
            AssignLane(probe, occupied)
        Next
        Return Math.Max(1, occupied.Count)
    End Function

    Private Shared Function MeasureNumberWidth(text As String) As Single
        Return Math.Max(12.0!, 5.7! * text.Length + 7.0!)
    End Function

    Private Shared Function FormatColumnId(value As Double) As String
        If Math.Abs(value - Math.Round(value)) <= 0.0000001R Then
            Return CInt(Math.Round(value)).ToString(CultureInfo.InvariantCulture)
        End If
        Return value.ToString("0.###", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function CountHorizontalLanes(geometry As PileLayoutGeometry, minX As Double,
                                                 scale As Double, xOffset As Double,
                                                 canvasWidth As Integer) As Integer
        Dim stations As Double() = geometry.HorizontalStationValues()
        Dim segments As Double() = geometry.HorizontalSegmentsMillimetres
        Dim occupied As New List(Of List(Of RectangleF))()
        For index As Integer = 0 To Math.Min(segments.Length, stations.Length - 1) - 1
            Dim text As String = FormatMillimetres(segments(index))
            Dim labelWidth As Single = MeasureLabelWidth(text)
            Dim centre As Single = CSng(xOffset + ((stations(index) + stations(index + 1)) / 2.0R - minX) * scale)
            Dim x As Single = Math.Max(2.0!, Math.Min(canvasWidth - labelWidth - 2.0!, centre - labelWidth / 2.0!))
            Dim probe As New RectangleF(x, 0.0!, labelWidth, LabelHeight)
            AssignLane(probe, occupied)
        Next
        Return Math.Max(1, occupied.Count)
    End Function

    Private Shared Sub AddHorizontalAnnotations(result As List(Of DimensionAnnotation),
                                                geometry As PileLayoutGeometry, minX As Double,
                                                scale As Double, xOffset As Double,
                                                planBottom As Single, canvasWidth As Integer,
                                                canvasHeight As Integer)
        Dim stations As Double() = geometry.HorizontalStationValues()
        Dim segments As Double() = geometry.HorizontalSegmentsMillimetres
        Dim occupied As New List(Of List(Of RectangleF))()
        For index As Integer = 0 To Math.Min(segments.Length, stations.Length - 1) - 1
            Dim startX As Single = CSng(xOffset + (stations(index) - minX) * scale)
            Dim endX As Single = CSng(xOffset + (stations(index + 1) - minX) * scale)
            Dim text As String = FormatMillimetres(segments(index))
            Dim labelWidth As Single = MeasureLabelWidth(text)
            Dim centre As Single = (startX + endX) / 2.0!
            Dim x As Single = Math.Max(2.0!, Math.Min(canvasWidth - labelWidth - 2.0!, centre - labelWidth / 2.0!))
            Dim probe As New RectangleF(x, 0.0!, labelWidth, LabelHeight)
            Dim lane As Integer = AssignLane(probe, occupied)
            Dim y As Single = Math.Min(canvasHeight - LabelHeight - 2.0!,
                                       planBottom + 7.0! + lane * (LabelHeight + LabelGap))
            Dim bounds As New RectangleF(x, y, labelWidth, LabelHeight)
            result.Add(New DimensionAnnotation(text, bounds,
                                               New PointF(startX, planBottom + 3.0!),
                                               New PointF(endX, planBottom + 3.0!), True))
        Next
    End Sub

    Private Shared Sub AddVerticalAnnotations(result As List(Of DimensionAnnotation),
                                              geometry As PileLayoutGeometry,
                                              minY As Double, maxY As Double,
                                              scale As Double, xOffset As Double,
                                              planOffsetY As Double, canvasWidth As Integer,
                                              canvasHeight As Integer)
        Dim stations As Double() = geometry.VerticalStationValues()
        Dim segments As Double() = geometry.VerticalSegmentsMillimetres
        Dim occupied As New List(Of List(Of RectangleF))()
        For index As Integer = 0 To Math.Min(segments.Length, stations.Length - 1) - 1
            Dim startY As Single = CSng(planOffsetY + (maxY - stations(index)) * scale)
            Dim endY As Single = CSng(planOffsetY + (maxY - stations(index + 1)) * scale)
            Dim text As String = FormatMillimetres(segments(index))
            Dim labelWidth As Single = MeasureLabelWidth(text)
            Dim centre As Single = (startY + endY) / 2.0!
            Dim probe As New RectangleF(0.0!, Math.Max(2.0!, Math.Min(canvasHeight - LabelHeight - 2.0!,
                                                                      centre - LabelHeight / 2.0!)),
                                        labelWidth, LabelHeight)
            Dim lane As Integer = AssignVerticalLane(probe, occupied)
            Dim x As Single = CSng(Math.Max(2.0R, xOffset - 8.0R - labelWidth -
                                                     lane * (labelWidth + LabelGap)))
            Dim bounds As New RectangleF(Math.Min(canvasWidth - labelWidth - 2.0!, x),
                                         probe.Y, labelWidth, LabelHeight)
            result.Add(New DimensionAnnotation(text, bounds,
                                               New PointF(CSng(xOffset - 4.0R), startY),
                                               New PointF(CSng(xOffset - 4.0R), endY), False))
        Next
    End Sub

    Private Shared Function AssignLane(probe As RectangleF,
                                       occupied As List(Of List(Of RectangleF))) As Integer
        For lane As Integer = 0 To occupied.Count - 1
            If Not occupied(lane).Any(Function(item) item.IntersectsWith(
                                          RectangleF.Inflate(probe, LabelGap, 0.0!))) Then
                occupied(lane).Add(probe)
                Return lane
            End If
        Next
        occupied.Add(New List(Of RectangleF) From {probe})
        Return occupied.Count - 1
    End Function

    Private Shared Function AssignVerticalLane(probe As RectangleF,
                                               occupied As List(Of List(Of RectangleF))) As Integer
        For lane As Integer = 0 To occupied.Count - 1
            If Not occupied(lane).Any(Function(item) item.IntersectsWith(
                                          RectangleF.Inflate(probe, 0.0!, LabelGap))) Then
                occupied(lane).Add(probe)
                Return lane
            End If
        Next
        occupied.Add(New List(Of RectangleF) From {probe})
        Return occupied.Count - 1
    End Function

    Private Shared Function MeasureLabelWidth(text As String) As Single
        Return Math.Max(19.0!, 5.7! * text.Length + 7.0!)
    End Function

    Private Shared Function FormatMillimetres(value As Double) As String
        Return value.ToString("0.###", CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub DrawBodyEdges(graphics As Graphics, pen As Pen, edges As Double(,),
                                     layout As PileLayoutResult)
        For row As Integer = 0 To edges.GetLength(0) - 1
            graphics.DrawLine(pen,
                MapFront(edges(row, 0), edges(row, 2), layout.XMinimum, layout.ZMinimum,
                         layout.PixelsPerMetre, layout.XOffset, layout.FrontOffsetY,
                         layout.HasPileBreak, layout.BreakTopZ, layout.BreakBottomZ),
                MapFront(edges(row, 3), edges(row, 5), layout.XMinimum, layout.ZMinimum,
                         layout.PixelsPerMetre, layout.XOffset, layout.FrontOffsetY,
                         layout.HasPileBreak, layout.BreakTopZ, layout.BreakBottomZ))
        Next
    End Sub

    Private Shared Sub DrawFrontGrillage(graphics As Graphics, pen As Pen, corners As PointF())
        For index As Integer = 0 To corners.Length - 1
            graphics.DrawLine(pen, corners(index), corners((index + 1) Mod corners.Length))
        Next
    End Sub

    Private Shared Sub DrawPileSymbol(graphics As Graphics, centre As PointF,
                                     diameter As Double, width As Double, scale As Double,
                                     outline As Pen, axisPen As Pen, fill As Brush)
        Dim actualSize As Double = If(width > 0.0R, width, diameter)
        Dim pixels As Single = CSng(Math.Max(2.0R, actualSize * scale))
        Dim bounds As New RectangleF(centre.X - pixels / 2.0!, centre.Y - pixels / 2.0!, pixels, pixels)
        If width > 0.0R Then
            graphics.FillRectangle(fill, bounds)
            graphics.DrawRectangle(outline, bounds.X, bounds.Y, bounds.Width, bounds.Height)
        Else
            graphics.FillEllipse(fill, bounds)
            graphics.DrawEllipse(outline, bounds)
        End If
        Dim axisLength As Single = Math.Max(2.5!, Math.Min(6.0!, pixels * 0.6!))
        graphics.DrawLine(axisPen, centre.X - axisLength, centre.Y, centre.X + axisLength, centre.Y)
        graphics.DrawLine(axisPen, centre.X, centre.Y - axisLength, centre.X, centre.Y + axisLength)
    End Sub

    Private Shared Sub DrawFrontPile(graphics As Graphics, piles As Double(,), row As Integer,
                                    layout As PileLayoutResult, outline As Pen, fill As Brush)
        Dim head As PointF = layout.FrontPileHeads(row)
        Dim toe As PointF = layout.FrontPileToes(row)
        If Not layout.HasPileBreak Then
            DrawFrontPileSegment(graphics, head, toe, piles(row, 6), piles(row, 7),
                                 layout.PixelsPerMetre, outline, fill)
            Return
        End If

        Dim upperIntersection As PointF = MapPileIntersection(piles, row, layout.BreakTopZ, layout)
        Dim lowerIntersection As PointF = MapPileIntersection(piles, row, layout.BreakBottomZ, layout)
        DrawFrontPileSegment(graphics, head, upperIntersection, piles(row, 6), piles(row, 7),
                             layout.PixelsPerMetre, outline, fill)
        DrawFrontPileSegment(graphics, lowerIntersection, toe, piles(row, 6), piles(row, 7),
                             layout.PixelsPerMetre, outline, fill)
        DrawPileBreakMark(graphics, outline, upperIntersection)
        DrawPileBreakMark(graphics, outline, lowerIntersection)
    End Sub

    Private Shared Function MapPileIntersection(piles As Double(,), row As Integer,
                                                z As Double, layout As PileLayoutResult) As PointF
        Dim dz As Double = piles(row, 5) - piles(row, 2)
        Dim factor As Double = If(Math.Abs(dz) <= Double.Epsilon, 0.0R,
                                  (z - piles(row, 2)) / dz)
        Dim x As Double = piles(row, 0) + (piles(row, 3) - piles(row, 0)) * factor
        Return MapFront(x, z, layout.XMinimum, layout.ZMinimum, layout.PixelsPerMetre,
                        layout.XOffset, layout.FrontOffsetY, layout.HasPileBreak,
                        layout.BreakTopZ, layout.BreakBottomZ)
    End Function

    Private Shared Sub DrawPileBreakMark(graphics As Graphics, pen As Pen, centre As PointF)
        graphics.DrawLines(pen, New PointF() {
            New PointF(centre.X - 4.0!, centre.Y + 1.5!),
            New PointF(centre.X - 1.5!, centre.Y - 1.5!),
            New PointF(centre.X + 1.5!, centre.Y + 1.5!),
            New PointF(centre.X + 4.0!, centre.Y - 1.5!)})
    End Sub

    Private Shared Sub DrawFrontPileSegment(graphics As Graphics, head As PointF, toe As PointF,
                                            diameter As Double, width As Double, scale As Double,
                                            outline As Pen, fill As Brush)
        Dim actualSize As Double = If(width > 0.0R, width, diameter)
        Dim pixels As Single = CSng(Math.Max(1.5R, actualSize * scale))
        Dim dx As Single = toe.X - head.X
        Dim dy As Single = toe.Y - head.Y
        Dim length As Single = CSng(Math.Sqrt(dx * dx + dy * dy))
        If length <= Single.Epsilon Then
            graphics.DrawEllipse(outline, head.X - pixels / 2.0!, head.Y - pixels / 2.0!, pixels, pixels)
            Return
        End If
        Dim offsetX As Single = -dy / length * pixels / 2.0!
        Dim offsetY As Single = dx / length * pixels / 2.0!
        Dim contour As PointF() = {
            New PointF(head.X + offsetX, head.Y + offsetY),
            New PointF(toe.X + offsetX, toe.Y + offsetY),
            New PointF(toe.X - offsetX, toe.Y - offsetY),
            New PointF(head.X - offsetX, head.Y - offsetY)
        }
        graphics.FillPolygon(fill, contour)
        graphics.DrawPolygon(outline, contour)
    End Sub

    Private Shared Sub DrawPileNumbers(graphics As Graphics, layout As PileLayoutResult,
                                       font As Font, brush As Brush)
        Using format As New StringFormat() With {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center
        }
            Dim numbers As String() = layout.FrontPileNumbers
            Dim bounds As RectangleF() = layout.FrontPileNumberBounds
            For index As Integer = 0 To Math.Min(numbers.Length, bounds.Length) - 1
                graphics.DrawString(numbers(index), font, brush, bounds(index), format)
            Next
        End Using
    End Sub

    Private Shared Sub DrawDimensions(graphics As Graphics, layout As PileLayoutResult,
                                      pen As Pen, font As Font, brush As Brush)
        Using format As New StringFormat() With {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center,
            .Trimming = StringTrimming.EllipsisCharacter
        }
            For Each annotation As DimensionAnnotation In layout.Annotations
                If annotation.Horizontal Then
                    graphics.DrawLine(pen, annotation.StartPoint, annotation.EndPoint)
                    graphics.DrawLine(pen, annotation.StartPoint.X, annotation.StartPoint.Y - 3.0!,
                                      annotation.StartPoint.X, annotation.StartPoint.Y + 3.0!)
                    graphics.DrawLine(pen, annotation.EndPoint.X, annotation.EndPoint.Y - 3.0!,
                                      annotation.EndPoint.X, annotation.EndPoint.Y + 3.0!)
                Else
                    graphics.DrawLine(pen, annotation.StartPoint, annotation.EndPoint)
                    graphics.DrawLine(pen, annotation.StartPoint.X - 3.0!, annotation.StartPoint.Y,
                                      annotation.StartPoint.X + 3.0!, annotation.StartPoint.Y)
                    graphics.DrawLine(pen, annotation.EndPoint.X - 3.0!, annotation.EndPoint.Y,
                                      annotation.EndPoint.X + 3.0!, annotation.EndPoint.Y)
                End If
                graphics.DrawString(annotation.Text, font, brush, annotation.Bounds, format)
            Next
        End Using
    End Sub
End Class
