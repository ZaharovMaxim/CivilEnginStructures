Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Linq

Friend NotInheritable Class PilePlanGeometry
    Private Const StationTolerance As Double = 0.0005R
    Private Const MissingGrillageStatus As String = "Ростверк отсутствует — размерные цепи до граней недоступны"
    Private Const MissingPilesStatus As String = "Сваи отсутствуют — размерные цепи не построены"

    Friend NotInheritable Class PileSymbol
        Friend ReadOnly X As Double
        Friend ReadOnly Y As Double
        Friend ReadOnly Diameter As Double
        Friend ReadOnly Width As Double

        Friend Sub New(x As Double, y As Double, diameter As Double, width As Double)
            Me.X = x
            Me.Y = y
            Me.Diameter = diameter
            Me.Width = width
        End Sub
    End Class

    Friend NotInheritable Class PlanPoint
        Friend ReadOnly X As Double
        Friend ReadOnly Y As Double

        Friend Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub
    End Class

    Private ReadOnly horizontalSegments As Double()
    Private ReadOnly verticalSegments As Double()
    Private ReadOnly horizontalStationValues As Double()
    Private ReadOnly verticalStationValues As Double()
    Private ReadOnly localCorners As PlanPoint()
    Private ReadOnly localPiles As PileSymbol()

    Private Sub New(corners As PlanPoint(), piles As PileSymbol(), horizontal As Double(),
                    vertical As Double(), horizontalAxes As Double(), verticalAxes As Double(),
                    status As String, hasGrillage As Boolean)
        localCorners = corners
        localPiles = piles
        horizontalSegments = horizontal
        verticalSegments = vertical
        horizontalStationValues = horizontalAxes
        verticalStationValues = verticalAxes
        StatusText = status
        Me.HasGrillage = hasGrillage
    End Sub

    Friend Shared Function Create(corners As Double(,), piles As Double(,)) As PilePlanGeometry
        Dim sourcePiles As List(Of Double()) = CopyValidPiles(piles)
        If Not HasValidCorners(corners) Then
            Return New PilePlanGeometry(Array.Empty(Of PlanPoint)(),
                                        CreateUnprojectedPileSymbols(sourcePiles),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        MissingGrillageStatus, False)
        End If

        Dim originX As Double = corners(0, 0)
        Dim originY As Double = corners(0, 1)
        Dim edgeX As Double = corners(1, 0) - originX
        Dim edgeY As Double = corners(1, 1) - originY
        Dim edgeLength As Double = Math.Sqrt(edgeX * edgeX + edgeY * edgeY)
        If Not IsFinite(edgeLength) OrElse edgeLength <= Double.Epsilon Then
            Return New PilePlanGeometry(Array.Empty(Of PlanPoint)(),
                                        CreateUnprojectedPileSymbols(sourcePiles),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        MissingGrillageStatus, False)
        End If

        Dim ux As Double = edgeX / edgeLength
        Dim uy As Double = edgeY / edgeLength
        Dim vx As Double = -uy
        Dim vy As Double = ux
        Dim towardFourth As Double = (corners(3, 0) - originX) * vx + (corners(3, 1) - originY) * vy
        If towardFourth < 0.0R Then
            vx = -vx
            vy = -vy
        End If

        Dim projectedCorners(3) As PlanPoint
        For index As Integer = 0 To 3
            projectedCorners(index) = Project(corners(index, 0), corners(index, 1),
                                              originX, originY, ux, uy, vx, vy)
        Next
        If Math.Abs(PolygonArea(projectedCorners)) <= 0.000000001R Then
            Return New PilePlanGeometry(Array.Empty(Of PlanPoint)(),
                                        CreateUnprojectedPileSymbols(sourcePiles),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        Array.Empty(Of Double)(), Array.Empty(Of Double)(),
                                        MissingGrillageStatus, False)
        End If

        Dim projectedPiles As PileSymbol() = Array.Empty(Of PileSymbol)()
        If sourcePiles.Count > 0 Then ReDim projectedPiles(sourcePiles.Count - 1)
        For index As Integer = 0 To sourcePiles.Count - 1
            Dim source As Double() = sourcePiles(index)
            Dim centre As PlanPoint = Project(source(0), source(1), originX, originY, ux, uy, vx, vy)
            projectedPiles(index) = New PileSymbol(centre.X, centre.Y, source(2), source(3))
        Next

        Dim status As String = If(projectedPiles.Length = 0, MissingPilesStatus, String.Empty)
        Dim horizontal As Double() = Array.Empty(Of Double)()
        Dim vertical As Double() = Array.Empty(Of Double)()
        Dim horizontalAxes As Double() = Array.Empty(Of Double)()
        Dim verticalAxes As Double() = Array.Empty(Of Double)()
        If projectedPiles.Length > 0 Then
            horizontalAxes = CreateDimensionStations(projectedCorners.Select(Function(point) point.X),
                                                      projectedPiles.Select(Function(pile) pile.X))
            verticalAxes = CreateDimensionStations(projectedCorners.Select(Function(point) point.Y),
                                                    projectedPiles.Select(Function(pile) pile.Y))
            horizontal = CreateDimensionSegments(horizontalAxes)
            vertical = CreateDimensionSegments(verticalAxes)
        End If
        Return New PilePlanGeometry(projectedCorners, projectedPiles, horizontal, vertical,
                                    horizontalAxes, verticalAxes,
                                    status, True)
    End Function

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

    Friend ReadOnly Property PileCount As Integer
        Get
            Return localPiles.Length
        End Get
    End Property

    Friend ReadOnly Property StatusText As String

    Friend ReadOnly Property HasGrillage As Boolean

    Friend Function Corners() As PlanPoint()
        Return DirectCast(localCorners.Clone(), PlanPoint())
    End Function

    Friend Function Piles() As PileSymbol()
        Return DirectCast(localPiles.Clone(), PileSymbol())
    End Function

    Friend Function HorizontalStations() As Double()
        Return DirectCast(horizontalStationValues.Clone(), Double())
    End Function

    Friend Function VerticalStations() As Double()
        Return DirectCast(verticalStationValues.Clone(), Double())
    End Function

    Private Shared Function HasValidCorners(corners As Double(,)) As Boolean
        If corners Is Nothing OrElse corners.Rank <> 2 OrElse
           corners.GetLength(0) < 4 OrElse corners.GetLength(1) < 2 Then Return False
        For row As Integer = 0 To 3
            If Not IsFinite(corners(row, 0)) OrElse Not IsFinite(corners(row, 1)) Then Return False
        Next
        Return True
    End Function

    Private Shared Function CopyValidPiles(piles As Double(,)) As List(Of Double())
        Dim result As New List(Of Double())()
        If piles Is Nothing OrElse piles.Rank <> 2 OrElse piles.GetLength(1) < 2 Then Return result
        For row As Integer = 0 To piles.GetLength(0) - 1
            Dim x As Double = piles(row, 0)
            Dim y As Double = piles(row, 1)
            If Not IsFinite(x) OrElse Not IsFinite(y) Then Continue For
            Dim diameter As Double = 0.0R
            Dim width As Double = 0.0R
            If piles.GetLength(1) > 2 AndAlso IsFinite(piles(row, 2)) Then diameter = Math.Max(0.0R, piles(row, 2))
            If piles.GetLength(1) > 3 AndAlso IsFinite(piles(row, 3)) Then width = Math.Max(0.0R, piles(row, 3))
            result.Add(New Double() {x, y, diameter, width})
        Next
        Return result
    End Function

    Private Shared Function CreateUnprojectedPileSymbols(source As List(Of Double())) As PileSymbol()
        Dim result As PileSymbol() = Array.Empty(Of PileSymbol)()
        If source.Count > 0 Then ReDim result(source.Count - 1)
        For index As Integer = 0 To source.Count - 1
            result(index) = New PileSymbol(source(index)(0), source(index)(1),
                                           source(index)(2), source(index)(3))
        Next
        Return result
    End Function

    Private Shared Function Project(x As Double, y As Double, originX As Double, originY As Double,
                                    ux As Double, uy As Double, vx As Double, vy As Double) As PlanPoint
        Dim dx As Double = x - originX
        Dim dy As Double = y - originY
        Return New PlanPoint(dx * ux + dy * uy, dx * vx + dy * vy)
    End Function

    Private Shared Function PolygonArea(points As PlanPoint()) As Double
        Dim result As Double = 0.0R
        For index As Integer = 0 To points.Length - 1
            Dim nextIndex As Integer = (index + 1) Mod points.Length
            result += points(index).X * points(nextIndex).Y - points(nextIndex).X * points(index).Y
        Next
        Return result / 2.0R
    End Function

    Private Shared Function CreateDimensionStations(edges As IEnumerable(Of Double),
                                                     centres As IEnumerable(Of Double)) As Double()
        Dim edgeValues As Double() = edges.Where(AddressOf IsFinite).ToArray()
        If edgeValues.Length = 0 Then Return Array.Empty(Of Double)()
        Dim stations As New List(Of Double)() From {edgeValues.Min(), edgeValues.Max()}
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

Friend NotInheritable Class PilePlanPreview
    Private Sub New()
    End Sub

    Friend Shared Function Render(geometry As PilePlanGeometry, requestedSize As Size) As Bitmap
        Dim width As Integer = If(requestedSize.Width > 0, requestedSize.Width, 320)
        Dim height As Integer = If(requestedSize.Height > 0, requestedSize.Height, 180)
        Dim bitmap As New Bitmap(Math.Max(1, width), Math.Max(1, height))
        Using graphics As Graphics = Graphics.FromImage(bitmap),
              titleFont As New Font("Segoe UI", If(width < 180, 7.0!, 10.0!), FontStyle.Bold, GraphicsUnit.Point),
              textFont As New Font("Segoe UI", If(width < 180, 6.0!, 8.0!), FontStyle.Regular, GraphicsUnit.Point),
              contourPen As New Pen(Color.FromArgb(55, 65, 81), 1.0!),
              pilePen As New Pen(Color.FromArgb(37, 99, 235), 1.2!),
              axisPen As New Pen(Color.FromArgb(150, 160, 176), 1.0!),
              dimensionPen As New Pen(Color.FromArgb(75, 85, 99), 1.0!),
              grillageBrush As New SolidBrush(Color.FromArgb(239, 242, 247)),
              pileBrush As New SolidBrush(Color.White),
              textBrush As New SolidBrush(Color.FromArgb(32, 50, 77))
            graphics.SmoothingMode = SmoothingMode.AntiAlias
            graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
            graphics.Clear(Color.White)
            graphics.DrawString("План свай · размеры по осям, мм", titleFont, textBrush,
                                New RectangleF(6.0!, 4.0!, Math.Max(1, width - 12), Math.Max(1, height - 8)))

            If geometry Is Nothing OrElse Not geometry.HasGrillage Then
                Dim status As String = If(geometry Is Nothing,
                                          "Ростверк отсутствует — размерные цепи до граней недоступны",
                                          geometry.StatusText)
                DrawStatus(graphics, status, textFont, textBrush, width, height)
                Return bitmap
            End If

            Dim corners As PilePlanGeometry.PlanPoint() = geometry.Corners()
            Dim piles As PilePlanGeometry.PileSymbol() = geometry.Piles()
            Dim leftMargin As Single = CSng(Math.Min(70.0R, Math.Max(14.0R, width * 0.13R)))
            Dim rightMargin As Single = CSng(Math.Min(22.0R, Math.Max(6.0R, width * 0.035R)))
            Dim topMargin As Single = CSng(Math.Min(78.0R, Math.Max(60.0R, height * 0.28R)))
            Dim bottomMargin As Single = CSng(Math.Min(28.0R, Math.Max(8.0R, height * 0.06R)))
            Dim plotWidth As Single = Math.Max(4.0!, width - leftMargin - rightMargin)
            Dim plotHeight As Single = Math.Max(4.0!, height - topMargin - bottomMargin)

            Dim minX As Double = corners.Min(Function(point) CDbl(point.X))
            Dim maxX As Double = corners.Max(Function(point) CDbl(point.X))
            Dim minY As Double = corners.Min(Function(point) CDbl(point.Y))
            Dim maxY As Double = corners.Max(Function(point) CDbl(point.Y))
            For Each pile As PilePlanGeometry.PileSymbol In piles
                Dim size As Double = If(pile.Width > 0.0R, pile.Width, pile.Diameter)
                minX = Math.Min(minX, pile.X - size / 2.0R)
                maxX = Math.Max(maxX, pile.X + size / 2.0R)
                minY = Math.Min(minY, pile.Y - size / 2.0R)
                maxY = Math.Max(maxY, pile.Y + size / 2.0R)
            Next
            Dim spanX As Double = Math.Max(0.001R, maxX - minX)
            Dim spanY As Double = Math.Max(0.001R, maxY - minY)
            Dim scale As Double = Math.Min(plotWidth / spanX, plotHeight / spanY)
            Dim diagramWidth As Double = spanX * scale
            Dim diagramHeight As Double = spanY * scale
            Dim offsetX As Double = leftMargin + (plotWidth - diagramWidth) / 2.0R
            Dim offsetY As Double = topMargin + (plotHeight - diagramHeight) / 2.0R

            Dim mappedCorners As PointF() = corners.Select(
                Function(point) MapPoint(point.X, point.Y, minX, maxY, scale, offsetX, offsetY)).ToArray()
            graphics.FillPolygon(grillageBrush, mappedCorners)
            graphics.DrawPolygon(contourPen, mappedCorners)

            DrawDimensions(graphics, geometry, corners, minX, maxY, scale, offsetX, offsetY,
                           dimensionPen, textFont, textBrush)
            For Each pile As PilePlanGeometry.PileSymbol In piles
                DrawPile(graphics, pile, minX, maxY, scale, offsetX, offsetY,
                         pilePen, axisPen, pileBrush)
            Next

            If Not String.IsNullOrEmpty(geometry.StatusText) Then
                DrawStatus(graphics, geometry.StatusText, textFont, textBrush, width, height)
            End If
        End Using
        Return bitmap
    End Function

    Private Shared Sub DrawDimensions(graphics As Graphics, geometry As PilePlanGeometry,
                                      corners As PilePlanGeometry.PlanPoint(), minX As Double, maxY As Double,
                                      scale As Double, offsetX As Double, offsetY As Double,
                                      pen As Pen, font As Font, brush As Brush)
        Dim xSegments As Double() = geometry.HorizontalSegmentsMillimetres
        Dim ySegments As Double() = geometry.VerticalSegmentsMillimetres
        Dim xStations As Double() = geometry.HorizontalStations()
        Dim yStations As Double() = geometry.VerticalStations()

        If xSegments.Length > 0 AndAlso xStations.Length = xSegments.Length + 1 Then
            Dim chainY As Single = CSng(offsetY - 14.0R)
            Dim station As Double = xStations(0)
            Dim startX As Single = MapX(station, minX, scale, offsetX)
            For index As Integer = 0 To xSegments.Length - 1
                Dim segment As Double = xSegments(index)
                Dim nextStation As Double = xStations(index + 1)
                Dim nextX As Single = MapX(nextStation, minX, scale, offsetX)
                DrawHorizontalDimension(graphics, pen, font, brush, startX, nextX, chainY, segment)
                station = nextStation
                startX = nextX
            Next
        End If

        If ySegments.Length > 0 AndAlso yStations.Length = ySegments.Length + 1 Then
            Dim chainX As Single = CSng(offsetX - 14.0R)
            Dim station As Double = yStations(0)
            Dim startY As Single = MapY(station, maxY, scale, offsetY)
            For index As Integer = 0 To ySegments.Length - 1
                Dim segment As Double = ySegments(index)
                Dim nextStation As Double = yStations(index + 1)
                Dim nextY As Single = MapY(nextStation, maxY, scale, offsetY)
                DrawVerticalDimension(graphics, pen, font, brush, chainX, startY, nextY, segment)
                station = nextStation
                startY = nextY
            Next
        End If
    End Sub

    Private Shared Sub DrawHorizontalDimension(graphics As Graphics, pen As Pen, font As Font,
                                               brush As Brush, startX As Single, endX As Single,
                                               y As Single, millimetres As Double)
        graphics.DrawLine(pen, startX, y, endX, y)
        graphics.DrawLine(pen, startX, y - 3.0!, startX, y + 3.0!)
        graphics.DrawLine(pen, endX, y - 3.0!, endX, y + 3.0!)
        Dim label As String = FormatMillimetres(millimetres)
        Dim size As SizeF = graphics.MeasureString(label, font)
        graphics.DrawString(label, font, brush, (startX + endX - size.Width) / 2.0!, y - size.Height - 1.0!)
    End Sub

    Private Shared Sub DrawVerticalDimension(graphics As Graphics, pen As Pen, font As Font,
                                             brush As Brush, x As Single, startY As Single,
                                             endY As Single, millimetres As Double)
        graphics.DrawLine(pen, x, startY, x, endY)
        graphics.DrawLine(pen, x - 3.0!, startY, x + 3.0!, startY)
        graphics.DrawLine(pen, x - 3.0!, endY, x + 3.0!, endY)
        Dim label As String = FormatMillimetres(millimetres)
        Dim state As GraphicsState = graphics.Save()
        graphics.TranslateTransform(x - 3.0!, (startY + endY) / 2.0!)
        graphics.RotateTransform(-90.0!)
        Dim size As SizeF = graphics.MeasureString(label, font)
        graphics.DrawString(label, font, brush, -size.Width / 2.0!, -size.Height)
        graphics.Restore(state)
    End Sub

    Private Shared Sub DrawPile(graphics As Graphics, pile As PilePlanGeometry.PileSymbol,
                                minX As Double, maxY As Double, scale As Double,
                                offsetX As Double, offsetY As Double, outline As Pen,
                                axisPen As Pen, fill As Brush)
        Dim centre As PointF = MapPoint(pile.X, pile.Y, minX, maxY, scale, offsetX, offsetY)
        Dim actualSize As Double = If(pile.Width > 0.0R, pile.Width, pile.Diameter)
        Dim pixels As Single = CSng(Math.Max(0.0R, actualSize * scale))
        If pixels >= 1.0! Then
            Dim bounds As New RectangleF(centre.X - pixels / 2.0!, centre.Y - pixels / 2.0!, pixels, pixels)
            If pile.Width > 0.0R Then
                graphics.FillRectangle(fill, bounds)
                graphics.DrawRectangle(outline, bounds.X, bounds.Y, bounds.Width, bounds.Height)
            Else
                graphics.FillEllipse(fill, bounds)
                graphics.DrawEllipse(outline, bounds)
            End If
        End If
        Dim axisLength As Single = Math.Max(3.0!, Math.Min(8.0!, If(pixels > 0.0!, pixels * 0.65!, 5.0!)))
        graphics.DrawLine(axisPen, centre.X - axisLength, centre.Y, centre.X + axisLength, centre.Y)
        graphics.DrawLine(axisPen, centre.X, centre.Y - axisLength, centre.X, centre.Y + axisLength)
    End Sub

    Private Shared Sub DrawStatus(graphics As Graphics, status As String, font As Font,
                                  brush As Brush, width As Integer, height As Integer)
        If String.IsNullOrEmpty(status) Then Return
        Using format As New StringFormat() With {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center,
            .Trimming = StringTrimming.EllipsisWord
        }
            graphics.DrawString(status, font, brush,
                                New RectangleF(8.0!, 30.0!, Math.Max(1, width - 16), Math.Max(1, height - 38)),
                                format)
        End Using
    End Sub

    Private Shared Function MapPoint(x As Double, y As Double, minX As Double, maxY As Double,
                                    scale As Double, offsetX As Double, offsetY As Double) As PointF
        Return New PointF(MapX(x, minX, scale, offsetX), MapY(y, maxY, scale, offsetY))
    End Function

    Private Shared Function MapX(x As Double, minX As Double, scale As Double, offsetX As Double) As Single
        Return CSng(offsetX + (x - minX) * scale)
    End Function

    Private Shared Function MapY(y As Double, maxY As Double, scale As Double, offsetY As Double) As Single
        Return CSng(offsetY + (maxY - y) * scale)
    End Function

    Private Shared Function FormatMillimetres(value As Double) As String
        Return Math.Round(value, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture)
    End Function
End Class
