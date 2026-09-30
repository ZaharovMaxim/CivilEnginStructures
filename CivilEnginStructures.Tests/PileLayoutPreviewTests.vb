Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class PileLayoutPreviewTests
        <Test>
        Public Sub GeometryNormalizesWorldCoordinatesWithoutMutatingInputs()
            Dim corners As Double(,) = RectangleCorners(10.0R, 20.0R, 6.0R, 4.0R, 2.0R)
            Dim piles As Double(,) = {
                {11.0R, 21.0R, 3.0R, 11.2R, 21.3R, 9.0R, 0.6R, 0.0R},
                {15.0R, 23.0R, 3.2R, 14.7R, 22.8R, 8.5R, 0.0R, 0.5R}
            }
            Dim edges As Double(,) = {
                {10.0R, 22.0R, 1.5R, 16.0R, 22.0R, 1.5R}
            }
            Dim originalCorners As Double(,) = DirectCast(corners.Clone(), Double(,))
            Dim originalPiles As Double(,) = DirectCast(piles.Clone(), Double(,))
            Dim originalEdges As Double(,) = DirectCast(edges.Clone(), Double(,))

            Dim geometry As Object = CreateGeometry(corners, piles, edges)

            Assert.Multiple(
                Sub()
                    Assert.That(BooleanProperty(geometry, "IsReady"), [Is].True)
                    Assert.That(IntegerProperty(geometry, "PileCount"), [Is].EqualTo(2))
                    AssertMatrix(MatrixProperty(geometry, "Corners"),
                                 New Double(,) {{0.0R, 0.0R, 2.0R},
                                                 {6.0R, 0.0R, 2.0R},
                                                 {6.0R, 4.0R, 2.0R},
                                                 {0.0R, 4.0R, 2.0R}})
                    AssertMatrix(MatrixProperty(geometry, "Piles"),
                                 New Double(,) {{1.0R, 1.0R, 3.0R, 1.2R, 1.3R, 9.0R, 0.6R, 0.0R},
                                                 {5.0R, 3.0R, 3.2R, 4.7R, 2.8R, 8.5R, 0.0R, 0.5R}})
                    AssertMatrix(MatrixProperty(geometry, "Edges"),
                                 New Double(,) {{0.0R, 2.0R, 1.5R, 6.0R, 2.0R, 1.5R}})
                    AssertMatrix(corners, originalCorners)
                    AssertMatrix(piles, originalPiles)
                    AssertMatrix(edges, originalEdges)
                End Sub)
        End Sub

        <Test>
        Public Sub TranslationAndRotationLeaveNormalizedGeometryUnchanged()
            Dim corners As Double(,) = RectangleCorners(10.0R, 20.0R, 6.0R, 4.0R, 2.0R)
            Dim piles As Double(,) = {
                {11.0R, 21.0R, 3.0R, 11.2R, 21.3R, 9.0R, 0.6R, 0.0R},
                {15.0R, 23.0R, 3.2R, 14.7R, 22.8R, 8.5R, 0.0R, 0.5R}
            }
            Dim edges As Double(,) = {{10.0R, 22.0R, 1.5R, 16.0R, 22.0R, 1.5R}}
            Dim expected As Object = CreateGeometry(corners, piles, edges)

            Dim transformed As Object = CreateGeometry(
                TransformXY(corners, 37.0R, 123.4R, -57.8R, 0, 1),
                TransformXY(piles, 37.0R, 123.4R, -57.8R, 0, 1, 3, 4),
                TransformXY(edges, 37.0R, 123.4R, -57.8R, 0, 1, 3, 4))

            Assert.Multiple(
                Sub()
                    AssertMatrix(MatrixProperty(transformed, "Corners"), MatrixProperty(expected, "Corners"))
                    AssertMatrix(MatrixProperty(transformed, "Piles"), MatrixProperty(expected, "Piles"))
                    AssertMatrix(MatrixProperty(transformed, "Edges"), MatrixProperty(expected, "Edges"))
                    AssertDoubleArrays(DoubleArrayProperty(transformed, "HorizontalSegmentsMillimetres"),
                                       DoubleArrayProperty(expected, "HorizontalSegmentsMillimetres"), 0.000001R)
                    AssertDoubleArrays(DoubleArrayProperty(transformed, "VerticalSegmentsMillimetres"),
                                       DoubleArrayProperty(expected, "VerticalSegmentsMillimetres"), 0.000001R)
                End Sub)
        End Sub

        <Test>
        Public Sub DimensionChainsUseActualAxesAt999And1001Millimetres()
            Dim piles As Double(,) = {
                {0.999R, 0.5R, 1.0R, 0.999R, 0.5R, 6.0R, 0.5R, 0.0R},
                {1.001R, 0.5R, 1.0R, 1.001R, 0.5R, 6.0R, 0.5R, 0.0R}
            }

            Dim geometry As Object = CreateGeometry(RectangleCorners(0.0R, 0.0R, 2.0R, 1.0R, 0.0R), piles, Nothing)

            Assert.Multiple(
                Sub()
                    Assert.That(DoubleArrayProperty(geometry, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {999.0R, 2.0R, 999.0R}).Within(0.01R),
                                "The chain must use the two calculated axes, not an invented 1000 mm station.")
                    Assert.That(DoubleArrayProperty(geometry, "VerticalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {500.0R, 500.0R}).Within(0.01R))
                End Sub)
        End Sub

        <Test>
        Public Sub NominalStationsRemoveProjectionJitterWithoutMovingRenderedPiles()
            Const columns As Integer = 21
            Const rows As Integer = 4
            Dim localPiles As Double(,) = Nothing
            Dim axisData As Double(,) = Nothing
            CreateRegularPileGrid(columns, rows, 0.5R, 0.4R, 1.0R, 0.5R,
                                  4.5R, 15.5R, True, localPiles, axisData)
            Dim corners As Double(,) = TransformXY(
                RectangleCorners(0.0R, 0.0R, 21.0R, 2.4R, 4.0R),
                37.0R, 123.4R, -57.8R, 0, 1)
            Dim worldPiles As Double(,) = TransformXY(localPiles, 37.0R, 123.4R, -57.8R, 0, 1, 3, 4)

            Dim geometry As Object = CreateGeometry(
                corners, worldPiles, Nothing, axisData, New Double() {21.0R, 2.4R})
            Dim normalizedPiles As Double(,) = MatrixProperty(geometry, "Piles")

            Assert.Multiple(
                Sub()
                    Assert.That(BooleanProperty(geometry, "IsReady"), [Is].True)
                    Assert.That(IntegerProperty(geometry, "PileCount"), [Is].EqualTo(columns * rows))
                    Assert.That(DoubleArrayProperty(geometry, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(ExpectedRegularSegments(500.0R, 20, 1000.0R, 500.0R)).Within(0.01R))
                    Assert.That(DoubleArrayProperty(geometry, "VerticalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {400.0R, 500.0R, 500.0R, 500.0R, 500.0R}).Within(0.01R))
                    Assert.That(normalizedPiles(0, 0), [Is].EqualTo(localPiles(0, 0)).Within(0.0000001R),
                                "Nominal stations affect dimensions only; pile graphics keep calculated coordinates.")
                    Assert.That(normalizedPiles(0, 1), [Is].EqualTo(localPiles(0, 1)).Within(0.0000001R))
                    Assert.That(normalizedPiles(columns * rows - 1, 0),
                                [Is].EqualTo(localPiles(columns * rows - 1, 0)).Within(0.0000001R))
                End Sub)
        End Sub

        <Test>
        Public Sub NominalStationsPreserveLegitimate999And1001MillimetreOffsets()
            Dim piles As Double(,) = {
                {0.999R, 0.5R, 1.0R, 0.999R, 0.5R, 6.0R, 0.5R, 0.0R},
                {1.001R, 0.5R, 1.0R, 1.001R, 0.5R, 6.0R, 0.5R, 0.0R}
            }
            Dim axisData As Double(,) = {
                {1.0R, 1.0R, 0.999R, 0.5R},
                {1.0R, 2.0R, 1.001R, 0.5R}
            }

            Dim geometry As Object = CreateGeometry(
                RectangleCorners(0.0R, 0.0R, 2.0R, 1.0R, 0.0R),
                piles, Nothing, axisData, New Double() {2.0R, 1.0R})

            Assert.Multiple(
                Sub()
                    Assert.That(DoubleArrayProperty(geometry, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {999.0R, 2.0R, 999.0R}).Within(0.01R),
                                "A real unequal nominal step must not be rounded into an invented regular grid.")
                    Assert.That(DoubleArrayProperty(geometry, "VerticalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {500.0R, 500.0R}).Within(0.01R))
                End Sub)
        End Sub

        <Test>
        Public Sub InvalidNominalMetadataFallsBackToCalculatedAxesAndKeepsSourceRowAlignment()
            Dim corners As Double(,) = RectangleCorners(0.0R, 0.0R, 2.0R, 1.0R, 0.0R)
            Dim piles As Double(,) = {
                {0.9R, 0.5R, 1.0R, 0.9R, 0.5R, 6.0R, 0.5R, 0.0R},
                {1.1R, 0.5R, 1.0R, 1.1R, 0.5R, 6.0R, 0.5R, 0.0R}
            }
            Dim rawGeometry As Object = CreateGeometry(corners, piles, Nothing)
            Dim wrongShape As Object = CreateGeometry(
                corners, piles, Nothing, New Double(,) {{1.0R, 1.0R, 0.5R, 0.5R}},
                New Double() {2.0R, 1.0R})
            Dim mismatchedCoordinates As Object = CreateGeometry(
                corners, piles, Nothing,
                New Double(,) {{1.0R, 1.0R, 0.5R, 0.5R}, {1.0R, 2.0R, 1.5R, 0.5R}},
                New Double() {2.0R, 1.0R})

            Dim rowsWithDegenerate As Double(,) = {
                {0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.5R, 0.0R},
                {0.5R, 0.5R, 1.0R, 0.5R, 0.5R, 6.0R, 0.5R, 0.0R},
                {1.5R, 0.5R, 1.0R, 1.5R, 0.5R, 6.0R, 0.5R, 0.0R}
            }
            Dim alignedMetadata As Double(,) = {
                {99.0R, 99.0R, 1.75R, 0.25R},
                {1.0R, 1.0R, 0.5R, 0.5R},
                {1.0R, 2.0R, 1.5R, 0.5R}
            }
            Dim filtered As Object = CreateGeometry(
                corners, rowsWithDegenerate, Nothing, alignedMetadata, New Double() {2.0R, 1.0R})

            Assert.Multiple(
                Sub()
                    Assert.That(DoubleArrayProperty(wrongShape, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(DoubleArrayProperty(rawGeometry, "HorizontalSegmentsMillimetres")).Within(0.01R))
                    Assert.That(DoubleArrayProperty(mismatchedCoordinates, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(DoubleArrayProperty(rawGeometry, "HorizontalSegmentsMillimetres")).Within(0.01R),
                                "Metadata far from calculated axes must not invent a scheme dimension chain.")
                    Assert.That(IntegerProperty(filtered, "PileCount"), [Is].EqualTo(2))
                    Assert.That(DoubleArrayProperty(filtered, "HorizontalSegmentsMillimetres"),
                                [Is].EqualTo(New Double() {500.0R, 1000.0R, 500.0R}).Within(0.01R),
                                "Filtering a degenerate pile must filter the matching metadata row.")
                End Sub)
        End Sub

        <Test>
        Public Sub LongDensePileLayoutUsesExplicitBreakAndKeepsPlanReadable()
            Const columns As Integer = 21
            Const rows As Integer = 4
            Dim piles As Double(,) = Nothing
            Dim axisData As Double(,) = Nothing
            CreateRegularPileGrid(columns, rows, 0.5R, 0.4R, 1.0R, 0.5R,
                                  4.5R, 15.5R, False, piles, axisData)
            Dim edges As Double(,) = {
                {0.0R, 0.0R, 0.0R, 21.0R, 0.0R, 0.0R},
                {0.0R, 0.0R, 4.0R, 21.0R, 0.0R, 4.0R},
                {0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 4.0R},
                {21.0R, 0.0R, 0.0R, 21.0R, 0.0R, 4.0R}
            }
            Dim geometry As Object = CreateGeometry(
                RectangleCorners(0.0R, 0.0R, 21.0R, 2.4R, 4.0R),
                piles, edges, axisData, New Double() {21.0R, 2.4R})

            For Each canvas As Size In New Size() {New Size(600, 500), New Size(480, 300)}
                Dim layout As Object = CreateLayout(geometry, canvas)
                Dim scale As Double = DoubleProperty(layout, "PixelsPerMetre")
                Dim frontHeads As PointF() = PointArrayProperty(layout, "FrontPileHeads")
                Dim frontToes As PointF() = PointArrayProperty(layout, "FrontPileToes")
                Dim planHeads As PointF() = PointArrayProperty(layout, "PlanPileHeads")
                Dim planCorners As PointF() = PointArrayProperty(layout, "PlanGrillageCorners")
                Dim labels As RectangleF() = RectangleArrayProperty(layout, "DimensionLabelBounds")
                Dim pileNumbers As String() = StringArrayProperty(layout, "FrontPileNumbers")
                Dim pileNumberBounds As RectangleF() = RectangleArrayProperty(layout, "FrontPileNumberBounds")
                Dim breakTop As Double = DoubleProperty(layout, "BreakTopZ")
                Dim breakBottom As Double = DoubleProperty(layout, "BreakBottomZ")
                Dim planTop As Single = planCorners.Min(Function(point) point.Y)

                Assert.Multiple(
                    Sub()
                        Assert.That(BooleanProperty(layout, "HasPileBreak"), [Is].True)
                        Assert.That(breakTop, [Is].GreaterThan(4.5R))
                        Assert.That(breakBottom, [Is].LessThan(15.5R))
                        Assert.That(breakBottom, [Is].GreaterThan(breakTop))
                        Assert.That(Math.Abs(planCorners(1).X - planCorners(0).X),
                                    [Is].GreaterThanOrEqualTo(canvas.Width * 0.7F),
                                    "The 21 m plan must use most of the available width.")
                        Assert.That(Math.Abs(planHeads(1).X - planHeads(0).X),
                                    [Is].EqualTo(scale).Within(0.03R))
                        Assert.That(Math.Abs(planHeads(columns).Y - planHeads(0).Y),
                                    [Is].EqualTo(0.5R * scale).Within(0.03R))
                        Assert.That(frontHeads(0).X, [Is].EqualTo(planHeads(0).X).Within(0.03F))
                        Assert.That(frontToes(0).X, [Is].EqualTo(frontHeads(0).X).Within(0.03F))
                        Assert.That(Math.Abs(frontToes(0).Y - frontHeads(0).Y),
                                    [Is].LessThan(11.0R * scale),
                                    "The omitted pile interval must be visible through the explicit break state.")
                        Assert.That(pileNumbers,
                                    [Is].EqualTo(Enumerable.Range(1, columns).
                                                 Select(Function(value) value.ToString()).ToArray()),
                                    "Four pile rows share the same 21 numbered front-view columns.")
                        Assert.That(pileNumberBounds.Length, [Is].EqualTo(columns))
                        For column As Integer = 0 To columns - 1
                            Assert.That(pileNumberBounds(column).Left + pileNumberBounds(column).Width / 2.0F,
                                        [Is].EqualTo(frontHeads(column).X).Within(0.6F),
                                        "A pile number must be centred on its front-view column.")
                            Assert.That(pileNumberBounds(column).Top,
                                        [Is].GreaterThanOrEqualTo(frontToes(column).Y))
                            Assert.That(pileNumberBounds(column).Bottom, [Is].LessThanOrEqualTo(planTop))
                        Next
                        AssertRectanglesDoNotOverlap(pileNumberBounds, "Front pile number")
                        AssertRectanglesDoNotIntersect(pileNumberBounds, labels,
                                                       "Front pile numbers must not overlap dimension labels.")
                        AssertLabelsInsideAndSeparate(labels, canvas)
                    End Sub)
            Next

            Dim shortPile As Double(,) = {{1.0R, 0.5R, 1.0R, 1.0R, 0.5R, 6.0R, 0.5R, 0.0R}}
            Dim shortLayout As Object = CreateLayout(
                CreateGeometry(RectangleCorners(0.0R, 0.0R, 4.0R, 2.0R, 0.5R), shortPile, Nothing),
                New Size(600, 500))
            Dim shortScale As Double = DoubleProperty(shortLayout, "PixelsPerMetre")
            Dim shortHeads As PointF() = PointArrayProperty(shortLayout, "FrontPileHeads")
            Dim shortToes As PointF() = PointArrayProperty(shortLayout, "FrontPileToes")
            Assert.Multiple(
                Sub()
                    Assert.That(BooleanProperty(shortLayout, "HasPileBreak"), [Is].False)
                    Assert.That(Math.Abs(shortToes(0).Y - shortHeads(0).Y),
                                [Is].EqualTo(5.0R * shortScale).Within(0.03R),
                                "Short piles retain the existing physical vertical scale.")
                End Sub)
        End Sub

        <Test>
        Public Sub LayoutUsesOneScaleAndOneHorizontalMappingForFrontAndPlan()
            Dim piles As Double(,) = {
                {1.0R, 0.5R, 1.0R, 1.25R, 0.75R, 6.0R, 0.5R, 0.0R},
                {3.0R, 1.5R, 1.0R, 3.0R, 1.5R, 5.0R, 0.0R, 0.45R}
            }
            Dim geometry As Object = CreateGeometry(RectangleCorners(0.0R, 0.0R, 4.0R, 2.0R, 0.5R), piles, Nothing)
            Dim layout As Object = CreateLayout(geometry, New Size(900, 620))
            Dim scale As Single = CSng(DoubleProperty(layout, "PixelsPerMetre"))
            Dim frontHeads As PointF() = PointArrayProperty(layout, "FrontPileHeads")
            Dim frontToes As PointF() = PointArrayProperty(layout, "FrontPileToes")
            Dim planHeads As PointF() = PointArrayProperty(layout, "PlanPileHeads")
            Dim frontCorners As PointF() = PointArrayProperty(layout, "FrontGrillageCorners")
            Dim planCorners As PointF() = PointArrayProperty(layout, "PlanGrillageCorners")
            Dim labels As RectangleF() = RectangleArrayProperty(layout, "DimensionLabelBounds")

            Assert.Multiple(
                Sub()
                    Assert.That(scale, [Is].GreaterThan(0.0F))
                    Assert.That(frontHeads.Length, [Is].EqualTo(2))
                    Assert.That(frontToes.Length, [Is].EqualTo(2))
                    Assert.That(planHeads.Length, [Is].EqualTo(2))
                    For index As Integer = 0 To frontHeads.Length - 1
                        Assert.That(frontHeads(index).X, [Is].EqualTo(planHeads(index).X).Within(0.01F),
                                    "The front and plan must share the same horizontal pile mapping.")
                    Next
                    For index As Integer = 0 To frontCorners.Length - 1
                        Assert.That(frontCorners(index).X, [Is].EqualTo(planCorners(index).X).Within(0.01F),
                                    "The front and plan must share the same horizontal grillage mapping.")
                    Next
                    Assert.That(Math.Abs(planHeads(1).X - planHeads(0).X), [Is].EqualTo(2.0F * scale).Within(0.02F))
                    Assert.That(Math.Abs(planHeads(1).Y - planHeads(0).Y), [Is].EqualTo(1.0F * scale).Within(0.02F),
                                "Plan X and Y must use the same aspect scale.")
                    Assert.That(Math.Abs(frontToes(0).X - frontHeads(0).X), [Is].EqualTo(0.25F * scale).Within(0.02F))
                    Assert.That(Math.Abs(frontToes(0).Y - frontHeads(0).Y), [Is].EqualTo(5.0F * scale).Within(0.02F),
                                "The actual toe must be preserved instead of drawing an invented vertical pile.")
                    Assert.That(labels, [Is].Not.Empty)
                    For Each bounds As RectangleF In labels
                        Assert.That(New RectangleF(0.0F, 0.0F, 900.0F, 620.0F).Contains(bounds), [Is].True,
                                    "Dimension labels must remain inside the combined canvas.")
                    Next
                End Sub)
        End Sub

        <Test>
        Public Sub InvalidOrEmptyCalculatedDataStaysNotReadyAndSafe()
            Dim validCorners As Double(,) = RectangleCorners(0.0R, 0.0R, 4.0R, 2.0R, 0.0R)
            Dim onePile As Double(,) = {{1.0R, 0.5R, 1.0R, 1.0R, 0.5R, 6.0R, 0.5R, 0.0R}}
            Dim invalidCorners As Double(,) = {{0.0R, 0.0R, 0.0R}, {0.0R, 1.0R, 0.0R}, {1.0R, 1.0R, 0.0R}}
            Dim nanPiles As Double(,) = DirectCast(onePile.Clone(), Double(,))
            nanPiles(0, 0) = Double.NaN
            Dim cases As Object() = {
                CreateGeometry(Nothing, Nothing, Nothing),
                CreateGeometry(validCorners, Nothing, Nothing),
                CreateGeometry(invalidCorners, onePile, Nothing),
                CreateGeometry(validCorners, nanPiles, Nothing)
            }

            For Each geometry As Object In cases
                Assert.That(BooleanProperty(geometry, "IsReady"), [Is].False)
                Assert.DoesNotThrow(Sub() CreateLayout(geometry, Size.Empty))
            Next
        End Sub

        <Test>
        Public Sub DegenerateCachedAxisIsIgnoredWithoutRejectingAValidPileAtWorldOrigin()
            Dim corners As Double(,) = RectangleCorners(0.0R, 0.0R, 4.0R, 2.0R, 0.0R)
            Dim zeroAxis As Double() = {0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 0.5R, 0.0R}
            Dim validOriginAxis As Double() = {0.0R, 0.0R, 0.0R, 0.0R, 0.0R, 6.0R, 0.5R, 0.0R}
            Dim onlyDefaultAxis As Object = CreateGeometry(corners, Rows(zeroAxis), Nothing)
            Dim mixedAxes As Object = CreateGeometry(corners, Rows(zeroAxis, validOriginAxis), Nothing)

            Assert.Multiple(
                Sub()
                    Assert.That(BooleanProperty(onlyDefaultAxis, "IsReady"), [Is].False)
                    Assert.That(IntegerProperty(onlyDefaultAxis, "PileCount"), [Is].Zero)
                    Assert.That(BooleanProperty(mixedAxes, "IsReady"), [Is].True)
                    Assert.That(IntegerProperty(mixedAxes, "PileCount"), [Is].EqualTo(1))
                    AssertMatrix(MatrixProperty(mixedAxes, "Piles"), Rows(validOriginAxis))
                End Sub)
        End Sub

        <Test>
        Public Sub RendererReturnsRequestedCombinedCanvasForValidGeometry()
            Dim piles As Double(,) = {{1.0R, 0.5R, 1.0R, 1.2R, 0.5R, 6.0R, 0.5R, 0.0R}}
            Dim geometry As Object = CreateGeometry(RectangleCorners(0.0R, 0.0R, 4.0R, 2.0R, 0.5R), piles, Nothing)

            Using rendered As Bitmap = Render(geometry, New Size(840, 560))
                Assert.Multiple(
                    Sub()
                        Assert.That(rendered.Size, [Is].EqualTo(New Size(840, 560)))
                        Assert.That(ContainsColor(rendered, Function(color) color.ToArgb() = Color.White.ToArgb()), [Is].True)
                        Assert.That(ContainsColor(rendered, Function(color) color.ToArgb() <> Color.White.ToArgb()), [Is].True)
                    End Sub)
            End Using
        End Sub

        Private Shared Function RectangleCorners(originX As Double, originY As Double,
                                                  length As Double, width As Double, z As Double) As Double(,)
            Return New Double(,) {{originX, originY, z},
                                  {originX + length, originY, z},
                                  {originX + length, originY + width, z},
                                  {originX, originY + width, z}}
        End Function

        Private Shared Function Rows(ParamArray values As Double()()) As Double(,)
            Dim result(values.Length - 1, values(0).Length - 1) As Double
            For row As Integer = 0 To values.Length - 1
                For column As Integer = 0 To values(row).Length - 1
                    result(row, column) = values(row)(column)
                Next
            Next
            Return result
        End Function

        Private Shared Sub CreateRegularPileGrid(columns As Integer, rows As Integer,
                                                 firstX As Double, firstY As Double,
                                                 stepX As Double, stepY As Double,
                                                 headZ As Double, toeZ As Double,
                                                 addJitter As Boolean,
                                                 ByRef piles As Double(,), ByRef axisData As Double(,))
            ReDim piles(columns * rows - 1, 7)
            ReDim axisData(columns * rows - 1, 3)
            Dim index As Integer = 0
            For row As Integer = 0 To rows - 1
                For column As Integer = 0 To columns - 1
                    Dim nominalX As Double = firstX + column * stepX
                    Dim nominalY As Double = firstY + row * stepY
                    Dim jitterX As Double = If(addJitter, If((row + column) Mod 2 = 0, -0.001R, 0.001R), 0.0R)
                    Dim jitterY As Double = If(addJitter, If((row * columns + column) Mod 3 = 0, 0.001R, -0.001R), 0.0R)
                    piles(index, 0) = nominalX + jitterX
                    piles(index, 1) = nominalY + jitterY
                    piles(index, 2) = headZ
                    piles(index, 3) = nominalX + jitterX
                    piles(index, 4) = nominalY + jitterY
                    piles(index, 5) = toeZ
                    piles(index, 6) = 0.5R
                    piles(index, 7) = 0.0R
                    axisData(index, 0) = row + 1
                    axisData(index, 1) = column + 1
                    axisData(index, 2) = nominalX
                    axisData(index, 3) = nominalY
                    index += 1
                Next
            Next
        End Sub

        Private Shared Function ExpectedRegularSegments(first As Double, repetitions As Integer,
                                                        repeated As Double, last As Double) As Double()
            Dim result(repetitions + 1) As Double
            result(0) = first
            For index As Integer = 1 To repetitions
                result(index) = repeated
            Next
            result(result.Length - 1) = last
            Return result
        End Function

        Private Shared Function TransformXY(source As Double(,), angleDegrees As Double,
                                             offsetX As Double, offsetY As Double,
                                             ParamArray xyColumnPairs As Integer()) As Double(,)
            Dim result As Double(,) = DirectCast(source.Clone(), Double(,))
            Dim radians As Double = angleDegrees * Math.PI / 180.0R
            Dim cosine As Double = Math.Cos(radians)
            Dim sine As Double = Math.Sin(radians)
            For row As Integer = 0 To result.GetLength(0) - 1
                For pairIndex As Integer = 0 To xyColumnPairs.Length - 1 Step 2
                    Dim xColumn As Integer = xyColumnPairs(pairIndex)
                    Dim yColumn As Integer = xyColumnPairs(pairIndex + 1)
                    Dim x As Double = source(row, xColumn)
                    Dim y As Double = source(row, yColumn)
                    result(row, xColumn) = x * cosine - y * sine + offsetX
                    result(row, yColumn) = x * sine + y * cosine + offsetY
                Next
            Next
            Return result
        End Function

        Private Shared Function CreateGeometry(corners As Double(,), piles As Double(,), edges As Double(,)) As Object
            Return InvokeStatic("CivilEnginStructures.PileLayoutGeometry", "Create", corners, piles, edges)
        End Function

        Private Shared Function CreateGeometry(corners As Double(,), piles As Double(,), edges As Double(,),
                                               axisData As Double(,), nominalSize As Double()) As Object
            Return InvokeStatic("CivilEnginStructures.PileLayoutGeometry", "Create",
                                corners, piles, edges, axisData, nominalSize)
        End Function

        Private Shared Function CreateLayout(geometry As Object, requestedSize As Size) As Object
            Return InvokeStatic("CivilEnginStructures.PileLayoutPreview", "CreateLayout", geometry, requestedSize)
        End Function

        Private Shared Function Render(geometry As Object, requestedSize As Size) As Bitmap
            Return DirectCast(InvokeStatic("CivilEnginStructures.PileLayoutPreview", "Render", geometry, requestedSize), Bitmap)
        End Function

        Private Shared Function InvokeStatic(typeName As String, methodName As String,
                                             ParamArray arguments As Object()) As Object
            Dim targetType As Type = GetType(PileLayoutPreviewTests).Assembly.GetType(typeName, False)
            If targetType Is Nothing Then Throw New AssertionException(typeName & " has not been implemented yet.")
            Dim method As MethodInfo = targetType.GetMethods(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static).
                FirstOrDefault(Function(candidate) candidate.Name = methodName AndAlso
                                                   candidate.GetParameters().Length = arguments.Length)
            If method Is Nothing Then Throw New AssertionException(typeName & "." & methodName & " is missing.")
            Try
                Return method.Invoke(Nothing, arguments)
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Function

        Private Shared Function PropertyValue(instance As Object, name As String) As Object
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(
                name, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            If propertyInfo Is Nothing Then Throw New AssertionException(name & " property is missing.")
            Return propertyInfo.GetValue(instance, Nothing)
        End Function

        Private Shared Function BooleanProperty(instance As Object, name As String) As Boolean
            Return CBool(PropertyValue(instance, name))
        End Function

        Private Shared Function IntegerProperty(instance As Object, name As String) As Integer
            Return CInt(PropertyValue(instance, name))
        End Function

        Private Shared Function DoubleProperty(instance As Object, name As String) As Double
            Return CDbl(PropertyValue(instance, name))
        End Function

        Private Shared Function MatrixProperty(instance As Object, name As String) As Double(,)
            Return DirectCast(PropertyValue(instance, name), Double(,))
        End Function

        Private Shared Function DoubleArrayProperty(instance As Object, name As String) As Double()
            Return DirectCast(PropertyValue(instance, name), Double())
        End Function

        Private Shared Function StringArrayProperty(instance As Object, name As String) As String()
            Return DirectCast(PropertyValue(instance, name), String())
        End Function

        Private Shared Function PointArrayProperty(instance As Object, name As String) As PointF()
            Return DirectCast(PropertyValue(instance, name), PointF())
        End Function

        Private Shared Function RectangleArrayProperty(instance As Object, name As String) As RectangleF()
            Return DirectCast(PropertyValue(instance, name), RectangleF())
        End Function

        Private Shared Sub AssertMatrix(actual As Double(,), expected As Double(,))
            Assert.That(actual, [Is].Not.Null)
            Assert.That(actual.GetLength(0), [Is].EqualTo(expected.GetLength(0)))
            Assert.That(actual.GetLength(1), [Is].EqualTo(expected.GetLength(1)))
            For row As Integer = 0 To expected.GetLength(0) - 1
                For column As Integer = 0 To expected.GetLength(1) - 1
                    Assert.That(actual(row, column), [Is].EqualTo(expected(row, column)).Within(0.0000001R),
                                String.Format("row {0}, column {1}", row, column))
                Next
            Next
        End Sub

        Private Shared Sub AssertDoubleArrays(actual As Double(), expected As Double(), tolerance As Double)
            Assert.That(actual.Length, [Is].EqualTo(expected.Length))
            For index As Integer = 0 To expected.Length - 1
                Assert.That(actual(index), [Is].EqualTo(expected(index)).Within(tolerance), "index " & index.ToString())
            Next
        End Sub

        Private Shared Sub AssertLabelsInsideAndSeparate(labels As RectangleF(), canvas As Size)
            Dim canvasBounds As New RectangleF(0.0F, 0.0F, canvas.Width, canvas.Height)
            Assert.That(labels, [Is].Not.Empty)
            For index As Integer = 0 To labels.Length - 1
                Assert.That(canvasBounds.Contains(labels(index)), [Is].True,
                            "Dimension label " & index.ToString() & " must stay inside the canvas.")
                For other As Integer = index + 1 To labels.Length - 1
                    Assert.That(labels(index).IntersectsWith(labels(other)), [Is].False,
                                String.Format("Dimension labels {0} and {1} overlap.", index, other))
                Next
            Next
        End Sub

        Private Shared Sub AssertRectanglesDoNotOverlap(bounds As RectangleF(), description As String)
            For index As Integer = 0 To bounds.Length - 1
                For other As Integer = index + 1 To bounds.Length - 1
                    Assert.That(bounds(index).IntersectsWith(bounds(other)), [Is].False,
                                String.Format("{0} bounds {1} and {2} overlap.", description, index, other))
                Next
            Next
        End Sub

        Private Shared Sub AssertRectanglesDoNotIntersect(first As RectangleF(), second As RectangleF(),
                                                          message As String)
            For Each firstBounds As RectangleF In first
                For Each secondBounds As RectangleF In second
                    Assert.That(firstBounds.IntersectsWith(secondBounds), [Is].False, message)
                Next
            Next
        End Sub

        Private Shared Function ContainsColor(bitmap As Bitmap, predicate As Func(Of Color, Boolean)) As Boolean
            For y As Integer = 0 To bitmap.Height - 1
                For x As Integer = 0 To bitmap.Width - 1
                    If predicate(bitmap.GetPixel(x, y)) Then Return True
                Next
            Next
            Return False
        End Function
    End Class
End Namespace
