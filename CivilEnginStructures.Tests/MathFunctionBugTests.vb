Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class MathFunctionBugTests
        Private Const Tolerance As Double = 0.000001

        <Test>
        Public Sub FuncIntersectionTwoRayRejectsIntersectionBehindVerticalFirstRay()
            Dim intersection As New Vector2D()

            Dim intersects As Boolean = MathFunction.FuncIntersectionTwoRay(
                New Vector2D(0, 0),
                New Vector2D(0, 10),
                New Vector2D(-5, -5),
                New Vector2D(5, -5),
                intersection)

            Assert.That(intersects, [Is].False,
                "The infinite lines cross at (0, -5), but that point is behind the first ray.")
        End Sub

        <Test>
        Public Sub FuncIntersectionTwoRayFindsIntersectionAheadOfVerticalFirstRay()
            Dim intersection As New Vector2D()

            Dim intersects As Boolean = MathFunction.FuncIntersectionTwoRay(
                New Vector2D(0, 0),
                New Vector2D(0, 10),
                New Vector2D(-5, 5),
                New Vector2D(5, 5),
                intersection)

            Assert.Multiple(
                Sub()
                    Assert.That(intersects, [Is].True)
                    Assert.That(intersection.X, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(intersection.Y, [Is].EqualTo(5).Within(Tolerance))
                End Sub)
        End Sub

        <TestCase(-1.0, 1.0, 2.356194490192345)>
        <TestCase(1.0, -1.0, 0.7853981633974483)>
        Public Sub FuncCalcAngleByToPoints2dReturnsAngleBetweenPositiveXAxisAndVector(
            x As Double,
            y As Double,
            expected As Double)

            Dim result As Double = MathFunction.funcCalcAngleByToPoints2d(
                New Vector2D(0, 0),
                New Vector2D(x, y))

            Assert.That(result, [Is].EqualTo(expected).Within(Tolerance))
        End Sub

        <TestCase(1.0, -1.0, 135.0)>
        <TestCase(-1.0, 1.0, 315.0)>
        Public Sub FuncCalcDirectionAngleByToPoints2dReturnsAzimuthInEveryQuadrant(
            x As Double,
            y As Double,
            expected As Double)

            Dim result As Double = MathFunction.funcCalcDirectionAngleByToPoints2d(
                New Vector2D(0, 0),
                New Vector2D(x, y))

            Assert.That(result, [Is].EqualTo(expected).Within(Tolerance))
        End Sub

        <TestCase(1.0, 0.0, 0.0)>
        <TestCase(0.0, 1.0, 1.5707963267948966)>
        <TestCase(-1.0, 0.0, 3.1415926535897931)>
        <TestCase(0.0, -1.0, 1.5707963267948966)>
        Public Sub FuncCalcAngleCardinalDirectionsAreUnsignedFromPositiveX(
            x As Double,
            y As Double,
            expected As Double)

            Dim result As Double = MathFunction.funcCalcAngleByToPoints2d(
                New Vector2D(0, 0),
                New Vector2D(x, y))

            Assert.That(result, [Is].EqualTo(expected).Within(Tolerance))
        End Sub

        <TestCase(0.0, 1.0, 0.0)>
        <TestCase(1.0, 0.0, 90.0)>
        <TestCase(0.0, -1.0, 180.0)>
        <TestCase(-1.0, 0.0, 270.0)>
        Public Sub FuncCalcDirectionCardinalDirectionsUseFullAzimuth(
            x As Double,
            y As Double,
            expected As Double)

            Dim result As Double = MathFunction.funcCalcDirectionAngleByToPoints2d(
                New Vector2D(0, 0),
                New Vector2D(x, y))

            Assert.That(result, [Is].EqualTo(expected).Within(Tolerance))
        End Sub

        <Test>
        Public Sub FuncSqrtPolylineByArrayReturnsRectangleArea()
            Dim rectangle As Double(,) =
                {
                    {0.0, 4.0, 4.0, 0.0},
                    {0.0, 0.0, 3.0, 3.0}
                }

            Dim area As Double = MathFunction.FuncSQRTPolylineByArray(rectangle)

            Assert.That(area, [Is].EqualTo(12.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub MatrixMultiplicationAcceptsTwoByThreeTimesThreeByOne()
            Dim matrixA As Double(,) =
                {
                    {1.0, 2.0, 3.0},
                    {4.0, 5.0, 6.0}
                }
            Dim matrixB As Double(,) =
                {
                    {7.0},
                    {8.0},
                    {9.0}
                }

            Dim result As Double(,) = MathFunction.MatrixMultiplication(matrixA, matrixB)

            Assert.Multiple(
                Sub()
                    Assert.That(result.GetLength(0), [Is].EqualTo(2))
                    Assert.That(result.GetLength(1), [Is].EqualTo(1))
                    Assert.That(result(0, 0), [Is].EqualTo(50.0).Within(Tolerance))
                    Assert.That(result(1, 0), [Is].EqualTo(122.0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub MatrixMultiplicationMultipliesFourByFourMatricesWithoutChangingShape()
            Dim matrixA As Double(,) =
                {
                    {1.0, 2.0, 0.0, 1.0},
                    {0.0, 1.0, 3.0, 0.0},
                    {2.0, 0.0, 1.0, 4.0},
                    {0.0, 0.0, 0.0, 1.0}
                }
            Dim matrixB As Double(,) =
                {
                    {2.0, 0.0, 1.0, 0.0},
                    {0.0, 3.0, 0.0, 1.0},
                    {1.0, 0.0, 4.0, 0.0},
                    {0.0, 0.0, 0.0, 1.0}
                }
            Dim expected As Double(,) =
                {
                    {2.0, 6.0, 1.0, 3.0},
                    {3.0, 3.0, 12.0, 1.0},
                    {5.0, 0.0, 6.0, 4.0},
                    {0.0, 0.0, 0.0, 1.0}
                }

            Dim result As Double(,) = MathFunction.MatrixMultiplication(matrixA, matrixB)

            Assert.Multiple(
                Sub()
                    Assert.That(result.GetLength(0), [Is].EqualTo(4))
                    Assert.That(result.GetLength(1), [Is].EqualTo(4))
                    Assert.That(result, [Is].EqualTo(expected))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncCalculatePositionAttributeWithZeroOffsetsReturnsVertex2()
            Dim vertex2 As New Vector2D(10.5, -3.25)

            Dim result As Vector2D = MathFunction.FuncCalculatePositionAttribute(
                New Vector2D(0, 0),
                vertex2,
                New Vector2D(20, 0),
                0,
                0)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(vertex2.X).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(vertex2.Y).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncCalculatePositionAttributeReturnsExistingCalculatedPointForNonzeroOffsets()
            Dim result As Vector2D = MathFunction.FuncCalculatePositionAttribute(
                New Vector2D(0, 0),
                New Vector2D(10, 0),
                New Vector2D(20, 0),
                2,
                3)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(10).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(-3).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncCrossPointPolylineRecognizesInteriorPointInFractionalRectangle()
            Dim rectangle As Double(,) =
                {
                    {0.2, 2.2, 2.2, 0.2},
                    {0.2, 0.2, 2.2, 2.2}
                }

            Dim inside As Boolean = MathFunction.FuncCrossPointPolyline(
                rectangle,
                New Vector2D(0.35, 2.05))

            Assert.That(inside, [Is].True)
        End Sub

        <Test>
        Public Sub FuncSortDblArrayReturnsFalseForOutOfRangeRowIndex()
            Dim values As Double(,) = {{2.0, 1.0}}

            Dim sorted As Boolean = MathFunction.FuncSortDblArray(values, 1)

            Assert.That(sorted, [Is].False)
        End Sub

        <Test>
        Public Sub FuncSortDblArray2PreservesFractionalSortKeys()
            Dim values As Double(,) =
                {
                    {1.4, 1.2, 1.3},
                    {0.0, 0.0, 0.0},
                    {14.0, 12.0, 13.0}
                }

            Dim sorted As Boolean = MathFunction.FuncSortDblArray2(values, 0, 1)
            Dim actual As Double() = {values(0, 0), values(0, 1), values(0, 2)}

            Assert.Multiple(
                Sub()
                    Assert.That(sorted, [Is].True)
                    Assert.That(actual, [Is].EqualTo(New Double() {1.2, 1.3, 1.4}))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncSortDblArray2UpdatesCurrentMinimumAfterSwap()
            Dim values As Double(,) =
                {
                    {3.0, 1.0, 2.0},
                    {0.0, 0.0, 0.0},
                    {30.0, 10.0, 20.0}
                }

            Dim sorted As Boolean = MathFunction.FuncSortDblArray2(values, 0, 1)
            Dim actual As Double() = {values(0, 0), values(0, 1), values(0, 2)}

            Assert.Multiple(
                Sub()
                    Assert.That(sorted, [Is].True)
                    Assert.That(actual, [Is].EqualTo(New Double() {1.0, 2.0, 3.0}))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncSortStrArraySortsNonNumericValues()
            Dim values As String(,) = {{"b", "a"}}

            Dim sorted As Boolean = MathFunction.FuncSortStrArray(values, 0)
            Dim actual As String() = {values(0, 0), values(0, 1)}

            Assert.Multiple(
                Sub()
                    Assert.That(sorted, [Is].True)
                    Assert.That(actual, [Is].EqualTo(New String() {"a", "b"}))
                End Sub)
        End Sub
    End Class
End Namespace
