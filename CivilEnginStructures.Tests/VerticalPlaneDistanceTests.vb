Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class VerticalPlaneDistanceTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub CrossingVerticalPlaneReturnsNegativeDistanceToIntersection()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(5, 0, -10),
                New Vector3D(5, 10, 20),
                New Vector3D(1, 2, 3),
                New Vector3D(9, 2, 3))

            Assert.That(result, [Is].EqualTo(-4.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub InclinedPlaneLineAndLargeHeightDifferenceUseInfiniteVerticalPlane()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(0, 0, -1000000),
                New Vector3D(10, 10, 1000000),
                New Vector3D(0, 10, 1000),
                New Vector3D(10, 0, -1000))

            Dim expected As Double = -System.Math.Sqrt(1000050.0)

            Assert.That(result, [Is].EqualTo(expected).Within(Tolerance))
        End Sub

        <Test>
        Public Sub SegmentThatStopsBeforePlaneUsesDistanceFromStartOnly()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(10, 0, 0),
                New Vector3D(10, 1, 0),
                New Vector3D(0, 0, 50),
                New Vector3D(4, 0, -50))

            Assert.That(result, [Is].EqualTo(10.0).Within(Tolerance),
                "The endpoint is 6 units from the plane, but only segmentStart is compared.")
        End Sub

        <Test>
        Public Sub SegmentMovingAwayFromPlaneUsesPositiveStartDistance()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(0, 0, 0),
                New Vector3D(0, 10, 0),
                New Vector3D(2, 4, 8),
                New Vector3D(8, 4, -8))

            Assert.That(result, [Is].EqualTo(2.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub SegmentParallelToPlaneUsesPositivePerpendicularStartDistance()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(0, 0, -100),
                New Vector3D(0, 10, 100),
                New Vector3D(3, 0, 0),
                New Vector3D(3, 100, 1000))

            Assert.That(result, [Is].EqualTo(3.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub SegmentStartingOnPlaneReturnsZero()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(0, 0, 0),
                New Vector3D(0, 10, 10),
                New Vector3D(0, 4, 20),
                New Vector3D(7, 4, -20))

            Assert.That(result, [Is].EqualTo(0.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub SegmentEndingOnPlaneCountsAsIntersection()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(0, 0, 0),
                New Vector3D(0, 10, 0),
                New Vector3D(3, 4, 5),
                New Vector3D(0, 4, 5))

            Assert.That(result, [Is].EqualTo(-3.0).Within(Tolerance))
        End Sub

        <Test>
        Public Sub CoincidentPlanePointsInXYReturnNaN()
            Dim result As Double = MathFunction.SignedDistanceFromSegmentStartToVerticalPlane(
                New Vector3D(1, 2, -1000),
                New Vector3D(1, 2, 1000),
                New Vector3D(0, 0, 0),
                New Vector3D(10, 10, 10))

            Assert.That(Double.IsNaN(result), [Is].True)
        End Sub
    End Class
End Namespace
