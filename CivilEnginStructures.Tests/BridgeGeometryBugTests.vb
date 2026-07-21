Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class BridgeGeometryBugTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub CrossPointInLineRejectsPointTwoMillimetersBeforeLongSegment()
            Dim result As Boolean = BridgeGeometry.crossPointInLine(
                New Vector3D(0, 0, 0),
                New Vector3D(1000, 0, 0),
                New Vector3D(-0.002, 0, 0))

            Assert.That(result, [Is].False)
        End Sub

        <Test>
        Public Sub CalculatePointPWithZeroAngleOffsetsAlongPositiveX()
            Dim result As Vector3D = BridgeGeometry.calculatePointP(
                New Vector3D(0, 0, 0),
                New Vector3D(10, 0, 0),
                2,
                3,
                0)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(2).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(3).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub CalculatePointP2IsRotationCovariantForSlopedBeams()
            Dim resultAlongX As Vector3D = BridgeGeometry.calculatePointP2(
                New Vector3D(0, 0, 0),
                New Vector3D(10, 0, 5),
                2)
            Dim resultAlongY As Vector3D = BridgeGeometry.calculatePointP2(
                New Vector3D(0, 0, 0),
                New Vector3D(0, 10, 5),
                2)

            Assert.Multiple(
                Sub()
                    Assert.That(resultAlongY.X, [Is].EqualTo(-resultAlongX.Y).Within(Tolerance))
                    Assert.That(resultAlongY.Y, [Is].EqualTo(resultAlongX.X).Within(Tolerance))
                    Assert.That(resultAlongY.Z, [Is].EqualTo(resultAlongX.Z).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub CreateRotatedRectanglePreservesThreeMillimeterLengthFraction()
            Const expectedLength As Double = 100000.003
            Dim points As List(Of Vector2D) = BridgeGeometry.createRotatedRectangle(
                New Vector2D(0, 0),
                expectedLength,
                10,
                0)
            Dim actualLength As Double = Math.Sqrt(
                Math.Pow(points(1).X - points(0).X, 2) +
                Math.Pow(points(1).Y - points(0).Y, 2))

            Assert.That(actualLength, [Is].EqualTo(expectedLength).Within(0.001))
        End Sub

        <Test>
        Public Sub CalculateRotatedRectangle3dPreservesThreeMillimeterLengthFraction()
            Const expectedLength As Double = 100000.003
            Dim points As List(Of Vector3D) = BridgeGeometry.calculateRotatedRectangle3d(
                New Vector3D(0, 0, 25),
                expectedLength,
                10,
                0)
            Dim actualLength As Double = Math.Sqrt(
                Math.Pow(points(1).X - points(0).X, 2) +
                Math.Pow(points(1).Y - points(0).Y, 2) +
                Math.Pow(points(1).Z - points(0).Z, 2))

            Assert.That(actualLength, [Is].EqualTo(expectedLength).Within(0.001))
        End Sub
    End Class
End Namespace
