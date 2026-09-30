Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities

Namespace Tests
    <TestFixture>
    Public Class BridgeGeometryBugTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub MoveLineToPointMovesStartToPositionRoundedToThreeDecimals()
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim position As New Vector2D(123.45678, -98.76543)

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(line, position)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(line.StartPoint.X, [Is].EqualTo(123.457).Within(Tolerance))
                    Assert.That(line.StartPoint.Y, [Is].EqualTo(-98.765).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointPreservesPlanVectorLengthAndRotation()
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim originalDeltaX As Double = line.EndPoint.X - line.StartPoint.X
            Dim originalDeltaY As Double = line.EndPoint.Y - line.StartPoint.Y
            Dim originalPlanLength As Double = Math.Sqrt(
                Math.Pow(originalDeltaX, 2) +
                Math.Pow(originalDeltaY, 2))
            Dim originalRotation As Double = line.Rotation

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543))
            Dim actualDeltaX As Double = line.EndPoint.X - line.StartPoint.X
            Dim actualDeltaY As Double = line.EndPoint.Y - line.StartPoint.Y
            Dim actualPlanLength As Double = Math.Sqrt(
                Math.Pow(actualDeltaX, 2) +
                Math.Pow(actualDeltaY, 2))

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(actualDeltaX, [Is].EqualTo(originalDeltaX).Within(Tolerance))
                    Assert.That(actualDeltaY, [Is].EqualTo(originalDeltaY).Within(Tolerance))
                    Assert.That(actualPlanLength, [Is].EqualTo(originalPlanLength).Within(Tolerance))
                    Assert.That(line.Rotation, [Is].EqualTo(originalRotation).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointPreservesThreeDimensionalLengthWithinOneMillimeter()
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim originalLength As Double = line.Length

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543))

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(line.Length, [Is].EqualTo(originalLength).Within(0.001))
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointPreservesStartAndEndElevations()
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim originalStartZ As Double = line.StartPoint.Z
            Dim originalEndZ As Double = line.EndPoint.Z

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543))

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(line.StartPoint.Z, [Is].EqualTo(originalStartZ).Within(Tolerance))
                    Assert.That(line.EndPoint.Z, [Is].EqualTo(originalEndZ).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointWithCustomRoundRoundsStartAndPreservesPlanVector()
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim originalDeltaX As Double = line.EndPoint.X - line.StartPoint.X
            Dim originalDeltaY As Double = line.EndPoint.Y - line.StartPoint.Y

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543),
                2)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(line.StartPoint.X, [Is].EqualTo(123.46).Within(Tolerance))
                    Assert.That(line.StartPoint.Y, [Is].EqualTo(-98.77).Within(Tolerance))
                    Assert.That(line.EndPoint.X - line.StartPoint.X, [Is].EqualTo(originalDeltaX).Within(Tolerance))
                    Assert.That(line.EndPoint.Y - line.StartPoint.Y, [Is].EqualTo(originalDeltaY).Within(Tolerance))
                End Sub)
        End Sub

        <TestCase(-1)>
        <TestCase(16)>
        Public Sub MoveLineToPointWithInvalidRoundReturnsFalseAndLeavesLineUnchanged(round As Integer)
            Dim line As DwgLine = CreateMoveLineToPointTestLine()
            Dim originalStart As Vector3D = line.StartPoint
            Dim originalEnd As Vector3D = line.EndPoint

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543),
                round)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    AssertLinePointsEqual(line, originalStart, originalEnd)
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointWhenResultOverflowsReturnsFalseAndLeavesLineUnchanged()
            Dim line As New DwgLine()
            line.StartPoint = New Vector3D(0, 2, 3)
            line.EndPoint = New Vector3D(Double.MaxValue, 5, 7)
            Dim originalStart As Vector3D = line.StartPoint
            Dim originalEnd As Vector3D = line.EndPoint

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(Double.MaxValue, 10))

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    AssertLinePointsEqual(line, originalStart, originalEnd)
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointWhenCoordinatesLoseLineLengthReturnsFalseAndLeavesLineUnchanged()
            Dim line As New DwgLine()
            line.StartPoint = New Vector3D(0, 0, 3)
            line.EndPoint = New Vector3D(1, 0, 7)
            Dim originalStart As Vector3D = line.StartPoint
            Dim originalEnd As Vector3D = line.EndPoint

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(1.0E+16, 10))

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    AssertLinePointsEqual(line, originalStart, originalEnd)
                End Sub)
        End Sub

        <Test>
        Public Sub MoveLineToPointReturnsFalseForNothing()
            Dim line As DwgLine = Nothing

            Dim result As Boolean = BridgeGeometry.moveLineToPoint(
                line,
                New Vector2D(123.45678, -98.76543))

            Assert.That(result, [Is].False)
        End Sub

        <Test>
        Public Sub CrossPointInLineRejectsPointTwoMillimetersBeforeLongSegment()
            Dim result As Boolean = BridgeGeometry.crossPointInLine(
                New Vector3D(0, 0, 0),
                New Vector3D(1000, 0, 0),
                New Vector3D(-0.002, 0, 0))

            Assert.That(result, [Is].False)
        End Sub

        <Test>
        Public Sub CrossPointInLineAcceptsPointHalfMillimeterBeforeLongSegment()
            Dim result As Boolean = BridgeGeometry.crossPointInLine(
                New Vector3D(0, 0, 0),
                New Vector3D(1000, 0, 0),
                New Vector3D(-0.0005, 0, 0))

            Assert.That(result, [Is].True,
                        "The endpoint tolerance is one millimeter in model units.")
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
        Public Sub CalculatePointPDefaultAngleOffsetsLeftOfAxis()
            Dim result As Vector3D = BridgeGeometry.calculatePointP(
                New Vector3D(0, 0, 10),
                New Vector3D(10, 0, 20),
                2,
                3)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(2).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(13).Within(Tolerance))
                End Sub)
        End Sub

        <TestCase(0.0, 2.0, 0.0)>
        <TestCase(1.5707963267948966, 0.0, 2.0)>
        <TestCase(-1.5707963267948966, 0.0, -2.0)>
        Public Sub CalculatePointPAngleIsMeasuredFromAxis(angle As Double,
                                                         expectedX As Double,
                                                         expectedY As Double)
            Dim result As Vector3D = BridgeGeometry.calculatePointP(
                New Vector3D(0, 0, 7),
                New Vector3D(10, 0, 9),
                2,
                5,
                angle)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(expectedX).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(expectedY).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(12).Within(Tolerance))
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

        Private Shared Function CreateMoveLineToPointTestLine() As DwgLine
            Dim line As New DwgLine()
            line.StartPoint = New Vector3D(10.1234, -20.5678, 7.25)
            line.EndPoint = New Vector3D(13.6234, -16.0678, 11.75)
            Return line
        End Function

        Private Shared Sub AssertLinePointsEqual(
            line As DwgLine,
            expectedStart As Vector3D,
            expectedEnd As Vector3D)

            Assert.That(line.StartPoint.X, [Is].EqualTo(expectedStart.X))
            Assert.That(line.StartPoint.Y, [Is].EqualTo(expectedStart.Y))
            Assert.That(line.StartPoint.Z, [Is].EqualTo(expectedStart.Z))
            Assert.That(line.EndPoint.X, [Is].EqualTo(expectedEnd.X))
            Assert.That(line.EndPoint.Y, [Is].EqualTo(expectedEnd.Y))
            Assert.That(line.EndPoint.Z, [Is].EqualTo(expectedEnd.Z))
        End Sub
    End Class
End Namespace
