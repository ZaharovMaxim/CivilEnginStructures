Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities

Namespace Tests
    <TestFixture>
    Public Class PerpendicularOffsetLineTests
        Private Const CoordinateTolerance As Double = 0.000000001
        Private Const AngleToleranceDegrees As Double = 0.000000001

        <Test>
        Public Sub SlopedThreeDimensionalLineCreatesExactNinetyDegreeOffsetWithoutRounding()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(0, 0, 0),
                New Vector3D(3, 4, 12))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 13)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    AssertPoint(result.StartPoint, -7.2, -9.6, 5)
                    AssertPoint(result.EndPoint, -4.2, -5.6, 17)
                    Assert.That(CalculateOffsetAngleDegrees(sourceLine, result),
                                [Is].EqualTo(90).Within(AngleToleranceDegrees))
                    Assert.That(CalculateOffsetLength(sourceLine, result),
                                [Is].EqualTo(13).Within(CoordinateTolerance))
                    AssertLinesAreParallelWithSameDirection(sourceLine, result)
                End Sub)
        End Sub

        <Test>
        Public Sub HorizontalLineOffsetsVerticallyAndRemainsParallel()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(10, -2, 7),
                New Vector3D(13, 2, 7))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 2.5)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    AssertPoint(result.StartPoint, 10, -2, 9.5)
                    AssertPoint(result.EndPoint, 13, 2, 9.5)
                    Assert.That(CalculateOffsetAngleDegrees(sourceLine, result),
                                [Is].EqualTo(90).Within(AngleToleranceDegrees))
                    Assert.That(CalculateOffsetLength(sourceLine, result),
                                [Is].EqualTo(2.5).Within(CoordinateTolerance))
                    AssertLinesAreParallelWithSameDirection(sourceLine, result)
                End Sub)
        End Sub

        <Test>
        Public Sub NegativeHeightMovesToOppositeSideAndUsesAbsoluteOffsetLength()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(0, 0, 0),
                New Vector3D(3, 4, 12))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, -13)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    AssertPoint(result.StartPoint, 7.2, 9.6, -5)
                    AssertPoint(result.EndPoint, 10.2, 13.6, 7)
                    Assert.That(CalculateOffsetAngleDegrees(sourceLine, result),
                                [Is].EqualTo(90).Within(AngleToleranceDegrees))
                    Assert.That(CalculateOffsetLength(sourceLine, result),
                                [Is].EqualTo(13).Within(CoordinateTolerance))
                    AssertLinesAreParallelWithSameDirection(sourceLine, result)
                End Sub)
        End Sub

        <Test>
        Public Sub ReversedSlopedLineKeepsPositiveHeightOnTheSameUpwardSide()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(3, 4, 12),
                New Vector3D(0, 0, 0))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 13)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    AssertPoint(result.StartPoint, -4.2, -5.6, 17)
                    AssertPoint(result.EndPoint, -7.2, -9.6, 5)
                    Assert.That(CalculateOffsetAngleDegrees(sourceLine, result),
                                [Is].EqualTo(90).Within(AngleToleranceDegrees))
                    Assert.That(CalculateOffsetLength(sourceLine, result),
                                [Is].EqualTo(13).Within(CoordinateTolerance))
                    AssertLinesAreParallelWithSameDirection(sourceLine, result)
                End Sub)
        End Sub

        <Test>
        Public Sub VerticalLineOffsetsAlongNegativeXUsingLegacyDirection()
            Using sourceLine As DwgLine = CreateLine(
                    New Vector3D(0, 0, 0),
                    New Vector3D(0, 0, 10))
                Using result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 5)
                    Assert.That(result, [Is].Not.Null)
                    Assert.Multiple(
                        Sub()
                            AssertPoint(result.StartPoint, -5, 0, 0)
                            AssertPoint(result.EndPoint, -5, 0, 10)
                            Assert.That(CalculateDirectionOffsetDotProduct(sourceLine, result),
                                        [Is].EqualTo(0).Within(CoordinateTolerance))
                            Assert.That(CalculateOffsetLength(sourceLine, result),
                                        [Is].EqualTo(5).Within(CoordinateTolerance))
                            AssertLinesAreParallelWithSameDirection(sourceLine, result)
                        End Sub)
                End Using
            End Using
        End Sub

        <Test>
        Public Sub ReversedVerticalLineOffsetsAlongPositiveXUsingLegacyDirection()
            Using sourceLine As DwgLine = CreateLine(
                    New Vector3D(0, 0, 10),
                    New Vector3D(0, 0, 0))
                Using result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 5)
                    Assert.That(result, [Is].Not.Null)
                    Assert.Multiple(
                        Sub()
                            AssertPoint(result.StartPoint, 5, 0, 10)
                            AssertPoint(result.EndPoint, 5, 0, 0)
                            Assert.That(CalculateDirectionOffsetDotProduct(sourceLine, result),
                                        [Is].EqualTo(0).Within(CoordinateTolerance))
                            Assert.That(CalculateOffsetLength(sourceLine, result),
                                        [Is].EqualTo(5).Within(CoordinateTolerance))
                            AssertLinesAreParallelWithSameDirection(sourceLine, result)
                        End Sub)
                End Using
            End Using
        End Sub

        <Test>
        Public Sub NegativeHeightReversesVerticalLineOffsetSide()
            Using sourceLine As DwgLine = CreateLine(
                    New Vector3D(0, 0, 0),
                    New Vector3D(0, 0, 10))
                Using result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, -5)
                    Assert.That(result, [Is].Not.Null)
                    Assert.Multiple(
                        Sub()
                            AssertPoint(result.StartPoint, 5, 0, 0)
                            AssertPoint(result.EndPoint, 5, 0, 10)
                            Assert.That(CalculateDirectionOffsetDotProduct(sourceLine, result),
                                        [Is].EqualTo(0).Within(CoordinateTolerance))
                            Assert.That(CalculateOffsetLength(sourceLine, result),
                                        [Is].EqualTo(5).Within(CoordinateTolerance))
                            AssertLinesAreParallelWithSameDirection(sourceLine, result)
                        End Sub)
                End Using
            End Using
        End Sub

        <Test>
        Public Sub ResultIsANewLineAndSourceLineRemainsUnchanged()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(1.25, -2.5, 3.75),
                New Vector3D(4.25, 1.5, 15.75))
            Dim originalStart As Vector3D = sourceLine.StartPoint
            Dim originalEnd As Vector3D = sourceLine.EndPoint

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 6.5)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].Not.SameAs(sourceLine))
                    AssertPoint(sourceLine.StartPoint, originalStart.X, originalStart.Y, originalStart.Z)
                    AssertPoint(sourceLine.EndPoint, originalEnd.X, originalEnd.Y, originalEnd.Z)
                End Sub)
        End Sub

        <Test>
        Public Sub ZeroHeightReturnsIndependentCoincidentLine()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(1, 2, 3),
                New Vector3D(4, 6, 15))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 0)

            Assert.That(result, [Is].Not.Null)
            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].Not.SameAs(sourceLine))
                    AssertPoint(result.StartPoint, 1, 2, 3)
                    AssertPoint(result.EndPoint, 4, 6, 15)
                    AssertLinesAreParallelWithSameDirection(sourceLine, result)
                End Sub)
        End Sub

        <Test>
        Public Sub ZeroLengthLineReturnsNothing()
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(1, 2, 3),
                New Vector3D(1, 2, 3))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, 5)

            Assert.That(result, [Is].Null)
        End Sub

        <Test>
        Public Sub NothingSourceLineReturnsNothing()
            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(Nothing, 5)

            Assert.That(result, [Is].Null)
        End Sub

        <TestCase(Double.NaN)>
        <TestCase(Double.PositiveInfinity)>
        <TestCase(Double.NegativeInfinity)>
        Public Sub NonFiniteHeightReturnsNothing(offsetHeight As Double)
            Dim sourceLine As DwgLine = CreateLine(
                New Vector3D(0, 0, 0),
                New Vector3D(3, 4, 12))

            Dim result As DwgLine = BridgeGeometry.createPerpendicularOffsetLine(sourceLine, offsetHeight)

            Assert.That(result, [Is].Null)
        End Sub

        Private Shared Function CreateLine(startPoint As Vector3D, endPoint As Vector3D) As DwgLine
            Dim line As New DwgLine()
            line.StartPoint = startPoint
            line.EndPoint = endPoint
            Return line
        End Function

        Private Shared Sub AssertPoint(actual As Vector3D,
                                       expectedX As Double,
                                       expectedY As Double,
                                       expectedZ As Double)
            Assert.Multiple(
                Sub()
                    Assert.That(actual.X, [Is].EqualTo(expectedX).Within(CoordinateTolerance))
                    Assert.That(actual.Y, [Is].EqualTo(expectedY).Within(CoordinateTolerance))
                    Assert.That(actual.Z, [Is].EqualTo(expectedZ).Within(CoordinateTolerance))
                End Sub)
        End Sub

        Private Shared Function CalculateOffsetLength(sourceLine As DwgLine, result As DwgLine) As Double
            Dim offsetX As Double = result.StartPoint.X - sourceLine.StartPoint.X
            Dim offsetY As Double = result.StartPoint.Y - sourceLine.StartPoint.Y
            Dim offsetZ As Double = result.StartPoint.Z - sourceLine.StartPoint.Z
            Return Math.Sqrt(offsetX * offsetX + offsetY * offsetY + offsetZ * offsetZ)
        End Function

        Private Shared Function CalculateDirectionOffsetDotProduct(sourceLine As DwgLine,
                                                                    result As DwgLine) As Double
            Dim directionX As Double = sourceLine.EndPoint.X - sourceLine.StartPoint.X
            Dim directionY As Double = sourceLine.EndPoint.Y - sourceLine.StartPoint.Y
            Dim directionZ As Double = sourceLine.EndPoint.Z - sourceLine.StartPoint.Z
            Dim offsetX As Double = result.StartPoint.X - sourceLine.StartPoint.X
            Dim offsetY As Double = result.StartPoint.Y - sourceLine.StartPoint.Y
            Dim offsetZ As Double = result.StartPoint.Z - sourceLine.StartPoint.Z
            Return directionX * offsetX + directionY * offsetY + directionZ * offsetZ
        End Function

        Private Shared Function CalculateOffsetAngleDegrees(sourceLine As DwgLine, result As DwgLine) As Double
            Dim directionX As Double = sourceLine.EndPoint.X - sourceLine.StartPoint.X
            Dim directionY As Double = sourceLine.EndPoint.Y - sourceLine.StartPoint.Y
            Dim directionZ As Double = sourceLine.EndPoint.Z - sourceLine.StartPoint.Z
            Dim offsetX As Double = result.StartPoint.X - sourceLine.StartPoint.X
            Dim offsetY As Double = result.StartPoint.Y - sourceLine.StartPoint.Y
            Dim offsetZ As Double = result.StartPoint.Z - sourceLine.StartPoint.Z
            Dim dotProduct As Double = directionX * offsetX + directionY * offsetY + directionZ * offsetZ
            Dim directionLength As Double = Math.Sqrt(
                directionX * directionX + directionY * directionY + directionZ * directionZ)
            Dim offsetLength As Double = Math.Sqrt(
                offsetX * offsetX + offsetY * offsetY + offsetZ * offsetZ)
            Dim cosine As Double = dotProduct / (directionLength * offsetLength)
            cosine = Math.Max(-1, Math.Min(1, cosine))
            Return Math.Acos(cosine) * 180 / Math.PI
        End Function

        Private Shared Sub AssertLinesAreParallelWithSameDirection(sourceLine As DwgLine, result As DwgLine)
            Assert.Multiple(
                Sub()
                    Assert.That(result.EndPoint.X - result.StartPoint.X,
                                [Is].EqualTo(sourceLine.EndPoint.X - sourceLine.StartPoint.X).
                                    Within(CoordinateTolerance))
                    Assert.That(result.EndPoint.Y - result.StartPoint.Y,
                                [Is].EqualTo(sourceLine.EndPoint.Y - sourceLine.StartPoint.Y).
                                    Within(CoordinateTolerance))
                    Assert.That(result.EndPoint.Z - result.StartPoint.Z,
                                [Is].EqualTo(sourceLine.EndPoint.Z - sourceLine.StartPoint.Z).
                                    Within(CoordinateTolerance))
                End Sub)
        End Sub
    End Class
End Namespace
