Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class AffineTransform3DTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub NewTransformPointReturnsTheOriginalPoint()
            Dim transform As New AffineTransform3D()
            Dim source As New AffineTransform3D.Point3D(12.5, -7.25, 103.75)

            Dim result As AffineTransform3D.Point3D = transform.TransformPoint(source)

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(source.X).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(source.Y).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(source.Z).Within(Tolerance))
                    Assert.That(transform.Matrix(0, 0), [Is].EqualTo(1.0))
                    Assert.That(transform.Matrix(1, 1), [Is].EqualTo(1.0))
                    Assert.That(transform.Matrix(2, 2), [Is].EqualTo(1.0))
                    Assert.That(transform.Matrix(3, 3), [Is].EqualTo(1.0))
                End Sub)
        End Sub

        <Test>
        Public Sub ComputeFrom4PointsForTranslationTransformsAnotherPoint()
            Dim transform As New AffineTransform3D()
            Dim source As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(1, 0, 0),
                    New AffineTransform3D.Point3D(0, 1, 0),
                    New AffineTransform3D.Point3D(0, 0, 1)
                }
            Dim target As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(10, -20, 5),
                    New AffineTransform3D.Point3D(11, -20, 5),
                    New AffineTransform3D.Point3D(10, -19, 5),
                    New AffineTransform3D.Point3D(10, -20, 6)
                }

            Dim computed As Boolean = transform.ComputeFrom4Points(source, target)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(2, -3, 4))

            Assert.Multiple(
                Sub()
                    Assert.That(computed, [Is].True)
                    Assert.That(result.X, [Is].EqualTo(12).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(-23).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(9).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub ComputeFrom4PointsForDegeneratePointsReturnsFalse()
            Dim transform As New AffineTransform3D()
            Dim source As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(1, 0, 0),
                    New AffineTransform3D.Point3D(0, 1, 0),
                    New AffineTransform3D.Point3D(1, 1, 0)
                }
            Dim target As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(5, 5, 5),
                    New AffineTransform3D.Point3D(6, 5, 5),
                    New AffineTransform3D.Point3D(5, 6, 5),
                    New AffineTransform3D.Point3D(6, 6, 5)
                }

            Dim computed As Boolean = transform.ComputeFrom4Points(source, target)

            Assert.That(computed, [Is].False)
        End Sub

        <Test>
        Public Sub ComputeFromPointsForStandardTetrahedronPreservesIdentity()
            Dim transform As New AffineTransform3D()
            Dim points As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(1, 0, 0),
                    New AffineTransform3D.Point3D(0, 1, 0),
                    New AffineTransform3D.Point3D(0, 0, 1)
                }

            Dim computed As Boolean = transform.ComputeFromPoints(points, points)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(0.2, 0.3, 0.4))

            Assert.Multiple(
                Sub()
                    Assert.That(computed, [Is].True)
                    Assert.That(result.X, [Is].EqualTo(0.2).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(0.3).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(0.4).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub ComputeFrom4PointsAcceptsWellConditionedSmallTetrahedron()
            Const scale As Double = 0.0001
            Dim transform As New AffineTransform3D()
            Dim points As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(scale, 0, 0),
                    New AffineTransform3D.Point3D(0, scale, 0),
                    New AffineTransform3D.Point3D(0, 0, scale)
                }

            Dim computed As Boolean = transform.ComputeFrom4Points(points, points)

            Assert.That(computed, [Is].True)
        End Sub

        <Test>
        Public Sub ComputeFrom4PointsIsInvariantToLargeTranslation()
            Const origin As Double = 100000000.0
            Const LargeCoordinateTolerance As Double = 0.000001
            Dim transform As New AffineTransform3D()
            Dim points As AffineTransform3D.Point3D() =
                {
                    New AffineTransform3D.Point3D(origin, origin, origin),
                    New AffineTransform3D.Point3D(origin + 1, origin, origin),
                    New AffineTransform3D.Point3D(origin, origin + 1, origin),
                    New AffineTransform3D.Point3D(origin, origin, origin + 1)
                }

            Dim computed As Boolean = transform.ComputeFrom4Points(points, points)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(origin + 0.2, origin + 0.3, origin + 0.4))

            Assert.Multiple(
                Sub()
                    Assert.That(computed, [Is].True)
                    Assert.That(result.X, [Is].EqualTo(origin + 0.2).Within(LargeCoordinateTolerance))
                    Assert.That(result.Y, [Is].EqualTo(origin + 0.3).Within(LargeCoordinateTolerance))
                    Assert.That(result.Z, [Is].EqualTo(origin + 0.4).Within(LargeCoordinateTolerance))
                End Sub)
        End Sub
    End Class
End Namespace
