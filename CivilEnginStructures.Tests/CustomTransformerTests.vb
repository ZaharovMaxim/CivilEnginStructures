Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class CustomTransformerTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub TransformPointBeforeInitializationThrowsInvalidOperationException()
            Dim transform As New CustomTransformer()
            Dim point As New AffineTransform3D.Point3D(1, 2, 3)

            Assert.That(
                Sub() transform.TransformPoint(point),
                Throws.TypeOf(Of InvalidOperationException)())
        End Sub

        <Test>
        Public Sub ComputeTransformForKnownScaleAndOffsetsTransformsPointAndHasZeroError()
            Dim transform As New CustomTransformer()
            Dim source As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(2, 4, 8),
                    New AffineTransform3D.Point3D(-4, 1, 9)
                }
            Dim target As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(10, -3, 0),
                    New AffineTransform3D.Point3D(12, 1, 6),
                    New AffineTransform3D.Point3D(6, -2, -12)
                }

            Dim computed As Boolean = transform.ComputeTransform(source, target)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(5, 8, 100))

            Assert.Multiple(
                Sub()
                    Assert.That(computed, [Is].True)
                    Assert.That(transform.IsInitialized, [Is].True)
                    Assert.That(transform.ScaleZ, [Is].EqualTo(3).Within(Tolerance))
                    Assert.That(result.X, [Is].EqualTo(15).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(5).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(15).Within(Tolerance))
                    Assert.That(transform.GetError(source, target), [Is].EqualTo(0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub ComputeTransformSubtractsOffsetWhenCalculatingScaleZ()
            Dim transform As New CustomTransformer()
            Dim source As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 0),
                    New AffineTransform3D.Point3D(2, 0, 0),
                    New AffineTransform3D.Point3D(4, 0, 0)
                }
            Dim target As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 5),
                    New AffineTransform3D.Point3D(2, 0, 11),
                    New AffineTransform3D.Point3D(4, 0, 17)
                }

            transform.ComputeTransform(source, target)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(6, 0, 0))

            Assert.Multiple(
                Sub()
                    Assert.That(transform.ScaleZ, [Is].EqualTo(3).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(23).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub ComputeTransformUsesTargetZAsOffsetAtZeroX()
            Dim transform As New CustomTransformer()
            Dim source As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 100),
                    New AffineTransform3D.Point3D(2, 0, 0)
                }
            Dim target As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(0, 0, 5),
                    New AffineTransform3D.Point3D(2, 0, 11)
                }

            transform.ComputeTransform(source, target)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(0, 0, -500))

            Assert.That(result.Z, [Is].EqualTo(5).Within(Tolerance))
        End Sub

        <Test>
        Public Sub ComputeTransformRecoversXYOffsetsWithoutZeroXAnchor()
            Dim transform As New CustomTransformer()
            Dim source As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(1, 5, 0),
                    New AffineTransform3D.Point3D(2, 7, 0)
                }
            Dim target As New List(Of AffineTransform3D.Point3D) From
                {
                    New AffineTransform3D.Point3D(11, 1, 2),
                    New AffineTransform3D.Point3D(12, 3, 4)
                }

            transform.ComputeTransform(source, target)
            Dim result As AffineTransform3D.Point3D =
                transform.TransformPoint(New AffineTransform3D.Point3D(3, 9, 0))

            Assert.Multiple(
                Sub()
                    Assert.That(result.X, [Is].EqualTo(13).Within(Tolerance))
                    Assert.That(result.Y, [Is].EqualTo(5).Within(Tolerance))
                    Assert.That(result.Z, [Is].EqualTo(6).Within(Tolerance))
                End Sub)
        End Sub
    End Class
End Namespace
