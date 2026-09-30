Imports NUnit.Framework
Imports System.Reflection
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class DelaunayTriangulationBugTests
        <Test>
        Public Sub TriangulateReturnsOneTriangleAtSmallCoordinateScale()
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(0, 0),
                    New Vector2D(0.0001, 0),
                    New Vector2D(0, 0.0001)
                }
            Dim triangulation As New DelaunayTriangulation()

            Dim result As Dictionary(Of Integer, List(Of Vector2D)) =
                triangulation.Triangulate(points)

            Assert.That(result.Count, [Is].EqualTo(1))
        End Sub

        <Test>
        Public Sub TriangulateReturnsOneTriangleAfterLargeTranslation()
            Const origin As Double = 1000000000.0
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(origin, origin),
                    New Vector2D(origin + 1, origin),
                    New Vector2D(origin, origin + 1)
                }
            Dim triangulation As New DelaunayTriangulation()

            Dim result As Dictionary(Of Integer, List(Of Vector2D)) =
                triangulation.Triangulate(points)

            Assert.That(result.Count, [Is].EqualTo(1))
        End Sub

        <Test>
        Public Sub VerifyDelaunayAcceptsCocircularSquareTriangulation()
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(0, 0),
                    New Vector2D(1, 0),
                    New Vector2D(1, 1),
                    New Vector2D(0, 1)
                }
            Dim triangulation As New DelaunayTriangulation()

            Dim result As Dictionary(Of Integer, List(Of Vector2D)) =
                triangulation.Triangulate(points)

            Assert.Multiple(
                Sub()
                    Assert.That(result.Count, [Is].EqualTo(2))
                    Assert.That(triangulation.VerifyDelaunay(), [Is].True)
                End Sub)
        End Sub

        <Test>
        Public Sub TriangulateRetriangulatesForPointStrictlyInsideCircumcircle()
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(0, 0),
                    New Vector2D(4, 0),
                    New Vector2D(0, 4),
                    New Vector2D(1, 1)
                }
            Dim triangulation As New DelaunayTriangulation()

            Dim result As Dictionary(Of Integer, List(Of Vector2D)) =
                triangulation.Triangulate(points)

            Assert.Multiple(
                Sub()
                    Assert.That(result.Count, [Is].EqualTo(3))
                    Assert.That(triangulation.VerifyDelaunay(), [Is].True)
                End Sub)
        End Sub

        <Test>
        Public Sub VerifyDelaunayRejectsStrictlyInteriorPointAtSmallScale()
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(0, 0),
                    New Vector2D(0.0000001, 0),
                    New Vector2D(0, 0.0000001),
                    New Vector2D(0.00000002, 0.00000002)
                }
            Dim triangles As New HashSet(Of Triangle) From
                {
                    New Triangle(0, 1, 2)
                }
            Dim triangulation As New DelaunayTriangulation()
            Dim pointsField As FieldInfo = GetType(DelaunayTriangulation).GetField(
                "Points", BindingFlags.NonPublic Or BindingFlags.Instance)
            Dim trianglesField As FieldInfo = GetType(DelaunayTriangulation).GetField(
                "Triangles", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(pointsField, [Is].Not.Null)
            Assert.That(trianglesField, [Is].Not.Null)
            pointsField.SetValue(triangulation, points)
            trianglesField.SetValue(triangulation, triangles)

            Assert.That(triangulation.VerifyDelaunay(), [Is].False,
                        "A relative tolerance must not hide a strictly interior point at small scale.")
        End Sub
    End Class
End Namespace
