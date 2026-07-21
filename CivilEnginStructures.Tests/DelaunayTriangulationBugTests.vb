Imports NUnit.Framework
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
    End Class
End Namespace
