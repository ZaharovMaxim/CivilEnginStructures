Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class PureMathCharacterizationTests
        Private Const Tolerance As Double = 0.000000001

        <Test>
        Public Sub ArePointsCoplanarDistinguishesPlaneFromTetrahedron()
            Dim p1 As New UPoint3D(0, 0, 2)
            Dim p2 As New UPoint3D(3, 0, 2)
            Dim p3 As New UPoint3D(0, 4, 2)

            Assert.Multiple(
                Sub()
                    Assert.That(UPlane3D.ArePointsCoplanar(p1, p2, p3, New UPoint3D(2, 1, 2)), [Is].True)
                    Assert.That(UPlane3D.ArePointsCoplanar(p1, p2, p3, New UPoint3D(2, 1, 3)), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub PlaneFromPointsIntersectsSegmentAtExpectedPoint()
            Dim plane As UPlane = UPlane3D.GetPlaneFromPoints(
                New UPoint3D(0, 0, 2),
                New UPoint3D(3, 0, 2),
                New UPoint3D(0, 4, 2))

            Dim intersection As UPoint3D = UPlane3D.FindIntersection(
                New UPoint3D(1, 1, 0),
                New UPoint3D(1, 1, 4),
                plane)

            Assert.Multiple(
                Sub()
                    Assert.That(plane.Normal.X, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(plane.Normal.Y, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(plane.Normal.Z, [Is].EqualTo(12).Within(Tolerance))
                    Assert.That(plane.D, [Is].EqualTo(-24).Within(Tolerance))
                    Assert.That(intersection.X, [Is].EqualTo(1).Within(Tolerance))
                    Assert.That(intersection.Y, [Is].EqualTo(1).Within(Tolerance))
                    Assert.That(intersection.Z, [Is].EqualTo(2).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub FindIntersectionReturnsNothingWhenSegmentDoesNotCrossPlane()
            Dim plane As New UPlane(0, 0, 1, -2)

            Dim parallel As UPoint3D = UPlane3D.FindIntersection(
                New UPoint3D(0, 0, 1),
                New UPoint3D(2, 0, 1),
                plane)
            Dim outsideSegment As UPoint3D = UPlane3D.FindIntersection(
                New UPoint3D(0, 0, 3),
                New UPoint3D(0, 0, 4),
                plane)

            Assert.Multiple(
                Sub()
                    Assert.That(parallel, [Is].Null)
                    Assert.That(outsideSegment, [Is].Null)
                End Sub)
        End Sub

        <Test>
        Public Sub CoordinateDistanceAndMidpointFunctionsApplyRequestedRounding()
            Dim point As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(
                New Vector2D(1.234, -2.346),
                0,
                2,
                2)
            Dim distance As Double = MathFunction.funcCalcDistanceByToPoints2d(
                New Vector2D(0, 0),
                New Vector2D(1, 1),
                3)
            Dim midpoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(
                New Vector2D(0.001, 0.003),
                New Vector2D(0.004, 0.008),
                3)

            Assert.Multiple(
                Sub()
                    Assert.That(point.X, [Is].EqualTo(3.23).Within(Tolerance))
                    Assert.That(point.Y, [Is].EqualTo(-2.35).Within(Tolerance))
                    Assert.That(distance, [Is].EqualTo(1.414).Within(Tolerance))
                    Assert.That(midpoint.X, [Is].EqualTo(0.002).Within(Tolerance))
                    Assert.That(midpoint.Y, [Is].EqualTo(0.006).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub VectorProductsAndNormalizationReturnExpectedComponents()
            Dim cross As Vector3D = MathFunction.CrossProduct(
                New Vector3D(1, 2, 3),
                New Vector3D(4, 5, 6))
            Dim dot As Double = MathFunction.DotProduct(
                New Vector3D(1, 2, 3),
                New Vector3D(4, 5, 6))
            Dim normal As Vector3D = MathFunction.GetNormal(New Vector3D(0, 3, 4))
            Dim zeroNormal As Vector3D = MathFunction.GetNormal(New Vector3D(0, 0, 0))

            Assert.Multiple(
                Sub()
                    Assert.That(cross.X, [Is].EqualTo(-3).Within(Tolerance))
                    Assert.That(cross.Y, [Is].EqualTo(6).Within(Tolerance))
                    Assert.That(cross.Z, [Is].EqualTo(-3).Within(Tolerance))
                    Assert.That(dot, [Is].EqualTo(32).Within(Tolerance))
                    Assert.That(normal.X, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(normal.Y, [Is].EqualTo(0.6).Within(Tolerance))
                    Assert.That(normal.Z, [Is].EqualTo(0.8).Within(Tolerance))
                    Assert.That(zeroNormal.Length, [Is].EqualTo(0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub TriangleCentersMatchThreeFourFiveRightTriangle()
            Dim a As New Vector2D(0, 0)
            Dim b As New Vector2D(4, 0)
            Dim c As New Vector2D(0, 3)

            Dim centroid As Vector2D = TriangleCenter.FindCentroid(a, b, c)
            Dim circumcenter As Vector2D = TriangleCenter.FindCircumcenter(a, b, c)
            Dim incenter As Vector2D = TriangleCenter.FindIncenter(a, b, c)
            Dim orthocenter As Vector2D = TriangleCenter.FindOrthocenter(a, b, c)

            Assert.Multiple(
                Sub()
                    Assert.That(centroid.X, [Is].EqualTo(4.0 / 3.0).Within(Tolerance))
                    Assert.That(centroid.Y, [Is].EqualTo(1).Within(Tolerance))
                    Assert.That(circumcenter.X, [Is].EqualTo(2).Within(Tolerance))
                    Assert.That(circumcenter.Y, [Is].EqualTo(1.5).Within(Tolerance))
                    Assert.That(incenter.X, [Is].EqualTo(1).Within(Tolerance))
                    Assert.That(incenter.Y, [Is].EqualTo(1).Within(Tolerance))
                    Assert.That(orthocenter.X, [Is].EqualTo(0).Within(Tolerance))
                    Assert.That(orthocenter.Y, [Is].EqualTo(0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub CircumcenterRejectsCollinearPoints()
            Assert.That(
                Sub() TriangleCenter.FindCircumcenter(
                    New Vector2D(0, 0),
                    New Vector2D(1, 1),
                    New Vector2D(2, 2)),
                Throws.TypeOf(Of ArgumentException)())
        End Sub
    End Class

    <TestFixture>
    <SetCulture("en-US")>
    Public Class TextFunctionCharacterizationTests
        <TestCase(-3, "Л-3")>
        <TestCase(0, "Ось")>
        <TestCase(4, "П-4")>
        Public Sub GetConditionalRowFormatsSideAndNumber(number As Integer, expected As String)
            Assert.That(FuncFormatZn.getConditionalRow(number), [Is].EqualTo(expected))
        End Sub

        <Test>
        Public Sub FuncFormatCoordinateUsesFixedRequestedPrecision()
            Assert.Multiple(
                Sub()
                    Assert.That(FuncFormatZn.FuncFormatCoordinate(12.3456, 3), [Is].EqualTo("12.346"))
                    Assert.That(FuncFormatZn.FuncFormatCoordinate(12, 0), [Is].EqualTo("12"))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncConvertStrNumberPointMergesAdjacentRangesAndKeepsGap()
            Dim result As String = FuncFormatZn.FuncConvertStrNumberPoint("1-3,4-5,8,9")

            Assert.That(result, [Is].EqualTo("1-5,8,9"))
        End Sub

        <Test>
        Public Sub CadastralHelpersSeparateObjectNumberAndContourNumber()
            Const value As String = " 77:01:000(контур 12) "

            Assert.Multiple(
                Sub()
                    Assert.That(FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(value), [Is].EqualTo("77:01:000"))
                    Assert.That(FuncFormatZn.FuncFindNumberCounterCadastralObject(value), [Is].EqualTo(12))
                    Assert.That(FuncFormatZn.FuncFindFullConditionalNumberCadastralObject(Nothing), [Is].Null)
                    Assert.That(FuncFormatZn.FuncFindNumberCounterCadastralObject(Nothing), [Is].EqualTo(0))
                End Sub)
        End Sub

        <Test>
        Public Sub FuncValidateJsonReportsValidContainersAndRejectsBlankInput()
            Dim objectType As String = Nothing
            Dim arrayType As String = Nothing
            Dim errorMessage As String = Nothing

            Dim objectValid As Boolean = FuncFormatZn.FuncValidateJson("{""name"":""bridge""}", jsonType:=objectType)
            Dim arrayValid As Boolean = FuncFormatZn.FuncValidateJson("[1,2,3]", jsonType:=arrayType)
            Dim blankValid As Boolean = FuncFormatZn.FuncValidateJson("   ", errorMessage:=errorMessage)

            Assert.Multiple(
                Sub()
                    Assert.That(objectValid, [Is].True)
                    Assert.That(objectType, [Is].EqualTo("Object"))
                    Assert.That(arrayValid, [Is].True)
                    Assert.That(arrayType, [Is].EqualTo("Array"))
                    Assert.That(blankValid, [Is].False)
                    Assert.That(errorMessage, [Is].EqualTo("Строка пустая или Nothing"))
                End Sub)
        End Sub
    End Class

    <TestFixture>
    Public Class TriangulationValueCharacterizationTests
        <Test>
        Public Sub TriangleCanonicalizesVertexOrderForEqualityAndDisplay()
            Dim triangle As New Triangle(5, 1, 3)
            Dim sameVertices As New Triangle(3, 5, 1)

            Assert.Multiple(
                Sub()
                    Assert.That(triangle.Vertices, [Is].EqualTo(New Integer() {1, 3, 5}))
                    Assert.That(triangle.ContainsVertex(3), [Is].True)
                    Assert.That(triangle.ContainsVertex(4), [Is].False)
                    Assert.That(triangle, [Is].EqualTo(sameVertices))
                    Assert.That(triangle.GetHashCode(), [Is].EqualTo(sameVertices.GetHashCode()))
                    Assert.That(triangle.ToString(), [Is].EqualTo("[1, 3, 5]"))
                End Sub)
        End Sub

        <Test>
        Public Sub EdgeCanonicalizesDirectionForEqualityAndDisplay()
            Dim edge As New Edge(7, 2)
            Dim reversed As New Edge(2, 7)

            Assert.Multiple(
                Sub()
                    Assert.That(edge.StartVertex, [Is].EqualTo(2))
                    Assert.That(edge.EndVertex, [Is].EqualTo(7))
                    Assert.That(edge, [Is].EqualTo(reversed))
                    Assert.That(edge.GetHashCode(), [Is].EqualTo(reversed.GetHashCode()))
                    Assert.That(edge.ToString(), [Is].EqualTo("[2-7]"))
                End Sub)
        End Sub

        <Test>
        Public Sub TriangulateRejectsMissingOrInsufficientPoints()
            Dim triangulation As New DelaunayTriangulation()

            Assert.Multiple(
                Sub()
                    Assert.That(
                        Sub() triangulation.Triangulate(Nothing),
                        Throws.TypeOf(Of ArgumentException)())
                    Assert.That(
                        Sub() triangulation.Triangulate(
                            New List(Of Vector2D) From {New Vector2D(0, 0), New Vector2D(1, 0)}),
                        Throws.TypeOf(Of ArgumentException)())
                End Sub)
        End Sub

        <Test>
        Public Sub TriangulateOrdinaryTriangleReturnsItsThreeVerticesWithoutMutatingInput()
            Dim points As New List(Of Vector2D) From
                {
                    New Vector2D(0, 0),
                    New Vector2D(10, 0),
                    New Vector2D(0, 10)
                }
            Dim triangulation As New DelaunayTriangulation()

            Dim result As Dictionary(Of Integer, List(Of Vector2D)) = triangulation.Triangulate(points)

            Assert.Multiple(
                Sub()
                    Assert.That(points.Count, [Is].EqualTo(3))
                    Assert.That(result.Count, [Is].EqualTo(1))
                    Assert.That(result(0), Has.Count.EqualTo(3))
                    Assert.That(triangulation.VerifyDelaunay(), [Is].True)
                End Sub)
        End Sub
    End Class
End Namespace
