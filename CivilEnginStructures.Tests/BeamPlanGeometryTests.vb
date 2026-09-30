Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class BeamPlanGeometryTests
        Private Const CoordinateTolerance As Double = 0.000001R
        Private Const AreaTolerance As Double = 0.00001R
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub BuildOutlineHasStablePublicMathOnlyContract()
            Dim geometryType As Type = MainType("CivilEnginStructures.BridgeBeamPlanGeometry")
            Dim method As MethodInfo = BuildOutlineMethod()
            Dim parameters As ParameterInfo() = method.GetParameters()

            Assert.Multiple(
                Sub()
                    Assert.That(geometryType.IsSealed, [Is].True,
                                "BridgeBeamPlanGeometry must remain NotInheritable.")
                    Assert.That(method.IsStatic, [Is].True)
                    Assert.That(method.ReturnType, [Is].EqualTo(GetType(Vector2D())))
                    Assert.That(parameters, Has.Length.EqualTo(2))
                    Assert.That(parameters(0).ParameterType,
                                [Is].EqualTo(GetType(IEnumerable(Of Vector2D))))
                    Assert.That(parameters(1).ParameterType,
                                [Is].EqualTo(GetType(IEnumerable(Of Vector2D))))
                End Sub)
        End Sub

        <Test>
        Public Sub CrossingRectanglesProduceSingleSteppedUnionInsteadOfConvexHull()
            Dim top As Vector2D() = Rectangle(1.0R, 9.0R, -1.0R, 1.0R)
            Dim bottom As Vector2D() = Rectangle(0.0R, 10.0R, -0.25R, 0.25R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            Dim expected As Vector2D() = {
                Point(0.0R, -0.25R), Point(1.0R, -0.25R), Point(1.0R, -1.0R),
                Point(9.0R, -1.0R), Point(9.0R, -0.25R), Point(10.0R, -0.25R),
                Point(10.0R, 0.25R), Point(9.0R, 0.25R), Point(9.0R, 1.0R),
                Point(1.0R, 1.0R), Point(1.0R, 0.25R), Point(0.0R, 0.25R)
            }
            AssertSingleRing(outline, 17.0R, 12)
            AssertContainsExactly(outline, expected)
        End Sub

        <TestCase(9.0R, 18.0R)>
        <TestCase(10.0R, 20.0R)>
        Public Sub ZeroEndExtensionsCollapseToOuterRectangle(length As Double,
                                                             expectedArea As Double)
            Dim top As Vector2D() = Rectangle(0.0R, length, -1.0R, 1.0R)
            Dim bottom As Vector2D() = Rectangle(0.0R, length, -0.25R, 0.25R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, expectedArea, 4)
            AssertContainsExactly(outline, top)
        End Sub

        <Test>
        Public Sub UnequalStartAndEndExtensionsPreserveBothSteps()
            Dim top As Vector2D() = Rectangle(1.0R, 8.0R, -1.0R, 1.0R)
            Dim bottom As Vector2D() = Rectangle(0.0R, 10.0R, -0.25R, 0.25R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 15.5R, 12)
            AssertContainsNode(outline, Point(1.0R, -0.25R))
            AssertContainsNode(outline, Point(8.0R, 0.25R))
            AssertContainsNode(outline, Point(10.0R, -0.25R))
        End Sub

        <Test>
        Public Sub VerticallyAsymmetricTopPreservesItsOffsetStepLevels()
            Dim top As Vector2D() = Rectangle(1.0R, 9.0R, -0.6R, 1.4R)
            Dim bottom As Vector2D() = Rectangle(0.0R, 10.0R, -0.25R, 0.25R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 17.0R, 12)
            AssertContainsNode(outline, Point(1.0R, -0.6R))
            AssertContainsNode(outline, Point(9.0R, -0.6R))
            AssertContainsNode(outline, Point(1.0R, 1.4R))
            AssertContainsNode(outline, Point(9.0R, 1.4R))
            AssertContainsNode(outline, Point(0.0R, -0.25R))
            AssertContainsNode(outline, Point(10.0R, 0.25R))
        End Sub

        <TestCase(0.0R, 9.0R)>
        <TestCase(1.0R, 10.0R)>
        Public Sub OneSidedExtensionProducesOnlyTheExposedStep(bottomStart As Double,
                                                               bottomEnd As Double)
            Dim top As Vector2D() = Rectangle(1.0R, 9.0R, -1.0R, 1.0R)
            Dim bottom As Vector2D() = Rectangle(bottomStart, bottomEnd, -0.25R, 0.25R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 16.5R)
            Dim exposedX As Double = If(bottomStart < 1.0R, bottomStart, bottomEnd)
            Dim junctionX As Double = If(bottomStart < 1.0R, 1.0R, 9.0R)
            AssertContainsNode(outline, Point(exposedX, -0.25R))
            AssertContainsNode(outline, Point(exposedX, 0.25R))
            AssertContainsNode(outline, Point(junctionX, -0.25R))
            AssertContainsNode(outline, Point(junctionX, 0.25R))
        End Sub

        <Test>
        Public Sub WiderBottomContainsTopWithoutLeavingInternalSeams()
            Dim top As Vector2D() = Rectangle(1.0R, 9.0R, -0.25R, 0.25R)
            Dim bottom As Vector2D() = Rectangle(0.0R, 10.0R, -1.0R, 1.0R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 20.0R, 4)
            AssertContainsExactly(outline, bottom)
        End Sub

        <Test>
        Public Sub SharedFullEdgeIsAValidSinglePolygonAndSharedEdgeIsRemoved()
            Dim top As Vector2D() = Rectangle(0.0R, 1.0R, 0.0R, 1.0R)
            Dim bottom As Vector2D() = Rectangle(1.0R, 2.0R, 0.0R, 1.0R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 2.0R)
            AssertContainsNode(outline, Point(0.0R, 0.0R))
            AssertContainsNode(outline, Point(2.0R, 1.0R))
            Assert.That(HasVerticalEdge(outline, 1.0R, 0.0R, 1.0R), [Is].False,
                        "The shared polygon edge is internal and must not be emitted.")
        End Sub

        <Test>
        Public Sub GeneralConcavePolygonIsPreservedWithoutConvexHullExpansion()
            Dim top As Vector2D() = {
                Point(0.0R, 0.0R), Point(3.0R, 0.0R), Point(3.0R, 1.0R),
                Point(1.0R, 1.0R), Point(1.0R, 3.0R), Point(0.0R, 3.0R)
            }
            Dim bottom As Vector2D() = Rectangle(0.1R, 0.9R, 0.1R, 0.9R)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 5.0R, 6)
            AssertContainsExactly(outline, top)
            Assert.That(ContainsNode(outline, Point(3.0R, 3.0R)), [Is].False,
                        "A concavity must not be replaced with a convex hull.")
        End Sub

        <Test>
        Public Sub RotatedTranslatedLargeCoordinatesRetainAreaAndStepNodes()
            Dim sourceTop As Vector2D() = Rectangle(1.0R, 9.0R, -1.0R, 1.0R)
            Dim sourceBottom As Vector2D() = Rectangle(0.0R, 10.0R, -0.25R, 0.25R)
            Dim angle As Double = 0.63R
            Dim offset As New Vector2D(1000000000.0R, -2000000000.0R)
            Dim top As Vector2D() = Transform(sourceTop, angle, offset)
            Dim bottom As Vector2D() = Transform(sourceBottom, angle, offset)

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 17.0R, 12, 0.0001R)
            AssertContainsNode(outline, Transform(Point(0.0R, -0.25R), angle, offset),
                               0.000001R)
            AssertContainsNode(outline, Transform(Point(9.0R, 1.0R), angle, offset),
                               0.000001R)
        End Sub

        <Test>
        Public Sub ReversedAndExplicitlyClosedInputsProduceTheSameGeometricRing()
            Dim top As Vector2D() = CloseRing(Rectangle(1.0R, 9.0R, -1.0R, 1.0R).Reverse().ToArray())
            Dim bottom As Vector2D() = CloseRing(Rectangle(0.0R, 10.0R, -0.25R, 0.25R).Reverse().ToArray())

            Dim outline As Vector2D() = BuildOutline(top, bottom)

            AssertSingleRing(outline, 17.0R, 12)
            AssertContainsNode(outline, Point(0.0R, -0.25R))
            AssertContainsNode(outline, Point(1.0R, -1.0R))
            AssertContainsNode(outline, Point(9.0R, 1.0R))
            AssertContainsNode(outline, Point(10.0R, 0.25R))
        End Sub

        <Test>
        Public Sub SlopedShiftedAxesUseXYUnionRatherThanIndependentEndDistances()
            Dim localTop As Vector2D() = Rectangle(1.25R, 9.25R, -1.0R, 1.0R)
            Dim localBottom As Vector2D() = Rectangle(0.0R, 10.0R, -0.25R, 0.25R)
            Dim angle As Double = Math.PI / 5.0R
            Dim offset As New Vector2D(350.0R, -125.0R)

            Dim outline As Vector2D() = BuildOutline(
                Transform(localTop, angle, offset),
                Transform(localBottom, angle, offset))

            AssertSingleRing(outline, 17.0R, 12)
            AssertContainsNode(outline, Transform(Point(1.25R, -1.0R), angle, offset))
            AssertContainsNode(outline, Transform(Point(9.25R, 0.25R), angle, offset))
            AssertContainsNode(outline, Transform(Point(10.0R, -0.25R), angle, offset))
        End Sub

        <TestCase("null top")>
        <TestCase("null bottom")>
        <TestCase("empty top")>
        <TestCase("empty bottom")>
        <TestCase("too few vertices")>
        <TestCase("NaN coordinate")>
        <TestCase("infinite coordinate")>
        <TestCase("self intersection")>
        <TestCase("zero area")>
        <TestCase("disjoint polygons")>
        <TestCase("point touch")>
        Public Sub InvalidOrNonPolygonalInputsFailWithUnderstandableReason(scenario As String)
            Dim inputs As Tuple(Of IEnumerable(Of Vector2D), IEnumerable(Of Vector2D)) =
                InvalidInputs(scenario)

            Dim exception As Exception = Assert.Catch(
                Sub() BuildOutline(inputs.Item1, inputs.Item2),
                scenario & " must not be hidden as an empty or successful outline.")

            Assert.Multiple(
                Sub()
                    Assert.That(exception,
                                [Is].InstanceOf(Of ArgumentException)().Or.
                                    InstanceOf(Of InvalidOperationException)(),
                                scenario)
                    Assert.That(exception.Message, [Is].Not.Null.And.Not.Empty, scenario)
                    Assert.That(exception.Message.Trim().Length, [Is].GreaterThanOrEqualTo(4),
                                scenario & " needs a useful failure reason.")
                End Sub)
        End Sub

        Private Shared Function InvalidInputs(scenario As String) As Tuple(Of IEnumerable(Of Vector2D), IEnumerable(Of Vector2D))
            Dim valid As Vector2D() = Rectangle(0.0R, 2.0R, 0.0R, 2.0R)
            Select Case scenario
                Case "null top"
                    Return Tuple.Create(DirectCast(Nothing, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "null bottom"
                    Return Tuple.Create(DirectCast(valid, IEnumerable(Of Vector2D)),
                                        DirectCast(Nothing, IEnumerable(Of Vector2D)))
                Case "empty top"
                    Return Tuple.Create(DirectCast(New Vector2D() {}, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "empty bottom"
                    Return Tuple.Create(DirectCast(valid, IEnumerable(Of Vector2D)),
                                        DirectCast(New Vector2D() {}, IEnumerable(Of Vector2D)))
                Case "too few vertices"
                    Return Tuple.Create(DirectCast(New Vector2D() {Point(0, 0), Point(1, 0)}, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "NaN coordinate"
                    Return Tuple.Create(DirectCast(New Vector2D() {
                                            Point(0, 0), Point(Double.NaN, 0), Point(0, 1)
                                        }, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "infinite coordinate"
                    Return Tuple.Create(DirectCast(New Vector2D() {
                                            Point(0, 0), Point(Double.PositiveInfinity, 0), Point(0, 1)
                                        }, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "self intersection"
                    Return Tuple.Create(DirectCast(New Vector2D() {
                                            Point(0, 0), Point(2, 2), Point(0, 2), Point(2, 0)
                                        }, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "zero area"
                    Return Tuple.Create(DirectCast(New Vector2D() {
                                            Point(0, 0), Point(1, 0), Point(2, 0)
                                        }, IEnumerable(Of Vector2D)),
                                        DirectCast(valid, IEnumerable(Of Vector2D)))
                Case "disjoint polygons"
                    Return Tuple.Create(DirectCast(Rectangle(0, 1, 0, 1), IEnumerable(Of Vector2D)),
                                        DirectCast(Rectangle(2, 3, 0, 1), IEnumerable(Of Vector2D)))
                Case "point touch"
                    Return Tuple.Create(DirectCast(Rectangle(0, 1, 0, 1), IEnumerable(Of Vector2D)),
                                        DirectCast(Rectangle(1, 2, 1, 2), IEnumerable(Of Vector2D)))
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(scenario))
            End Select
        End Function

        Private Shared Function BuildOutline(top As IEnumerable(Of Vector2D),
                                             bottom As IEnumerable(Of Vector2D)) As Vector2D()
            Dim result As Object = InvokeWithoutWrapper(
                BuildOutlineMethod(), Nothing, New Object() {top, bottom})
            Assert.That(result, [Is].Not.Null)
            Return DirectCast(result, Vector2D())
        End Function

        Private Shared Function BuildOutlineMethod() As MethodInfo
            Dim geometryType As Type = MainType("CivilEnginStructures.BridgeBeamPlanGeometry")
            Dim method As MethodInfo = geometryType.GetMethod(
                "BuildOutline",
                BindingFlags.Public Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {
                    GetType(IEnumerable(Of Vector2D)), GetType(IEnumerable(Of Vector2D))
                },
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeBeamPlanGeometry.BuildOutline(top, bottom) must exist in the production assembly.")
            Return method
        End Function

        Private Shared Function InvokeWithoutWrapper(method As MethodInfo,
                                                     target As Object,
                                                     arguments As Object()) As Object
            Try
                Return method.Invoke(target, arguments)
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then Throw exception.InnerException
                Throw
            End Try
        End Function

        Private Shared Sub AssertSingleRing(outline As Vector2D(),
                                            expectedArea As Double,
                                            Optional expectedUniqueVertices As Integer = -1,
                                            Optional areaToleranceOverride As Double = AreaTolerance)
            Assert.That(outline, [Is].Not.Null)
            Assert.That(outline.Length, [Is].GreaterThanOrEqualTo(4))
            Assert.That(AreSamePoint(outline(0), outline(outline.Length - 1)), [Is].True,
                        "The outline must be explicitly closed.")
            For Each vertex As Vector2D In outline
                Assert.That(IsFinite(vertex.X) AndAlso IsFinite(vertex.Y), [Is].True,
                            "Every outline coordinate must be finite.")
            Next

            Dim signedArea As Double = CalculateSignedArea(outline)
            Assert.That(signedArea, [Is].GreaterThan(0.0R),
                        "The closed outline must use counter-clockwise orientation.")
            Assert.That(signedArea, [Is].EqualTo(expectedArea).Within(areaToleranceOverride))
            AssertNoSelfIntersections(outline)
            If expectedUniqueVertices >= 0 Then
                Assert.That(UniqueVertices(outline).Count,
                            [Is].EqualTo(expectedUniqueVertices))
            End If
        End Sub

        Private Shared Sub AssertContainsExactly(outline As Vector2D(), expectedOpenRing As Vector2D())
            Dim actual As List(Of Vector2D) = UniqueVertices(outline)
            Dim expected As List(Of Vector2D) = UniqueVertices(expectedOpenRing)
            Assert.That(actual.Count, [Is].EqualTo(expected.Count))
            For Each vertex As Vector2D In expected
                AssertContainsNode(outline, vertex)
            Next
        End Sub

        Private Shared Sub AssertContainsNode(outline As IEnumerable(Of Vector2D),
                                              expected As Vector2D,
                                              Optional tolerance As Double = CoordinateTolerance)
            Assert.That(ContainsNode(outline, expected, tolerance), [Is].True,
                        "Expected outline node (" & expected.X & ", " & expected.Y & ").")
        End Sub

        Private Shared Function ContainsNode(points As IEnumerable(Of Vector2D),
                                             expected As Vector2D,
                                             Optional tolerance As Double = CoordinateTolerance) As Boolean
            Return points.Any(Function(pointValue) AreSamePoint(pointValue, expected, tolerance))
        End Function

        Private Shared Function UniqueVertices(points As IEnumerable(Of Vector2D)) As List(Of Vector2D)
            Dim result As New List(Of Vector2D)()
            For Each vertex As Vector2D In points
                If Not result.Any(Function(existing) AreSamePoint(existing, vertex)) Then
                    result.Add(vertex)
                End If
            Next
            Return result
        End Function

        Private Shared Function CalculateSignedArea(closedRing As Vector2D()) As Double
            Dim origin As Vector2D = closedRing(0)
            Dim twiceArea As Double = 0.0R
            For index As Integer = 0 To closedRing.Length - 2
                Dim currentX As Double = closedRing(index).X - origin.X
                Dim currentY As Double = closedRing(index).Y - origin.Y
                Dim nextX As Double = closedRing(index + 1).X - origin.X
                Dim nextY As Double = closedRing(index + 1).Y - origin.Y
                twiceArea += currentX * nextY - nextX * currentY
            Next
            Return twiceArea / 2.0R
        End Function

        Private Shared Sub AssertNoSelfIntersections(closedRing As Vector2D())
            Dim segmentCount As Integer = closedRing.Length - 1
            For firstIndex As Integer = 0 To segmentCount - 1
                For secondIndex As Integer = firstIndex + 1 To segmentCount - 1
                    Dim adjacent As Boolean = secondIndex = firstIndex + 1 OrElse
                        (firstIndex = 0 AndAlso secondIndex = segmentCount - 1)
                    If Not adjacent Then
                        Assert.That(SegmentsIntersect(
                                        closedRing(firstIndex), closedRing(firstIndex + 1),
                                        closedRing(secondIndex), closedRing(secondIndex + 1)),
                                    [Is].False,
                                    "A single outline ring must not self-intersect.")
                    End If
                Next
            Next
        End Sub

        Private Shared Function SegmentsIntersect(a As Vector2D,
                                                  b As Vector2D,
                                                  c As Vector2D,
                                                  d As Vector2D) As Boolean
            Dim abC As Double = Cross(a, b, c)
            Dim abD As Double = Cross(a, b, d)
            Dim cdA As Double = Cross(c, d, a)
            Dim cdB As Double = Cross(c, d, b)
            If OppositeSigns(abC, abD) AndAlso OppositeSigns(cdA, cdB) Then Return True
            Return Math.Abs(abC) <= CoordinateTolerance AndAlso OnSegment(a, b, c) OrElse
                   Math.Abs(abD) <= CoordinateTolerance AndAlso OnSegment(a, b, d) OrElse
                   Math.Abs(cdA) <= CoordinateTolerance AndAlso OnSegment(c, d, a) OrElse
                   Math.Abs(cdB) <= CoordinateTolerance AndAlso OnSegment(c, d, b)
        End Function

        Private Shared Function Cross(a As Vector2D, b As Vector2D, c As Vector2D) As Double
            Return (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X)
        End Function

        Private Shared Function OppositeSigns(first As Double, second As Double) As Boolean
            Return first > CoordinateTolerance AndAlso second < -CoordinateTolerance OrElse
                   first < -CoordinateTolerance AndAlso second > CoordinateTolerance
        End Function

        Private Shared Function OnSegment(a As Vector2D, b As Vector2D, pointValue As Vector2D) As Boolean
            Return pointValue.X >= Math.Min(a.X, b.X) - CoordinateTolerance AndAlso
                   pointValue.X <= Math.Max(a.X, b.X) + CoordinateTolerance AndAlso
                   pointValue.Y >= Math.Min(a.Y, b.Y) - CoordinateTolerance AndAlso
                   pointValue.Y <= Math.Max(a.Y, b.Y) + CoordinateTolerance
        End Function

        Private Shared Function HasVerticalEdge(closedRing As Vector2D(),
                                                x As Double,
                                                minimumY As Double,
                                                maximumY As Double) As Boolean
            For index As Integer = 0 To closedRing.Length - 2
                Dim first As Vector2D = closedRing(index)
                Dim second As Vector2D = closedRing(index + 1)
                If Math.Abs(first.X - x) <= CoordinateTolerance AndAlso
                   Math.Abs(second.X - x) <= CoordinateTolerance AndAlso
                   Math.Abs(Math.Min(first.Y, second.Y) - minimumY) <= CoordinateTolerance AndAlso
                   Math.Abs(Math.Max(first.Y, second.Y) - maximumY) <= CoordinateTolerance Then
                    Return True
                End If
            Next
            Return False
        End Function

        Private Shared Function Rectangle(minimumX As Double,
                                          maximumX As Double,
                                          minimumY As Double,
                                          maximumY As Double) As Vector2D()
            Return {
                Point(minimumX, minimumY), Point(maximumX, minimumY),
                Point(maximumX, maximumY), Point(minimumX, maximumY)
            }
        End Function

        Private Shared Function CloseRing(openRing As Vector2D()) As Vector2D()
            Dim result(openRing.Length) As Vector2D
            Array.Copy(openRing, result, openRing.Length)
            result(result.Length - 1) = openRing(0)
            Return result
        End Function

        Private Shared Function Transform(points As IEnumerable(Of Vector2D),
                                          angle As Double,
                                          offset As Vector2D) As Vector2D()
            Return points.Select(Function(pointValue) Transform(pointValue, angle, offset)).ToArray()
        End Function

        Private Shared Function Transform(pointValue As Vector2D,
                                          angle As Double,
                                          offset As Vector2D) As Vector2D
            Dim cosine As Double = Math.Cos(angle)
            Dim sine As Double = Math.Sin(angle)
            Return New Vector2D(
                offset.X + pointValue.X * cosine - pointValue.Y * sine,
                offset.Y + pointValue.X * sine + pointValue.Y * cosine)
        End Function

        Private Shared Function Point(x As Double, y As Double) As Vector2D
            Return New Vector2D(x, y)
        End Function

        Private Shared Function AreSamePoint(first As Vector2D,
                                             second As Vector2D,
                                             Optional tolerance As Double = CoordinateTolerance) As Boolean
            Return Math.Abs(first.X - second.X) <= tolerance AndAlso
                   Math.Abs(first.Y - second.Y) <= tolerance
        End Function

        Private Shared Function IsFinite(value As Double) As Boolean
            Return Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value)
        End Function

        Private Shared Function MainType(fullName As String) As Type
            Dim result As Type = MainAssembly().GetType(fullName, throwOnError:=False)
            Assert.That(result, [Is].Not.Null, fullName & " must exist in the production assembly.")
            Return result
        End Function

        Private Shared Function MainAssembly() As Assembly
            If _mainAssembly IsNot Nothing Then Return _mainAssembly

            Dim root As String = RepositoryRoot()
            Dim candidates As String() = {
                Path.Combine(root, "bin", "Debug", "CivilEnginStructures.dll"),
                Path.Combine(root, "bin", "Release", "CivilEnginStructures.dll")
            }
            Dim assemblyPath As String = candidates.
                Where(Function(candidate) File.Exists(candidate)).
                OrderByDescending(Function(candidate) File.GetLastWriteTimeUtc(candidate)).
                FirstOrDefault()
            Assert.That(assemblyPath, [Is].Not.Null,
                        "Build CivilEnginStructures.vbproj before running beam plan geometry tests.")
            _mainAssembly = Assembly.LoadFrom(assemblyPath)
            Return _mainAssembly
        End Function

        Private Shared Function RepositoryRoot() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.TestDirectory)
            While directory IsNot Nothing
                If File.Exists(Path.Combine(directory.FullName, "CivilEnginStructures.vbproj")) Then
                    Return directory.FullName
                End If
                directory = directory.Parent
            End While
            Assert.Fail("Could not locate the repository root from the test directory.")
            Return Nothing
        End Function
    End Class
End Namespace
