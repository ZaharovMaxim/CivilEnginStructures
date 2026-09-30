Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Drawing
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Stg

Namespace Tests
    <TestFixture>
    Public Class BridgeBeamPlanEntityTests
        Private Const Tolerance As Double = 0.000000001R
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub SetOutlineCopiesInputAndOutlineReturnsAnIndependentSnapshot()
            Dim entity As DwgEntity = NewPlanEntity()
            Dim input As Vector2D() = RectangleOutline()

            SetOutline(entity, input)
            input(0) = New Vector2D(100.0R, 200.0R)
            Dim firstSnapshot As Vector2D() = Outline(entity)
            firstSnapshot(1) = New Vector2D(-100.0R, -200.0R)
            Dim secondSnapshot As Vector2D() = Outline(entity)

            Assert.Multiple(
                Sub()
                    AssertPoint(secondSnapshot(0), 0.0R, 0.0R)
                    AssertPoint(secondSnapshot(1), 4.0R, 0.0R)
                    Assert.That(secondSnapshot, [Is].Not.SameAs(firstSnapshot))
                End Sub)
        End Sub

        <Test>
        Public Sub SetOutlineRejectsNullOpenShortAndNonFiniteContours()
            Dim entity As DwgEntity = NewPlanEntity()

            AssertInvocationInner(Of ArgumentNullException)(Sub() SetOutline(entity, Nothing))
            AssertInvocationInner(Of ArgumentException)(
                Sub() SetOutline(entity, {New Vector2D(0, 0), New Vector2D(1, 0), New Vector2D(0, 0)}))
            AssertInvocationInner(Of ArgumentException)(
                Sub() SetOutline(entity, {New Vector2D(0, 0), New Vector2D(1, 0),
                                          New Vector2D(1, 1), New Vector2D(0, 1)}))

            Dim invalidCoordinates As Double() = {
                Double.NaN,
                Double.PositiveInfinity,
                Double.NegativeInfinity
            }
            For Each invalidCoordinate As Double In invalidCoordinates
                Dim invalidX As Vector2D() = RectangleOutline()
                invalidX(1) = New Vector2D(invalidCoordinate, invalidX(1).Y)
                AssertInvocationInner(Of ArgumentException)(Sub() SetOutline(entity, invalidX))

                Dim invalidY As Vector2D() = RectangleOutline()
                invalidY(2) = New Vector2D(invalidY(2).X, invalidCoordinate)
                AssertInvocationInner(Of ArgumentException)(Sub() SetOutline(entity, invalidY))
            Next
        End Sub

        <Test>
        Public Sub HatchPropertiesNormalizeSupportedPatternsAndRejectInvalidNumbers()
            Dim entity As DwgEntity = NewPlanEntity()

            SetProperty(entity, "HatchPatternName", " ansi31 ")
            SetProperty(entity, "HatchScale", 0.000001R)
            SetProperty(entity, "HatchAngle", -Math.PI / 3.0R)

            Assert.Multiple(
                Sub()
                    Assert.That(PropertyValue(Of String)(entity, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                    Assert.That(PropertyValue(Of Double)(entity, "HatchScale"), [Is].EqualTo(0.000001R))
                    Assert.That(PropertyValue(Of Double)(entity, "HatchAngle"), [Is].EqualTo(-Math.PI / 3.0R))
                End Sub)

            SetProperty(entity, "HatchPatternName", "solid")
            Assert.That(PropertyValue(Of String)(entity, "HatchPatternName"), [Is].EqualTo("SOLID"))

            For Each invalidPattern As String In {String.Empty, "   ", "AR-CONC"}
                AssertInvocationInner(Of ArgumentException)(
                    Sub() SetProperty(entity, "HatchPatternName", invalidPattern))
            Next
            For Each invalidScale As Double In {
                0.0R, -1.0R, Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity
            }
                AssertInvocationInner(Of ArgumentOutOfRangeException)(
                    Sub() SetProperty(entity, "HatchScale", invalidScale))
            Next
            For Each invalidAngle As Double In {
                Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity
            }
                AssertInvocationInner(Of ArgumentOutOfRangeException)(
                    Sub() SetProperty(entity, "HatchAngle", invalidAngle))
            Next
        End Sub

        <Test>
        Public Sub ConcaveSteppedOutlineHitTestsRespectTheGapInsideItsBounds()
            Dim entity As DwgEntity = NewPlanEntity()
            SetOutline(entity, SteppedOutline())

            Assert.Multiple(
                Sub()
                    Assert.That(ContainsPoint(entity, New Vector2D(0.5R, 3.0R)), [Is].True)
                    Assert.That(ContainsPoint(entity, New Vector2D(2.0R, 0.5R)), [Is].True)
                    Assert.That(ContainsPoint(entity, New Vector2D(1.0R, 2.0R)), [Is].True,
                                "A point on the step boundary belongs to the contour.")
                    Assert.That(ContainsPoint(entity, New Vector2D(2.0R, 2.0R)), [Is].False,
                                "The empty step gap is outside even though it lies inside Bounds.")
                    Assert.That(ContainsPoint(entity, New Vector2D(5.0R, 5.0R)), [Is].False)
                End Sub)

            Assert.Multiple(
                Sub()
                    Assert.That(entity.IntersectWith(Box(0.2R, 2.0R, 0.8R, 3.0R), 1.0R), [Is].True)
                    Assert.That(entity.IntersectWith(Box(0.8R, 1.5R, 1.2R, 2.5R), 1.0R), [Is].True)
                    Assert.That(entity.IntersectWith(Box(-1.0R, -1.0R, 5.0R, 5.0R), 1.0R), [Is].True)
                    Assert.That(entity.IntersectWith(Box(1.5R, 1.5R, 2.5R, 2.5R), 1.0R), [Is].False,
                                "Bounds overlap alone must not select the empty step gap.")
                    Assert.That(entity.IntersectWith(Box(5.0R, 5.0R, 6.0R, 6.0R), 1.0R), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub TransformUpdatesEveryOutlinePointAndRecalculatesBounds()
            Dim entity As DwgEntity = NewPlanEntity()
            Dim original As Vector2D() = RectangleOutline()
            SetOutline(entity, original)
            Dim transformValue As Matrix = Matrix.CreateTranslation(10.0R, -3.0R, 0.0R) *
                                           Matrix.CreateScale(2.0R, 0.5R, 1.0R)
            Dim expected As Vector2D() = original.
                Select(Function(pointValue) Vector2D.Transform(pointValue, transformValue)).
                ToArray()

            entity.Transform(transformValue)

            Dim actual As Vector2D() = Outline(entity)
            For index As Integer = 0 To expected.Length - 1
                AssertPoint(actual(index), expected(index).X, expected(index).Y)
            Next
            Assert.Multiple(
                Sub()
                    Assert.That(entity.Bounds.Left, [Is].EqualTo(expected.Min(Function(p) p.X)).Within(Tolerance))
                    Assert.That(entity.Bounds.Right, [Is].EqualTo(expected.Max(Function(p) p.X)).Within(Tolerance))
                    Assert.That(entity.Bounds.Bottom, [Is].EqualTo(expected.Min(Function(p) p.Y)).Within(Tolerance))
                    Assert.That(entity.Bounds.Top, [Is].EqualTo(expected.Max(Function(p) p.Y)).Within(Tolerance))
                    AssertPoint(actual(actual.Length - 1), actual(0).X, actual(0).Y)
                End Sub)
        End Sub

        <Test>
        Public Sub AssignCopiesOwnStateWithoutAliasingTheSource()
            Dim source As DwgEntity = NewPlanEntity()
            SetOutline(source, RectangleOutline())
            SetProperty(source, "FillColor", New CadColor(Color.FromArgb(30, 80, 130)))
            SetProperty(source, "HatchPatternName", "ANSI31")
            SetProperty(source, "HatchScale", 2.25R)
            SetProperty(source, "HatchAngle", 0.75R)
            Dim assigned As DwgEntity = NewPlanEntity()

            assigned.Assign(source)
            assigned.Transform(Matrix.CreateTranslation(100.0R, 0.0R, 0.0R))
            SetProperty(assigned, "HatchScale", 4.5R)

            Assert.Multiple(
                Sub()
                    AssertPoint(Outline(source)(0), 0.0R, 0.0R)
                    AssertPoint(Outline(assigned)(0), 100.0R, 0.0R)
                    Assert.That(PropertyValue(Of Double)(source, "HatchScale"), [Is].EqualTo(2.25R))
                    Assert.That(PropertyValue(Of String)(assigned, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                    Assert.That(PropertyValue(Of Double)(assigned, "HatchAngle"), [Is].EqualTo(0.75R))
                    Assert.That(PropertyValue(Of CadColor)(assigned, "FillColor"),
                                [Is].EqualTo(PropertyValue(Of CadColor)(source, "FillColor")))
                End Sub)

            Assert.Throws(Of ArgumentException)(Sub() assigned.Assign(New DwgLine()))
        End Sub

        <Test>
        Public Sub OwnStateBinaryRoundTripPreservesOutlineFillHatchAndBounds()
            Dim original As DwgEntity = NewPlanEntity()
            SetOutline(original, SteppedOutline())
            SetProperty(original, "FillColor", New CadColor(Color.FromArgb(45, 90, 135)))
            SetProperty(original, "HatchPatternName", "ANSI31")
            SetProperty(original, "HatchScale", 3.5R)
            SetProperty(original, "HatchAngle", -0.25R)
            Dim sourceDocument As New StgDocument()
            InvokeOwnState(original, "OnSaveToStg", sourceDocument.Body)

            Dim loadedDocument As New StgDocument()
            Using stream As New MemoryStream()
                sourceDocument.SaveToStreamAsBinary(stream)
                stream.Position = 0
                loadedDocument.LoadFromStreamAsBinary(stream)
            End Using
            Dim restored As DwgEntity = NewPlanEntity()
            InvokeOwnState(restored, "OnLoadFromStg", loadedDocument.Body)

            AssertOutlineEqual(Outline(restored), Outline(original))
            Assert.Multiple(
                Sub()
                    Assert.That(PropertyValue(Of CadColor)(restored, "FillColor"),
                                [Is].EqualTo(PropertyValue(Of CadColor)(original, "FillColor")))
                    Assert.That(PropertyValue(Of String)(restored, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                    Assert.That(PropertyValue(Of Double)(restored, "HatchScale"), [Is].EqualTo(3.5R))
                    Assert.That(PropertyValue(Of Double)(restored, "HatchAngle"), [Is].EqualTo(-0.25R))
                    Assert.That(restored.Bounds, [Is].EqualTo(original.Bounds))
                End Sub)
        End Sub

        <Test>
        Public Sub DefaultEntityUsesSolidFillAndHasPlanSelectionMetadata()
            Dim entity As DwgEntity = NewPlanEntity()
            Dim controller As Object = Activator.CreateInstance(
                MainType("CivilEnginStructures.BridgeBeamPlanEntityController"))

            Assert.Multiple(
                Sub()
                    Assert.That(PropertyValue(Of String)(entity, "HatchPatternName"), [Is].EqualTo("SOLID"))
                    Assert.That(entity.EntityName, [Is].EqualTo("Контур балки моста"))
                    Assert.That(entity.IsBackgroud, [Is].True)
                    Assert.That(PropertyValue(Of Boolean)(controller, "SupportPaint3d"), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub DetachedLayoutBuildsFillAndPerimeterAndHighlightDoesNotMutateStoredStyle()
            Dim entity As DwgEntity = NewPlanEntity()
            SetOutline(entity, RectangleOutline())
            entity.Color = New CadColor(Color.FromArgb(20, 40, 60))
            SetProperty(entity, "FillColor", New CadColor(Color.FromArgb(80, 100, 120)))
            SetProperty(entity, "HatchPatternName", "ANSI31")
            SetProperty(entity, "HatchScale", 2.0R)
            SetProperty(entity, "HatchAngle", 0.5R)

            Dim normal As List(Of DwgEntity) = CreateLayout(entity, False)
            Dim highlighted As List(Of DwgEntity) = CreateLayout(entity, True)
            Dim repeatedNormal As List(Of DwgEntity) = CreateLayout(entity, False)

            Assert.That(normal, Has.Count.EqualTo(2))
            Assert.That(normal(0), [Is].TypeOf(Of DwgHatch)())
            Assert.That(normal(1), [Is].TypeOf(Of DwgPolyline)())
            Dim normalHatch As DwgHatch = DirectCast(normal(0), DwgHatch)
            Dim normalPerimeter As DwgPolyline = DirectCast(normal(1), DwgPolyline)
            Dim highlightedHatch As DwgHatch = DirectCast(highlighted(0), DwgHatch)
            Dim highlightedPerimeter As DwgPolyline = DirectCast(highlighted(1), DwgPolyline)
            Dim repeatedHatch As DwgHatch = DirectCast(repeatedNormal(0), DwgHatch)
            Dim repeatedPerimeter As DwgPolyline = DirectCast(repeatedNormal(1), DwgPolyline)

            Assert.Multiple(
                Sub()
                    Assert.That(normalHatch.Color, [Is].EqualTo(PropertyValue(Of CadColor)(entity, "FillColor")))
                    Assert.That(normalPerimeter.Closed, [Is].True)
                    Assert.That(normalPerimeter.Color, [Is].EqualTo(entity.Color))
                    Assert.That(highlightedHatch.Color, [Is].Not.EqualTo(normalHatch.Color))
                    Assert.That(highlightedPerimeter.Color, [Is].Not.EqualTo(normalPerimeter.Color))
                    Assert.That(repeatedHatch.Color, [Is].EqualTo(normalHatch.Color))
                    Assert.That(repeatedPerimeter.Color, [Is].EqualTo(normalPerimeter.Color))
                    Assert.That(entity.Color, [Is].EqualTo(normalPerimeter.Color))
                    Assert.That(PropertyValue(Of CadColor)(entity, "FillColor"), [Is].EqualTo(normalHatch.Color))
                End Sub)
        End Sub

        Private Shared Function NewPlanEntity() As DwgEntity
            Return DirectCast(Activator.CreateInstance(MainType("CivilEnginStructures.BridgeBeamPlanEntity")), DwgEntity)
        End Function

        Private Shared Sub SetOutline(entity As DwgEntity, points As IEnumerable(Of Vector2D))
            entity.GetType().GetMethod("SetOutline", BindingFlags.Public Or BindingFlags.Instance).
                Invoke(entity, New Object() {points})
        End Sub

        Private Shared Function Outline(entity As DwgEntity) As Vector2D()
            Return PropertyValue(Of Vector2D())(entity, "Outline")
        End Function

        Private Shared Function ContainsPoint(entity As DwgEntity, pointValue As Vector2D) As Boolean
            Dim method As MethodInfo = entity.GetType().GetMethod(
                "ContainsPoint", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(method, [Is].Not.Null)
            Return CBool(method.Invoke(entity, New Object() {pointValue}))
        End Function

        Private Shared Function CreateLayout(entity As DwgEntity, highlight As Boolean) As List(Of DwgEntity)
            Dim method As MethodInfo = entity.GetType().GetMethod(
                "CreateLayoutEntities", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(method, [Is].Not.Null)
            Return DirectCast(method.Invoke(entity, New Object() {highlight}), List(Of DwgEntity))
        End Function

        Private Shared Sub InvokeOwnState(entity As DwgEntity, methodName As String, node As StgNode)
            Dim method As MethodInfo = entity.GetType().GetMethod(
                methodName, BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing, New Type() {GetType(StgNode)}, Nothing)
            Assert.That(method, [Is].Not.Null)
            method.Invoke(entity, New Object() {node})
        End Sub

        Private Shared Sub SetProperty(instance As Object, propertyName As String, value As Object)
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(
                propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null.And.Property("CanWrite").True)
            propertyInfo.SetValue(instance, value, Nothing)
        End Sub

        Private Shared Function PropertyValue(Of T)(instance As Object, propertyName As String) As T
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(
                propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null)
            Return DirectCast(propertyInfo.GetValue(instance, Nothing), T)
        End Function

        Private Shared Sub AssertInvocationInner(Of TException As Exception)(action As TestDelegate)
            Dim thrown As TargetInvocationException = Assert.Throws(Of TargetInvocationException)(action)
            Assert.That(thrown.InnerException, [Is].TypeOf(Of TException)())
        End Sub

        Private Shared Sub AssertOutlineEqual(actual As Vector2D(), expected As Vector2D())
            Assert.That(actual, Has.Length.EqualTo(expected.Length))
            For index As Integer = 0 To expected.Length - 1
                AssertPoint(actual(index), expected(index).X, expected(index).Y)
            Next
        End Sub

        Private Shared Sub AssertPoint(actual As Vector2D, expectedX As Double, expectedY As Double)
            Assert.Multiple(
                Sub()
                    Assert.That(actual.X, [Is].EqualTo(expectedX).Within(Tolerance))
                    Assert.That(actual.Y, [Is].EqualTo(expectedY).Within(Tolerance))
                End Sub)
        End Sub

        Private Shared Function RectangleOutline() As Vector2D()
            Return {
                New Vector2D(0.0R, 0.0R),
                New Vector2D(4.0R, 0.0R),
                New Vector2D(4.0R, 2.0R),
                New Vector2D(0.0R, 2.0R),
                New Vector2D(0.0R, 0.0R)
            }
        End Function

        Private Shared Function SteppedOutline() As Vector2D()
            Return {
                New Vector2D(0.0R, 0.0R),
                New Vector2D(4.0R, 0.0R),
                New Vector2D(4.0R, 1.0R),
                New Vector2D(1.0R, 1.0R),
                New Vector2D(1.0R, 4.0R),
                New Vector2D(0.0R, 4.0R),
                New Vector2D(0.0R, 0.0R)
            }
        End Function

        Private Shared Function Box(left As Double, bottom As Double,
                                    right As Double, top As Double) As BoundingBox2D
            Return New BoundingBox2D(New Vector2D(left, bottom), New Vector2D(right, top))
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
                        "Build CivilEnginStructures.vbproj before running bridge plan entity tests.")
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
