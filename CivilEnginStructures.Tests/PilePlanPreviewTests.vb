Imports System
Imports System.Drawing
Imports System.Linq
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class PilePlanPreviewTests
        Private Const GrillageMissingStatus As String = "Ростверк отсутствует — размерные цепи до граней недоступны"

        <Test>
        Public Sub SevenByThreeSchemeUsesActualPileAxesAndReportsMillimetres()
            Dim plan As Object = CreatePlan(RectangleCorners(4.1R, 1.5R), RegularPiles(7, 3, 0.25R, 0.25R, 0.6R, 0.5R))

            Assert.Multiple(
                Sub()
                    Assert.That(IntegerProperty(plan, "PileCount"), [Is].EqualTo(21))
                    AssertSegments(plan, "HorizontalSegmentsMillimetres", 250, 600, 600, 600, 600, 600, 600, 250)
                    AssertSegments(plan, "VerticalSegmentsMillimetres", 250, 500, 500, 250)
                    Assert.That(StringProperty(plan, "StatusText"), [Is].Empty)
                End Sub)
        End Sub

        <Test>
        Public Sub TranslationAndRotationDoNotChangeDimensionChains()
            Dim corners As Double(,) = RectangleCorners(4.1R, 1.5R)
            Dim piles As Double(,) = RegularPiles(7, 3, 0.25R, 0.25R, 0.6R, 0.5R)
            Dim transformedCorners As Double(,) = TransformPoints(corners, 37.0R, 123.4R, -57.8R, 2)
            Dim transformedPiles As Double(,) = TransformPoints(piles, 37.0R, 123.4R, -57.8R, 4)

            Dim plan As Object = CreatePlan(transformedCorners, transformedPiles)

            Assert.Multiple(
                Sub()
                    AssertSegments(plan, "HorizontalSegmentsMillimetres", 250, 600, 600, 600, 600, 600, 600, 250)
                    AssertSegments(plan, "VerticalSegmentsMillimetres", 250, 500, 500, 250)
                    Assert.That(IntegerProperty(plan, "PileCount"), [Is].EqualTo(21))
                End Sub)
        End Sub

        <Test>
        Public Sub StaggeredMissingAndDuplicateStationsUseSortedActualAxes()
            Dim piles As Double(,) = {
                {0.3R, 0.2R, 0.5R, 0.0R},
                {1.1R, 1.7R, 0.0R, 0.45R},
                {2.8R, 2.6R, 0.5R, 0.0R},
                {4.4R, 0.2R, 0.5R, 0.0R},
                {1.1R, 2.6R, 0.5R, 0.0R},
                {1.1004R, 1.7R, 0.5R, 0.0R}
            }

            Dim plan As Object = CreatePlan(RectangleCorners(5.0R, 3.0R), piles)

            Assert.Multiple(
                Sub()
                    Assert.That(IntegerProperty(plan, "PileCount"), [Is].EqualTo(6))
                    AssertSegments(plan, "VerticalSegmentsMillimetres", 200, 1500, 900, 400)
                    Dim horizontal As Double() = DoubleArrayProperty(plan, "HorizontalSegmentsMillimetres")
                    Assert.That(horizontal.Length, [Is].EqualTo(5),
                                "Axes closer than 0.5 mm must not create a spurious dimension segment.")
                    Assert.That(horizontal.Sum(), [Is].EqualTo(5000.0R).Within(0.01R))
                    Assert.That(horizontal.All(Function(value) value > 0.0R), [Is].True)
                End Sub)
        End Sub

        <Test>
        Public Sub MissingOrInvalidGrillageReturnsAnInformativeSafePlan()
            Dim onePile As Double(,) = {{12.0R, -4.0R, 0.5R, 0.0R}}
            Dim invalidCorners As Double(,) = {{0.0R, 0.0R}, {4.0R, 0.0R}, {4.0R, 1.5R}}
            Dim missingPlan As Object = Nothing
            Dim invalidPlan As Object = Nothing

            Assert.DoesNotThrow(Sub() missingPlan = CreatePlan(Nothing, onePile))
            Assert.DoesNotThrow(Sub() invalidPlan = CreatePlan(invalidCorners, onePile))
            AssertInvalidGrillagePlan(missingPlan)
            AssertInvalidGrillagePlan(invalidPlan)
        End Sub

        <Test>
        Public Sub EmptyPileSetKeepsTheGrillageSafeWithoutInventingStations()
            Dim plan As Object = Nothing

            Assert.DoesNotThrow(Sub() plan = CreatePlan(RectangleCorners(4.1R, 1.5R), Nothing))
            Assert.Multiple(
                Sub()
                    Assert.That(IntegerProperty(plan, "PileCount"), [Is].Zero)
                    Assert.That(DoubleArrayProperty(plan, "HorizontalSegmentsMillimetres"), [Is].Empty)
                    Assert.That(DoubleArrayProperty(plan, "VerticalSegmentsMillimetres"), [Is].Empty)
                    Assert.That(StringProperty(plan, "StatusText"), [Is].Not.Empty)
                End Sub)
        End Sub

        <Test>
        Public Sub RendererDrawsRoundSquareAndZeroSizePileCentresOnWhiteCanvas()
            Dim piles As Double(,) = {
                {0.5R, 0.5R, 0.5R, 0.0R},
                {2.0R, 0.75R, 0.0R, 0.4R},
                {3.6R, 1.1R, 0.0R, 0.0R}
            }
            Dim plan As Object = CreatePlan(RectangleCorners(4.1R, 1.5R), piles)

            Using normal As Bitmap = Render(plan, New Size(800, 500)),
                  narrow As Bitmap = Render(plan, New Size(72, 240))
                Assert.Multiple(
                    Sub()
                        Assert.That(IntegerProperty(plan, "PileCount"), [Is].EqualTo(3),
                                    "A zero-size pile still has an actual centre and must not be dropped.")
                        AssertBitmapHasWhiteCanvasAndInk(normal, New Size(800, 500))
                        AssertBitmapHasWhiteCanvasAndInk(narrow, New Size(72, 240))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub RendererReturnsDisposableNonZeroBitmapForEmptyAndZeroSizeRequests()
            Dim plan As Object = CreatePlan(Nothing, Nothing)
            Dim rendered As Bitmap = Nothing

            Assert.DoesNotThrow(Sub() rendered = Render(plan, Size.Empty))
            Assert.That(rendered, [Is].Not.Null)
            Try
                Assert.Multiple(
                    Sub()
                        Assert.That(rendered.Width, [Is].GreaterThan(0))
                        Assert.That(rendered.Height, [Is].GreaterThan(0))
                        Assert.That(ContainsNonWhitePixel(rendered), [Is].True,
                                    "The empty state must be informative rather than a blank bitmap.")
                    End Sub)
            Finally
                rendered.Dispose()
            End Try
            Assert.Throws(Of ArgumentException)(Sub() rendered.GetPixel(0, 0),
                                                 "The returned bitmap belongs to the caller and can be disposed safely.")
        End Sub

        Private Shared Function CreatePlan(corners As Double(,), piles As Double(,)) As Object
            Dim geometryType As Type = RequireType("CivilEnginStructures.PilePlanGeometry")
            Dim factory As MethodInfo = geometryType.GetMethod("Create", BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
            If factory Is Nothing Then Throw New AssertionException("PilePlanGeometry.Create(Double(,), Double(,)) is missing.")
            Return Invoke(factory, Nothing, corners, piles)
        End Function

        Private Shared Function Render(plan As Object, requestedSize As Size) As Bitmap
            Dim rendererType As Type = RequireType("CivilEnginStructures.PilePlanPreview")
            Dim method As MethodInfo = rendererType.GetMethod("Render", BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
            If method Is Nothing Then Throw New AssertionException("PilePlanPreview.Render(PilePlanGeometry, Size) is missing.")
            Return DirectCast(Invoke(method, Nothing, plan, requestedSize), Bitmap)
        End Function

        Private Shared Function RequireType(fullName As String) As Type
            Dim target As Type = GetType(PilePlanPreviewTests).Assembly.GetType(fullName, False)
            If target Is Nothing Then Throw New AssertionException(fullName & " has not been implemented yet.")
            Return target
        End Function

        Private Shared Function Invoke(method As MethodInfo, target As Object, ParamArray arguments As Object()) As Object
            Try
                Return method.Invoke(target, arguments)
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Function

        Private Shared Function DoubleArrayProperty(instance As Object, name As String) As Double()
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(name, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            If propertyInfo Is Nothing Then Throw New AssertionException(name & " property is missing.")
            Return DirectCast(propertyInfo.GetValue(instance, Nothing), Double())
        End Function

        Private Shared Function IntegerProperty(instance As Object, name As String) As Integer
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(name, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            If propertyInfo Is Nothing Then Throw New AssertionException(name & " property is missing.")
            Return CInt(propertyInfo.GetValue(instance, Nothing))
        End Function

        Private Shared Function StringProperty(instance As Object, name As String) As String
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(name, BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            If propertyInfo Is Nothing Then Throw New AssertionException(name & " property is missing.")
            Return DirectCast(propertyInfo.GetValue(instance, Nothing), String)
        End Function

        Private Shared Sub AssertSegments(plan As Object, propertyName As String, ParamArray expected As Double())
            Dim actual As Double() = DoubleArrayProperty(plan, propertyName)
            Assert.That(actual.Length, [Is].EqualTo(expected.Length), propertyName & " count")
            For index As Integer = 0 To expected.Length - 1
                Assert.That(actual(index), [Is].EqualTo(expected(index)).Within(0.01R),
                            propertyName & " segment " & index.ToString())
            Next
        End Sub

        Private Shared Sub AssertInvalidGrillagePlan(plan As Object)
            Assert.Multiple(
                Sub()
                    Assert.That(DoubleArrayProperty(plan, "HorizontalSegmentsMillimetres"), [Is].Empty)
                    Assert.That(DoubleArrayProperty(plan, "VerticalSegmentsMillimetres"), [Is].Empty)
                    Assert.That(StringProperty(plan, "StatusText"), [Is].EqualTo(GrillageMissingStatus))
                End Sub)
        End Sub

        Private Shared Sub AssertBitmapHasWhiteCanvasAndInk(bitmap As Bitmap, requestedSize As Size)
            Assert.That(bitmap, [Is].Not.Null)
            Assert.That(bitmap.Size, [Is].EqualTo(requestedSize))
            Assert.That(ContainsPixel(bitmap, Function(color) color.ToArgb() = Color.White.ToArgb()), [Is].True,
                        "The preview canvas must contain a white background.")
            Assert.That(ContainsNonWhitePixel(bitmap), [Is].True,
                        "The preview must contain visible geometry, axes, dimensions, or labels.")
        End Sub

        Private Shared Function ContainsNonWhitePixel(bitmap As Bitmap) As Boolean
            Return ContainsPixel(bitmap, Function(color) color.ToArgb() <> Color.White.ToArgb())
        End Function

        Private Shared Function ContainsPixel(bitmap As Bitmap, predicate As Func(Of Color, Boolean)) As Boolean
            For y As Integer = 0 To bitmap.Height - 1
                For x As Integer = 0 To bitmap.Width - 1
                    If predicate(bitmap.GetPixel(x, y)) Then Return True
                Next
            Next
            Return False
        End Function

        Private Shared Function RectangleCorners(length As Double, width As Double) As Double(,)
            Return New Double(,) {{0.0R, 0.0R}, {length, 0.0R}, {length, width}, {0.0R, width}}
        End Function

        Private Shared Function RegularPiles(columns As Integer, rows As Integer,
                                             firstX As Double, firstY As Double,
                                             stepX As Double, stepY As Double) As Double(,)
            Dim result(columns * rows - 1, 3) As Double
            Dim index As Integer = 0
            For row As Integer = 0 To rows - 1
                For column As Integer = 0 To columns - 1
                    result(index, 0) = firstX + column * stepX
                    result(index, 1) = firstY + row * stepY
                    result(index, 2) = 0.5R
                    result(index, 3) = 0.0R
                    index += 1
                Next
            Next
            Return result
        End Function

        Private Shared Function TransformPoints(source As Double(,), angleDegrees As Double,
                                                translateX As Double, translateY As Double,
                                                columns As Integer) As Double(,)
            Dim rows As Integer = source.GetLength(0)
            Dim result(rows - 1, columns - 1) As Double
            Dim angle As Double = angleDegrees * Math.PI / 180.0R
            Dim cosine As Double = Math.Cos(angle)
            Dim sine As Double = Math.Sin(angle)
            For row As Integer = 0 To rows - 1
                result(row, 0) = translateX + source(row, 0) * cosine - source(row, 1) * sine
                result(row, 1) = translateY + source(row, 0) * sine + source(row, 1) * cosine
                For column As Integer = 2 To columns - 1
                    result(row, column) = source(row, column)
                Next
            Next
            Return result
        End Function
    End Class
End Namespace
