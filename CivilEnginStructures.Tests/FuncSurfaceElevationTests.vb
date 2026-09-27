Imports System.Collections.Generic
Imports System.IO
Imports System.Text.RegularExpressions
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class FuncSurfaceElevationTests
        Private Const Tolerance As Double = 0.000000001
        Private Const FuncSurfaceRelativePath As String = "FuncSurface\FuncSurface.vb"
        Private Const RegexOptionsValue As RegexOptions = RegexOptions.IgnoreCase Or
                                                           RegexOptions.CultureInvariant Or
                                                           RegexOptions.Singleline

        <Test>
        Public Sub GetElevationUsesNullableValueWhenSurfaceReturnsElevation()
            Dim methodSource As String = ReadGetElevationToSurfaceSource()
            Dim nullableAssignmentPattern As String =
                "Dim\s+(?<value>[A-Za-z_]\w*)\s+As\s+Nullable\s*\(\s*Of\s+Double\s*\)\s*=\s*" &
                "surface\s*\.\s*GetElevation\s*\(\s*point\s*\)"
            Dim assignment As Match = Regex.Match(methodSource,
                                                   nullableAssignmentPattern,
                                                   RegexOptionsValue)

            Assert.That(assignment.Success, [Is].True,
                        "Surface.GetElevation returns Nullable(Of Double) and must be captured without coercing Nothing to zero.")
            If Not assignment.Success Then Return

            Dim variableName As String = Regex.Escape(assignment.Groups("value").Value)

            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(methodSource,
                                              variableName & "\s*\.\s*HasValue",
                                              RegexOptionsValue),
                                [Is].True,
                                "The normal Surface result must be accepted only when Nullable.HasValue is true.")
                    Assert.That(Regex.IsMatch(methodSource,
                                              "elevation\s*=\s*" & variableName & "\s*\.\s*Value",
                                              RegexOptionsValue),
                                [Is].True,
                                "The successful normal path must return the Nullable.Value elevation.")
                End Sub)
        End Sub

        <Test>
        Public Sub TriangleInteriorUsesLinearInterpolationInXY()
            Dim elevation As Double = -999.0
            Dim squaredDistance As Double = -1.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(2, 3),
                New Vector3D(0, 0, 10),
                New Vector3D(10, 0, 20),
                New Vector3D(0, 10, 30),
                False,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(elevation, [Is].EqualTo(18.0).Within(Tolerance))
                    Assert.That(squaredDistance, [Is].EqualTo(0.0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub TriangleBoundaryCountsAsInside()
            Dim elevation As Double = -999.0
            Dim squaredDistance As Double = -1.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(7.5, 2.5),
                New Vector3D(0, 0, 10),
                New Vector3D(10, 0, 20),
                New Vector3D(0, 10, 30),
                False,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(elevation, [Is].EqualTo(22.5).Within(Tolerance))
                    Assert.That(squaredDistance, [Is].EqualTo(0.0).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub PointOutsideTriangleUsesClosestPointOnTriangleAndInterpolatesItsElevation()
            Dim elevation As Double = -999.0
            Dim squaredDistance As Double = -1.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(8, 8),
                New Vector3D(0, 0, 10),
                New Vector3D(10, 0, 20),
                New Vector3D(0, 10, 30),
                True,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(elevation, [Is].EqualTo(25.0).Within(Tolerance),
                                "The closest XY point is (5, 5), whose linearly interpolated elevation is 25.")
                    Assert.That(squaredDistance, [Is].EqualTo(18.0).Within(Tolerance),
                                "Squared XY distance from (8, 8) to (5, 5) is 18.")
                End Sub)
        End Sub

        <Test>
        Public Sub PointOutsideTriangleIsRejectedWhenClampingIsDisabled()
            Dim elevation As Double = 123.0
            Dim squaredDistance As Double = 456.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(8, 8),
                New Vector3D(0, 0, 10),
                New Vector3D(10, 0, 20),
                New Vector3D(0, 10, 30),
                False,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    Assert.That(elevation, [Is].EqualTo(123.0),
                                "A failed calculation must not overwrite the caller's elevation.")
                    Assert.That(squaredDistance, [Is].EqualTo(456.0),
                                "A failed calculation must not overwrite the caller's distance.")
                End Sub)
        End Sub

        <Test>
        Public Sub DegenerateTriangleIsRejectedWithoutChangingOutputs()
            Dim elevation As Double = 123.0
            Dim squaredDistance As Double = 456.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(1, 1),
                New Vector3D(0, 0, 10),
                New Vector3D(1, 1, 20),
                New Vector3D(2, 2, 30),
                True,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    Assert.That(elevation, [Is].EqualTo(123.0))
                    Assert.That(squaredDistance, [Is].EqualTo(456.0))
                End Sub)
        End Sub

        <Test>
        Public Sub TriangleContainingNonFiniteCoordinateIsRejected()
            Dim elevation As Double = 123.0
            Dim squaredDistance As Double = 456.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetElevationOnTriangle(
                New Vector2D(1, 1),
                New Vector3D(0, 0, 10),
                New Vector3D(Double.NaN, 0, 20),
                New Vector3D(0, 10, 30),
                True,
                elevation,
                squaredDistance)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    Assert.That(elevation, [Is].EqualTo(123.0))
                    Assert.That(squaredDistance, [Is].EqualTo(456.0))
                End Sub)
        End Sub

        <Test>
        Public Sub NearestSurfacePointSkipsNonFinitePointsAndKeepsFirstAtEqualDistance()
            Dim points As New List(Of Vector3D) From {
                New Vector3D(Double.NaN, 4, 999),
                New Vector3D(3, 4, 50),
                New Vector3D(5, 4, 70),
                New Vector3D(4, 8, 90)
            }
            Dim elevation As Double = -999.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetNearestSurfacePointElevation(
                New Vector2D(4, 4),
                points,
                elevation)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].True)
                    Assert.That(elevation, [Is].EqualTo(50.0),
                                "The first of the two equally near valid points must win.")
                End Sub)
        End Sub

        <Test>
        Public Sub EmptySurfacePointSetReturnsFalseWithoutChangingElevation()
            Dim elevation As Double = 321.0

            Dim result As Boolean = SurfaceElevationGeometry.TryGetNearestSurfacePointElevation(
                New Vector2D(4, 4),
                New List(Of Vector3D)(),
                elevation)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].False)
                    Assert.That(elevation, [Is].EqualTo(321.0))
                End Sub)
        End Sub

        Private Shared Function ReadGetElevationToSurfaceSource() As String
            Dim source As String = ReadFuncSurfaceSource()
            Dim methodMatch As Match = Regex.Match(
                source,
                "Public\s+Shared\s+Function\s+getElevationToSurface\b(?<method>.*?)End\s+Function",
                RegexOptionsValue)

            Assert.That(methodMatch.Success, [Is].True,
                        "Could not find FuncSurface.getElevationToSurface in the active checkout.")
            If Not methodMatch.Success Then Return String.Empty
            Return "Public Shared Function getElevationToSurface" & methodMatch.Groups("method").Value
        End Function

        Private Shared Function ReadFuncSurfaceSource() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.WorkDirectory)

            While directory IsNot Nothing
                Dim candidate As String = Path.Combine(directory.FullName, FuncSurfaceRelativePath)
                If File.Exists(candidate) Then Return File.ReadAllText(candidate)
                directory = directory.Parent
            End While

            Assert.Fail("Could not locate " & FuncSurfaceRelativePath &
                        " above test work directory " & TestContext.CurrentContext.WorkDirectory & ".")
            Return String.Empty
        End Function
    End Class
End Namespace
