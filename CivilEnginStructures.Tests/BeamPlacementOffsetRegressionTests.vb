Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BeamPlacementOffsetRegressionTests
        Private Const GlobalHorizontalOffset As Double = 0.25R
        Private Const GlobalVerticalOffset As Double = 0.05R
        Private Const PavementThickness As Double = 0.1R
        Private Shared ReadOnly RowOffsets As Double() = {
            -9.0R, -6.99R, -4.98R, -2.97R, -0.96R,
            0.96R, 2.97R, 4.98R, 6.99R, 9.0R
        }
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub TrajectoryConstructorMapsSurfaceAndAlignmentOffsetsByTheirDeclaredMeaning()
            Dim trajectoryType As Type = MainType("CivilEnginStructures.TrajectoryPlacementBeams")
            Dim trajectoryKind As Type = trajectoryType.GetNestedType("TypeTrajectoryPlacementBeams")
            Dim noneValue As Object = [Enum].Parse(trajectoryKind, "None")

            Dim trajectory As Object = Activator.CreateInstance(
                trajectoryType,
                New Object() {7, 0.314R, -2.718R, noneValue})

            Assert.Multiple(
                Sub()
                    Assert.That(ReadInteger(trajectory, "NumberRows"), [Is].EqualTo(7))
                    Assert.That(ReadDouble(trajectory, "OffsetProjectSurface"),
                                [Is].EqualTo(0.314R).Within(0.000000001R))
                    Assert.That(ReadDouble(trajectory, "OffsetProjectAlignment"),
                                [Is].EqualTo(-2.718R).Within(0.000000001R))
                End Sub)
        End Sub

        <Test>
        Public Sub ApplyingTenRowTrajectoriesReplacesOffsetsAndSurvivesJsonRoundTrip()
            Dim beamType As Type = MainType("CivilEnginStructures.BeamI")

            For index As Integer = 0 To RowOffsets.Length - 1
                Dim rowNumber As Integer = index + 1
                Dim expectedAxisOffset As Double = GlobalHorizontalOffset + RowOffsets(index)
                Dim expectedSurfaceOffset As Double = GlobalVerticalOffset + PavementThickness
                Dim beam As Object = Activator.CreateInstance(beamType)
                WriteProperty(beam, "numberRow", rowNumber)
                WriteProperty(beam, "axisOffset", 123.456R)
                WriteProperty(beam, "offsetSurface", 654.321R)
                WriteProperty(beam, "height", 2.34R)
                WriteProperty(beam, "lenght", 33.4R)
                Dim trajectory As Object = CreateTrajectory(
                    rowNumber, expectedSurfaceOffset, expectedAxisOffset)

                ApplyTrajectoryOffsets(beam, trajectory)
                ApplyTrajectoryOffsets(beam, trajectory)

                Dim json As String = JsonConvert.SerializeObject(beam)
                Dim restored As Object = JsonConvert.DeserializeObject(json, beamType)

                Assert.Multiple(
                    Sub()
                        Assert.That(ReadDouble(beam, "axisOffset"),
                                    [Is].EqualTo(expectedAxisOffset).Within(0.000000001R),
                                    "row " & rowNumber & " must replace the previous horizontal offset")
                        Assert.That(ReadDouble(beam, "offsetSurface"),
                                    [Is].EqualTo(expectedSurfaceOffset).Within(0.000000001R),
                                    "row " & rowNumber & " must replace the previous surface offset")
                        Assert.That(ReadDouble(restored, "axisOffset"),
                                    [Is].EqualTo(expectedAxisOffset).Within(0.000000001R),
                                    "row " & rowNumber & " horizontal offset must persist in BeamI JSON")
                        Assert.That(ReadDouble(restored, "offsetSurface"),
                                    [Is].EqualTo(expectedSurfaceOffset).Within(0.000000001R),
                                    "row " & rowNumber & " surface offset must persist in BeamI JSON")
                        Assert.That(ReadDouble(beam, "height"), [Is].EqualTo(2.34R))
                        Assert.That(ReadDouble(beam, "lenght"), [Is].EqualTo(33.4R))
                        Assert.That(ReadInteger(beam, "numberRow"), [Is].EqualTo(rowNumber))
                        Assert.That(ReadDouble(trajectory, "OffsetProjectAlignment"),
                                    [Is].EqualTo(expectedAxisOffset).Within(0.000000001R))
                        Assert.That(ReadDouble(trajectory, "OffsetProjectSurface"),
                                    [Is].EqualTo(expectedSurfaceOffset).Within(0.000000001R))
                    End Sub)
            Next
        End Sub

        <Test>
        Public Sub SavedRowsRemoveGlobalOffsetsBeforeRestoringMillimetreInputs()
            Dim method As MethodInfo = GetType(FormPlacementBeams).GetMethod(
                "GetSavedRowValues",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {
                    GetType(Double), GetType(Double), GetType(Double), GetType(Double)
                },
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "GetSavedRowValues must convert persisted total offsets back to row-local UI values.")

            For Each rowOffset As Double In RowOffsets
                Dim axisOffset As Double = GlobalHorizontalOffset + rowOffset
                Dim surfaceOffset As Double = GlobalVerticalOffset + PavementThickness
                Dim values As Double() = DirectCast(method.Invoke(
                    Nothing,
                    New Object() {
                        axisOffset, surfaceOffset, GlobalHorizontalOffset, GlobalVerticalOffset
                    }), Double())

                Assert.Multiple(
                    Sub()
                        Assert.That(values, Has.Length.EqualTo(2))
                        Assert.That(values(0),
                                    [Is].EqualTo(Math.Abs(rowOffset) * 1000.0R).Within(0.000000001R),
                                    "The row distance must not include the global horizontal shift twice.")
                        Assert.That(values(1),
                                    [Is].EqualTo(PavementThickness * 1000.0R).Within(0.000000001R),
                                    "The pavement thickness must not include the global vertical shift twice.")
                    End Sub)
            Next
        End Sub

        Private Shared Function CreateTrajectory(rowNumber As Integer,
                                                 surfaceOffset As Double,
                                                 alignmentOffset As Double) As Object
            Dim trajectoryType As Type = MainType("CivilEnginStructures.TrajectoryPlacementBeams")
            Dim trajectoryKind As Type = trajectoryType.GetNestedType("TypeTrajectoryPlacementBeams")
            Dim noneValue As Object = [Enum].Parse(trajectoryKind, "None")
            Return Activator.CreateInstance(
                trajectoryType,
                New Object() {rowNumber, surfaceOffset, alignmentOffset, noneValue})
        End Function

        Private Shared Sub ApplyTrajectoryOffsets(beam As Object, trajectory As Object)
            Dim calculationType As Type = MainType("CivilEnginStructures.CalculationBeams")
            Dim method As MethodInfo = calculationType.GetMethod(
                "ApplyTrajectoryOffsets",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {beam.GetType(), trajectory.GetType()},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "CalculationBeams.ApplyTrajectoryOffsets must copy the selected row trajectory into BeamI before geometry and JSON are calculated.")
            InvokeWithoutWrapper(method, Nothing, New Object() {beam, trajectory})
        End Sub

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

        Private Shared Function ReadDouble(target As Object, propertyName As String) As Double
            Return CDbl(target.GetType().GetProperty(propertyName).GetValue(target, Nothing))
        End Function

        Private Shared Function ReadInteger(target As Object, propertyName As String) As Integer
            Return CInt(target.GetType().GetProperty(propertyName).GetValue(target, Nothing))
        End Function

        Private Shared Sub WriteProperty(target As Object, propertyName As String, value As Object)
            target.GetType().GetProperty(propertyName).SetValue(target, value, Nothing)
        End Sub

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
                        "Build CivilEnginStructures.vbproj before running beam offset regression tests.")
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
