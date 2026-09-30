Imports System.IO
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BridgeModelSettingsInitializationTests
        Private Shared _mainAssembly As Assembly

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub NullOrBlankCandidateDoesNotInitializeProjectSurface(candidatePath As String)
            Dim settings As Object = CreateSettings()
            WriteProperty(settings, "EarthSurfaceRelativePath", "earth/surface.sfc")

            Dim initialized As Boolean = TryInitialize(settings, candidatePath)

            Assert.Multiple(
                Sub()
                    Assert.That(initialized, [Is].False)
                    Assert.That(ReadProperty(settings, "ProjectSurfaceRelativePath"), [Is].Null)
                    Assert.That(ReadProperty(settings, "EarthSurfaceRelativePath"),
                                [Is].EqualTo("earth/surface.sfc"))
                End Sub)
        End Sub

        <Test>
        Public Sub ExistingProjectSurfaceIsNotOverwritten()
            Dim settings As Object = CreateSettings()
            WriteProperty(settings, "ProjectSurfaceRelativePath", "surfaces/existing.sfc")
            WriteProperty(settings, "EarthSurfaceRelativePath", "surfaces/earth.sfc")

            Dim initialized As Boolean = TryInitialize(settings, "surfaces/candidate.sfc")

            Assert.Multiple(
                Sub()
                    Assert.That(initialized, [Is].False)
                    Assert.That(ReadProperty(settings, "ProjectSurfaceRelativePath"),
                                [Is].EqualTo("surfaces/existing.sfc"))
                    Assert.That(ReadProperty(settings, "EarthSurfaceRelativePath"),
                                [Is].EqualTo("surfaces/earth.sfc"))
                End Sub)
        End Sub

        <Test>
        Public Sub BlankProjectSurfaceIsInitializedExactlyOnce()
            Dim settings As Object = CreateSettings()
            WriteProperty(settings, "ProjectSurfaceRelativePath", "   ")
            WriteProperty(settings, "EarthSurfaceRelativePath", "surfaces/earth.sfc")
            Const firstPath As String = "..\Surfaces\Project Surface.sfc"
            Const secondPath As String = "surfaces/replacement.sfc"

            Dim firstResult As Boolean = TryInitialize(settings, firstPath)
            Dim secondResult As Boolean = TryInitialize(settings, secondPath)

            Assert.Multiple(
                Sub()
                    Assert.That(firstResult, [Is].True)
                    Assert.That(secondResult, [Is].False)
                    Assert.That(ReadProperty(settings, "ProjectSurfaceRelativePath"),
                                [Is].EqualTo(firstPath))
                    Assert.That(ReadProperty(settings, "EarthSurfaceRelativePath"),
                                [Is].EqualTo("surfaces/earth.sfc"))
                End Sub)
        End Sub

        Private Shared Function CreateSettings() As Object
            Return Activator.CreateInstance(MainType("CivilEnginStructures.BridgeModelSettings"))
        End Function

        Private Shared Function TryInitialize(settings As Object, candidatePath As String) As Boolean
            Dim method As MethodInfo = settings.GetType().GetMethod(
                "TryInitializeProjectSurfaceReference",
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {GetType(String)},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeModelSettings must expose the agreed instance initialization helper.")
            Return CBool(InvokeWithoutWrapper(method, settings, New Object() {candidatePath}))
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

        Private Shared Function ReadProperty(target As Object, propertyName As String) As Object
            Return target.GetType().GetProperty(propertyName).GetValue(target, Nothing)
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
                        "Build CivilEnginStructures.vbproj before running bridge settings initialization tests.")
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
