Imports System.IO
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class SurfaceReferenceClassificationTests
        Private Shared _mainAssembly As Assembly

        <TestCase("Поверхность 1.1", False, TestName:="Dotted_legacy_surface_name_is_not_an_exact_reference")>
        <TestCase("Поверхность проекта.sfcx", True)>
        <TestCase("Поверхность проекта.roadx", True)>
        <TestCase("Поверхность проекта.site", True)>
        <TestCase("Поверхность проекта.algx", True)>
        <TestCase("folder\Поверхность 1.1", True)>
        <TestCase("Поверхность проекта", False)>
        Public Sub ExactSurfaceReferenceDistinguishesPathsAndSupportedExtensions(value As String,
                                                                                 expected As Boolean)
            Dim funcSurfaceType As Type = MainAssembly().GetType("CivilEnginStructures.FuncSurface", False)
            Assert.That(funcSurfaceType, [Is].Not.Null)
            Dim classifier As MethodInfo = funcSurfaceType.GetMethod(
                "IsExactSurfaceReference",
                BindingFlags.NonPublic Or BindingFlags.Static,
                Nothing,
                New Type() {GetType(String)},
                Nothing)
            Assert.That(classifier, [Is].Not.Null,
                        "The regression test exercises the private pure classifier used before native surface lookup.")

            Assert.That(CBool(classifier.Invoke(Nothing, New Object() {value})), [Is].EqualTo(expected))
        End Sub

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
                        "Build CivilEnginStructures.vbproj before running surface reference tests.")
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
