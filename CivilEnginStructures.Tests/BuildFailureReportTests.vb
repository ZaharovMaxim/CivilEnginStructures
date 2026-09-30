Imports System.IO
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BuildFailureReportTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub ReportSummaryContainsActionableBuildContext()
            Dim report As Object = CreateReport(
                "Построение балок",
                "Чтение проектной отметки",
                "Отметка отсутствует",
                "Назначьте проектную поверхность",
                "BridgeBuilder.Build",
                Nothing)

            Dim summary As String = ReportText(report, "Summary")

            Assert.Multiple(
                Sub()
                    Assert.That(summary, Does.Contain("Построение балок"))
                    Assert.That(summary, Does.Contain("Чтение проектной отметки"))
                    Assert.That(summary, Does.Contain("Отметка отсутствует"))
                    Assert.That(summary, Does.Contain("Назначьте проектную поверхность"))
                    Assert.That(report.GetType().GetProperty("Summary").CanWrite, [Is].False)
                    Assert.That(report.GetType().GetProperty("TechnicalDetails").CanWrite, [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub TechnicalDetailsIncludesSourceMethodWhenProvided()
            Dim report As Object = CreateReport(
                "Построение опоры",
                "Расчёт геометрии",
                "Не удалось построить контур",
                "Проверьте исходные точки",
                "PillarBuilder.CreateContour",
                Nothing)

            Assert.That(ReportText(report, "TechnicalDetails"),
                        Does.Contain("PillarBuilder.CreateContour"))
        End Sub

        <Test>
        Public Sub TechnicalDetailsIncludesFullExceptionStringWithInnerExceptionAndStackTrace()
            Dim failure As Exception = CaptureWrappedException()
            Dim report As Object = CreateReport(
                "Построение моста",
                "Создание тела",
                "Операция завершилась исключением",
                "Проверьте исходную геометрию",
                "BridgeBuilder.CreateSolid",
                failure)

            Dim details As String = ReportText(report, "TechnicalDetails")

            Assert.Multiple(
                Sub()
                    Assert.That(failure.StackTrace, [Is].Not.Null.And.Not.Empty)
                    Assert.That(details, Does.Contain(failure.ToString()))
                    Assert.That(details, Does.Contain("outer-build-marker"))
                    Assert.That(details, Does.Contain("inner-value-marker"))
                End Sub)
        End Sub

        <Test>
        Public Sub TryRequireFiniteAcceptsZeroAndWritesTheReadValue()
            Dim value As Double = 73.25
            Dim reason As String = "old failure"

            Dim accepted As Boolean = InvokeTryRequireFinite(
                Function() New Double?(0),
                "Проектная отметка",
                value,
                reason)

            Assert.Multiple(
                Sub()
                    Assert.That(accepted, [Is].True)
                    Assert.That(value, [Is].EqualTo(0))
                    Assert.That(reason, [Is].Null.Or.Empty)
                End Sub)
        End Sub

        <TestCase(125.75)>
        <TestCase(-18.5)>
        Public Sub TryRequireFiniteAcceptsOrdinaryFiniteValues(expected As Double)
            Dim value As Double = 999
            Dim reason As String = Nothing

            Dim accepted As Boolean = InvokeTryRequireFinite(
                Function() New Double?(expected),
                "Отметка земли",
                value,
                reason)

            Assert.Multiple(
                Sub()
                    Assert.That(accepted, [Is].True)
                    Assert.That(value, [Is].EqualTo(expected))
                    Assert.That(reason, [Is].Null.Or.Empty)
                End Sub)
        End Sub

        <Test>
        Public Sub TryRequireFiniteRejectsMissingAndNonfiniteValuesWithoutChangingOutput()
            Dim readers As Func(Of Double?)() =
                {
                    Function() CType(Nothing, Double?),
                    Function() New Double?(Double.NaN),
                    Function() New Double?(Double.PositiveInfinity),
                    Function() New Double?(Double.NegativeInfinity)
                }

            For Each reader As Func(Of Double?) In readers
                AssertInvalidValue(reader, "Отметка оси")
            Next
        End Sub

        <Test>
        Public Sub TryRequireFinitePropagatesTheReaderExceptionInstance()
            Dim expected As New InvalidOperationException("reader-failure-marker")
            Dim value As Double = 42
            Dim reason As String = Nothing
            MainType("CivilEnginStructures.BuildValueValidation")

            Dim thrown As TargetInvocationException = Assert.Throws(Of TargetInvocationException)(
                Sub()
                    InvokeTryRequireFinite(
                        Function() As Double?
                            Throw expected
                        End Function,
                        "Отметка насыпи",
                        value,
                        reason)
                End Sub)

            Assert.That(thrown.InnerException, [Is].SameAs(expected),
                        "MethodInfo adds one reflection wrapper; the production helper must not replace the original exception.")
        End Sub

        Private Shared Sub AssertInvalidValue(reader As Func(Of Double?), valueName As String)
            Const sentinel As Double = 314.159
            Dim value As Double = sentinel
            Dim reason As String = Nothing

            Dim accepted As Boolean = InvokeTryRequireFinite(reader, valueName, value, reason)

            Assert.That(accepted, [Is].False)
            Assert.That(value, [Is].EqualTo(sentinel),
                        "A failed read must not overwrite the caller's last valid value.")
            Assert.That(reason, Does.Contain(valueName))
        End Sub

        Private Shared Function CreateReport(operation As String,
                                             stage As String,
                                             reason As String,
                                             suggestedAction As String,
                                             sourceMethod As String,
                                             failure As Exception) As Object
            Dim reportType As Type = MainType("CivilEnginStructures.BuildFailureReport")
            Dim constructor As ConstructorInfo = reportType.GetConstructor(
                New Type() {
                    GetType(String),
                    GetType(String),
                    GetType(String),
                    GetType(String),
                    GetType(String),
                    GetType(Exception)
                })
            Assert.That(constructor, [Is].Not.Null,
                        "BuildFailureReport must expose the agreed six-value constructor.")
            Return constructor.Invoke(
                New Object() {operation, stage, reason, suggestedAction, sourceMethod, failure})
        End Function

        Private Shared Function ReportText(report As Object, propertyName As String) As String
            Dim propertyInfo As PropertyInfo = report.GetType().GetProperty(
                propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null,
                        "BuildFailureReport." & propertyName & " must be Public ReadOnly.")
            Return CStr(propertyInfo.GetValue(report, Nothing))
        End Function

        Private Shared Function InvokeTryRequireFinite(reader As Func(Of Double?),
                                                       valueName As String,
                                                       ByRef value As Double,
                                                       ByRef failureReason As String) As Boolean
            Dim validationType As Type = MainType("CivilEnginStructures.BuildValueValidation")
            Dim method As MethodInfo = validationType.GetMethod(
                "TryRequireFinite",
                BindingFlags.Public Or BindingFlags.Static,
                Nothing,
                New Type() {
                    GetType(Func(Of Double?)),
                    GetType(String),
                    GetType(Double).MakeByRefType(),
                    GetType(String).MakeByRefType()
                },
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BuildValueValidation.TryRequireFinite must expose the agreed pure validation contract.")

            Dim arguments As Object() = {reader, valueName, value, failureReason}
            Dim result As Boolean = CBool(method.Invoke(Nothing, arguments))
            value = CDbl(arguments(2))
            failureReason = TryCast(arguments(3), String)
            Return result
        End Function

        Private Shared Function CaptureWrappedException() As Exception
            Try
                Throw New ArgumentException("inner-value-marker")
            Catch innerFailure As Exception
                Try
                    Throw New InvalidOperationException("outer-build-marker", innerFailure)
                Catch outerFailure As Exception
                    Return outerFailure
                End Try
            End Try
            Return Nothing
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
                        "Build CivilEnginStructures.vbproj before running build failure report tests.")
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
