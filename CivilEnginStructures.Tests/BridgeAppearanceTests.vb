Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation

Namespace Tests
    <TestFixture>
    Public Class BridgeAppearanceTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub ValuesConstructorPreservesAllValuesAndExposesReadOnlyProperties()
            Dim cadColorValue As Integer = CadColor.ByLayer.ToCompressValue()
            Dim values As Object = CreateValues(
                cadColorValue,
                "Мост-Балки",
                "Осевая",
                2.5,
                35,
                New Double?(0.4),
                New Integer?(Color.FromArgb(255, 20, 40, 60).ToArgb()))

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Integer)(values, "CadColorValue"), [Is].EqualTo(cadColorValue))
                    Assert.That(Value(Of String)(values, "LayerName"), [Is].EqualTo("Мост-Балки"))
                    Assert.That(Value(Of String)(values, "LinetypeName"), [Is].EqualTo("Осевая"))
                    Assert.That(Value(Of Double)(values, "LinetypeScale"), [Is].EqualTo(2.5))
                    Assert.That(Value(Of Integer)(values, "Lineweight"), [Is].EqualTo(35))
                    Assert.That(Value(Of Double?)(values, "Width"), [Is].EqualTo(New Double?(0.4)))
                    Assert.That(Value(Of Integer?)(values, "BodyArgb"),
                                [Is].EqualTo(New Integer?(Color.FromArgb(255, 20, 40, 60).ToArgb())))
                    For Each propertyInfo As PropertyInfo In values.GetType().GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                        Assert.That(propertyInfo.CanWrite, [Is].False,
                                    propertyInfo.Name & " must be read-only on an immutable appearance snapshot.")
                    Next
                End Sub)
        End Sub

        <Test>
        Public Sub SparsePatchChangesOnlySpecifiedValuesWithoutMutatingSource()
            Dim source As Object = StandardValues(CadColor.ByBlock.ToCompressValue())
            Dim patch As Object = CreatePatch()
            SetPatch(patch, "LayerName", "Мост-Пролёт")
            SetPatch(patch, "Width", 0.75)

            Dim result As Object = ApplyTo(patch, source)

            Assert.Multiple(
                Sub()
                    Assert.That(result, [Is].Not.SameAs(source))
                    Assert.That(Value(Of String)(result, "LayerName"), [Is].EqualTo("Мост-Пролёт"))
                    Assert.That(Value(Of Double?)(result, "Width"), [Is].EqualTo(New Double?(0.75)))
                    AssertUnchangedExcept(source, result, "LayerName", "Width")
                    Assert.That(Value(Of String)(source, "LayerName"), [Is].EqualTo("Исходный слой"))
                    Assert.That(Value(Of Double?)(source, "Width"), [Is].EqualTo(New Double?(0.25)))
                End Sub)
        End Sub

        <Test>
        Public Sub CadAndBodyColorsCanBePatchedIndependently()
            Dim originalCad As Integer = CadColor.ByLayer.ToCompressValue()
            Dim replacementCad As Integer = New CadColor(255).ToCompressValue()
            Dim originalBody As Integer = Color.FromArgb(255, 10, 20, 30).ToArgb()
            Dim replacementBody As Integer = Color.FromArgb(255, 80, 90, 100).ToArgb()
            Dim source As Object = CreateValues(
                originalCad, "Слой", "Continuous", 1.0, 25, Nothing, New Integer?(originalBody))

            Dim bodyPatch As Object = CreatePatch()
            SetPatch(bodyPatch, "BodyArgb", replacementBody)
            Dim bodyResult As Object = ApplyTo(bodyPatch, source)

            Dim cadPatch As Object = CreatePatch()
            SetPatch(cadPatch, "CadColorValue", replacementCad)
            Dim cadResult As Object = ApplyTo(cadPatch, source)

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Integer)(bodyResult, "CadColorValue"), [Is].EqualTo(originalCad))
                    Assert.That(Value(Of Integer?)(bodyResult, "BodyArgb"), [Is].EqualTo(New Integer?(replacementBody)))
                    Assert.That(Value(Of Integer)(cadResult, "CadColorValue"), [Is].EqualTo(replacementCad))
                    Assert.That(Value(Of Integer?)(cadResult, "BodyArgb"), [Is].EqualTo(New Integer?(originalBody)))
                    Assert.That(Value(Of Integer)(source, "CadColorValue"), [Is].EqualTo(originalCad))
                    Assert.That(Value(Of Integer?)(source, "BodyArgb"), [Is].EqualTo(New Integer?(originalBody)))
                End Sub)
        End Sub

        <TestCaseSource(NameOf(CompressedCadColors))>
        Public Sub SparsePatchPreservesCompressedCadColorExactly(cadColorValue As Integer)
            Dim source As Object = StandardValues(cadColorValue)
            Dim patch As Object = CreatePatch()
            SetPatch(patch, "Width", 1.125)

            Dim result As Object = ApplyTo(patch, source)

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Integer)(result, "CadColorValue"), [Is].EqualTo(cadColorValue))
                    Assert.That(CadColor.FromCompressValue(Value(Of Integer)(result, "CadColorValue")),
                                [Is].EqualTo(CadColor.FromCompressValue(cadColorValue)))
                    Assert.That(Value(Of Double?)(result, "Width"), [Is].EqualTo(New Double?(1.125)))
                End Sub)
        End Sub

        <Test>
        Public Sub NumericValidationAcceptsPositiveScaleAndMissingOrZeroWidth()
            Dim validValues As Tuple(Of Double, Double?)() = {
                Tuple.Create(0.000001, CType(Nothing, Double?)),
                Tuple.Create(1.0, New Double?(0.0)),
                Tuple.Create(250.0, New Double?(12.5))
            }

            For Each item As Tuple(Of Double, Double?) In validValues
                Dim patch As Object = CreatePatch()
                SetPatch(patch, "LinetypeScale", item.Item1)
                If item.Item2.HasValue Then SetPatch(patch, "Width", item.Item2.Value)
                Dim reason As String = "old error"

                Assert.That(TryValidate(patch, reason), [Is].True)
                Assert.That(reason, [Is].Null.Or.Empty)
                Assert.That(ApplyTo(patch, StandardValues(CadColor.ByLayer.ToCompressValue())), [Is].Not.Null)
            Next
        End Sub

        <Test>
        Public Sub InvalidNumericPatchIsRejectedWithoutMutatingSource()
            Dim invalidValues As Tuple(Of String, Double)() = {
                Tuple.Create("LinetypeScale", 0.0),
                Tuple.Create("LinetypeScale", -1.0),
                Tuple.Create("LinetypeScale", Double.NaN),
                Tuple.Create("LinetypeScale", Double.PositiveInfinity),
                Tuple.Create("LinetypeScale", Double.NegativeInfinity),
                Tuple.Create("Width", -0.001),
                Tuple.Create("Width", Double.NaN),
                Tuple.Create("Width", Double.PositiveInfinity),
                Tuple.Create("Width", Double.NegativeInfinity)
            }

            For Each item As Tuple(Of String, Double) In invalidValues
                Dim source As Object = StandardValues(CadColor.ByBlock.ToCompressValue())
                Dim before As Object() = Snapshot(source)
                Dim patch As Object = CreatePatch()
                SetPatch(patch, item.Item1, item.Item2)
                Dim reason As String = Nothing

                Assert.That(TryValidate(patch, reason), [Is].False, item.Item1 & "=" & item.Item2)
                Assert.That(reason, [Is].Not.Null.And.Not.Empty)
                Dim thrown As TargetInvocationException = Assert.Throws(Of TargetInvocationException)(
                    Sub() ApplyTo(patch, source))
                Assert.That(thrown.InnerException, [Is].TypeOf(Of ArgumentException)())
                Assert.That(Snapshot(source), [Is].EqualTo(before),
                            "Validation failure must occur before changing the source snapshot.")
            Next
        End Sub

        <Test>
        Public Sub ForTargetDropsUnsupportedWidthAndBodyWithoutMutatingOriginalPatch()
            Dim patch As Object = CreatePatch()
            SetPatch(patch, "CadColorValue", CadColor.ByLayer.ToCompressValue())
            SetPatch(patch, "LayerName", "Мост-Балки")
            SetPatch(patch, "LinetypeName", "Continuous")
            SetPatch(patch, "LinetypeScale", 2.0)
            SetPatch(patch, "Lineweight", 35)
            SetPatch(patch, "Width", 0.75)
            SetPatch(patch, "BodyArgb", Color.FromArgb(255, 10, 20, 30).ToArgb())

            Dim effective As Object = ForTarget(patch, False, False)

            Assert.Multiple(
                Sub()
                    Assert.That(effective, [Is].Not.SameAs(patch))
                    Assert.That(Value(Of Integer?)(effective, "CadColorValue"),
                                [Is].EqualTo(Value(Of Integer?)(patch, "CadColorValue")))
                    Assert.That(Value(Of String)(effective, "LayerName"), [Is].EqualTo("Мост-Балки"))
                    Assert.That(Value(Of String)(effective, "LinetypeName"), [Is].EqualTo("Continuous"))
                    Assert.That(Value(Of Double?)(effective, "LinetypeScale"), [Is].EqualTo(New Double?(2.0)))
                    Assert.That(Value(Of Integer?)(effective, "Lineweight"), [Is].EqualTo(New Integer?(35)))
                    Assert.That(Value(Of Double?)(effective, "Width"), [Is].Null)
                    Assert.That(Value(Of Integer?)(effective, "BodyArgb"), [Is].Null)
                    Assert.That(Value(Of Double?)(patch, "Width"), [Is].EqualTo(New Double?(0.75)))
                    Assert.That(Value(Of Integer?)(patch, "BodyArgb"),
                                [Is].EqualTo(New Integer?(Color.FromArgb(255, 10, 20, 30).ToArgb())))
                End Sub)
        End Sub

        <Test>
        Public Sub ForTargetKeepsOnlyCompatibleFieldsAndCanProduceAnEmptyPatch()
            Dim fullPatch As Object = CreatePatch()
            SetPatch(fullPatch, "Width", 1.25)
            SetPatch(fullPatch, "BodyArgb", Color.Red.ToArgb())

            Dim widthOnly As Object = ForTarget(fullPatch, True, False)
            Dim bodyOnly As Object = ForTarget(fullPatch, False, True)
            Dim both As Object = ForTarget(fullPatch, True, True)

            Dim incompatibleOnly As Object = CreatePatch()
            SetPatch(incompatibleOnly, "Width", 1.25)
            SetPatch(incompatibleOnly, "BodyArgb", Color.Red.ToArgb())
            Dim empty As Object = ForTarget(incompatibleOnly, False, False)

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Double?)(widthOnly, "Width"), [Is].EqualTo(New Double?(1.25)))
                    Assert.That(Value(Of Integer?)(widthOnly, "BodyArgb"), [Is].Null)
                    Assert.That(Value(Of Double?)(bodyOnly, "Width"), [Is].Null)
                    Assert.That(Value(Of Integer?)(bodyOnly, "BodyArgb"), [Is].EqualTo(New Integer?(Color.Red.ToArgb())))
                    Assert.That(Value(Of Double?)(both, "Width"), [Is].EqualTo(New Double?(1.25)))
                    Assert.That(Value(Of Integer?)(both, "BodyArgb"), [Is].EqualTo(New Integer?(Color.Red.ToArgb())))
                    Assert.That(IsEmptyPatch(empty), [Is].True,
                                "A patch containing only incompatible fields must be a no-op for that target.")
                End Sub)
        End Sub

        <Test>
        Public Sub ExtendedValuesAndSparsePatchPreserveIndependentFillSettings()
            Dim source As Object = CreateExtendedValues(
                New Integer?(New CadColor(5).ToCompressValue()), "SOLID", New Double?(1.0R), New Double?(0.0R))
            Dim patch As Object = CreatePatch()
            Dim replacementFill As Integer = New CadColor(Color.FromArgb(40, 80, 120)).ToCompressValue()
            SetPatch(patch, "FillCadColorValue", replacementFill)
            SetPatch(patch, "HatchPatternName", "ANSI31")
            SetPatch(patch, "HatchScale", 2.5R)
            SetPatch(patch, "HatchAngle", -0.75R)

            Dim result As Object = ApplyTo(patch, source)

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Integer?)(result, "FillCadColorValue"),
                                [Is].EqualTo(New Integer?(replacementFill)))
                    Assert.That(Value(Of String)(result, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                    Assert.That(Value(Of Double?)(result, "HatchScale"), [Is].EqualTo(New Double?(2.5R)))
                    Assert.That(Value(Of Double?)(result, "HatchAngle"), [Is].EqualTo(New Double?(-0.75R)))
                    AssertUnchangedExcept(source, result,
                                          "FillCadColorValue", "HatchPatternName", "HatchScale", "HatchAngle")
                    Assert.That(Value(Of String)(source, "HatchPatternName"), [Is].EqualTo("SOLID"))
                End Sub)
        End Sub

        <Test>
        Public Sub FillValidationAcceptsSupportedValuesAndRejectsInvalidPatternScaleAndAngle()
            Dim valid As Object = CreatePatch()
            SetPatch(valid, "HatchPatternName", " ansi31 ")
            SetPatch(valid, "HatchScale", 0.000001R)
            SetPatch(valid, "HatchAngle", -Math.PI)
            Dim reason As String = "old"
            Assert.That(TryValidate(valid, reason), [Is].True)
            Assert.That(reason, [Is].Null.Or.Empty)

            For Each invalidPattern As String In {String.Empty, "   ", "AR-CONC"}
                Dim patch As Object = CreatePatch()
                SetPatch(patch, "HatchPatternName", invalidPattern)
                reason = Nothing
                Assert.That(TryValidate(patch, reason), [Is].False, "pattern=" & invalidPattern)
                Assert.That(reason, [Is].Not.Null.And.Not.Empty)
            Next
            For Each invalidScale As Double In {
                0.0R, -1.0R, Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity
            }
                Dim patch As Object = CreatePatch()
                SetPatch(patch, "HatchScale", invalidScale)
                reason = Nothing
                Assert.That(TryValidate(patch, reason), [Is].False, "scale=" & invalidScale)
            Next
            For Each invalidAngle As Double In {
                Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity
            }
                Dim patch As Object = CreatePatch()
                SetPatch(patch, "HatchAngle", invalidAngle)
                reason = Nothing
                Assert.That(TryValidate(patch, reason), [Is].False, "angle=" & invalidAngle)
            Next
        End Sub

        <Test>
        Public Sub ForTargetKeepsFillOnlyForFillCapableTargetsWithoutMutatingPatch()
            Dim patch As Object = CreatePatch()
            SetPatch(patch, "FillCadColorValue", New CadColor(4).ToCompressValue())
            SetPatch(patch, "HatchPatternName", "ANSI31")
            SetPatch(patch, "HatchScale", 3.0R)
            SetPatch(patch, "HatchAngle", 0.5R)

            Dim capable As Object = ForTarget(patch, False, False, True)
            Dim unsupported As Object = ForTarget(patch, False, False, False)

            Assert.Multiple(
                Sub()
                    Assert.That(Value(Of Integer?)(capable, "FillCadColorValue"),
                                [Is].EqualTo(Value(Of Integer?)(patch, "FillCadColorValue")))
                    Assert.That(Value(Of String)(capable, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                    Assert.That(Value(Of Double?)(capable, "HatchScale"), [Is].EqualTo(New Double?(3.0R)))
                    Assert.That(Value(Of Double?)(capable, "HatchAngle"), [Is].EqualTo(New Double?(0.5R)))
                    Assert.That(IsEmptyPatch(capable), [Is].False)
                    Assert.That(IsEmptyPatch(unsupported), [Is].True)
                    Assert.That(Value(Of String)(patch, "HatchPatternName"), [Is].EqualTo("ANSI31"))
                End Sub)
        End Sub

        Private Shared Function CompressedCadColors() As IEnumerable
            Return New Object() {
                New TestCaseData(CadColor.ByLayer.ToCompressValue()).SetName("Sparse_patch_preserves_ByLayer"),
                New TestCaseData(CadColor.ByBlock.ToCompressValue()).SetName("Sparse_patch_preserves_ByBlock"),
                New TestCaseData(CadColor.Empty.ToCompressValue()).SetName("Sparse_patch_preserves_Empty"),
                New TestCaseData(CadColor.Background.ToCompressValue()).SetName("Sparse_patch_preserves_Background"),
                New TestCaseData((New CadColor(1)).ToCompressValue()).SetName("Sparse_patch_preserves_indexed_1"),
                New TestCaseData((New CadColor(255)).ToCompressValue()).SetName("Sparse_patch_preserves_indexed_255"),
                New TestCaseData((New CadColor(Color.FromArgb(12, 34, 56))).ToCompressValue()).SetName("Sparse_patch_preserves_RGB")
            }
        End Function

        Private Shared Function StandardValues(cadColorValue As Integer) As Object
            Return CreateValues(
                cadColorValue,
                "Исходный слой",
                "Continuous",
                1.5,
                30,
                New Double?(0.25),
                New Integer?(Color.FromArgb(255, 30, 60, 90).ToArgb()))
        End Function

        Private Shared Function CreateValues(cadColorValue As Integer,
                                              layerName As String,
                                              linetypeName As String,
                                              linetypeScale As Double,
                                              lineweight As Integer,
                                              width As Double?,
                                              bodyArgb As Integer?) As Object
            Dim valuesType As Type = MainType("CivilEnginStructures.BridgeAppearanceValues")
            Dim constructor As ConstructorInfo = valuesType.GetConstructor(
                New Type() {
                    GetType(Integer),
                    GetType(String),
                    GetType(String),
                    GetType(Double),
                    GetType(Integer),
                    GetType(Double?),
                    GetType(Integer?)
                })
            Assert.That(constructor, [Is].Not.Null,
                        "BridgeAppearanceValues must expose the agreed seven-value constructor.")
            Return constructor.Invoke(
                New Object() {cadColorValue, layerName, linetypeName, linetypeScale, lineweight, width, bodyArgb})
        End Function

        Private Shared Function CreateExtendedValues(fillCadColorValue As Integer?,
                                                     hatchPatternName As String,
                                                     hatchScale As Double?,
                                                     hatchAngle As Double?) As Object
            Dim valuesType As Type = MainType("CivilEnginStructures.BridgeAppearanceValues")
            Dim constructor As ConstructorInfo = valuesType.GetConstructor(
                New Type() {
                    GetType(Integer), GetType(String), GetType(String), GetType(Double),
                    GetType(Integer), GetType(Double?), GetType(Integer?), GetType(Integer?),
                    GetType(String), GetType(Double?), GetType(Double?)
                })
            Assert.That(constructor, [Is].Not.Null,
                        "BridgeAppearanceValues must expose the fill-aware constructor without replacing the legacy one.")
            Return constructor.Invoke(
                New Object() {
                    CadColor.ByLayer.ToCompressValue(), "Исходный слой", "Continuous", 1.5R,
                    30, New Double?(0.25R), New Integer?(Color.Navy.ToArgb()), fillCadColorValue,
                    hatchPatternName, hatchScale, hatchAngle
                })
        End Function

        Private Shared Function CreatePatch() As Object
            Return Activator.CreateInstance(MainType("CivilEnginStructures.BridgeAppearancePatch"))
        End Function

        Private Shared Sub SetPatch(patch As Object, propertyName As String, value As Object)
            Dim propertyInfo As PropertyInfo = patch.GetType().GetProperty(
                propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null.And.Property("CanWrite").True,
                        "BridgeAppearancePatch." & propertyName & " must be a public settable sparse field.")
            propertyInfo.SetValue(patch, value, Nothing)
        End Sub

        Private Shared Function ApplyTo(patch As Object, source As Object) As Object
            Dim method As MethodInfo = patch.GetType().GetMethod(
                "ApplyTo",
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                New Type() {source.GetType()},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeAppearancePatch.ApplyTo must accept and return BridgeAppearanceValues.")
            Assert.That(method.ReturnType, [Is].EqualTo(source.GetType()))
            Return method.Invoke(patch, New Object() {source})
        End Function

        Private Shared Function ForTarget(patch As Object,
                                          supportsWidth As Boolean,
                                          isModel3D As Boolean) As Object
            Dim method As MethodInfo = patch.GetType().GetMethod(
                "ForTarget",
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(Boolean), GetType(Boolean)},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeAppearancePatch.ForTarget must expose the agreed pure compatibility filter.")
            Assert.That(method.ReturnType, [Is].EqualTo(patch.GetType()))
            Return method.Invoke(patch, New Object() {supportsWidth, isModel3D})
        End Function

        Private Shared Function ForTarget(patch As Object,
                                          supportsWidth As Boolean,
                                          isModel3D As Boolean,
                                          supportsFill As Boolean) As Object
            Dim method As MethodInfo = patch.GetType().GetMethod(
                "ForTarget",
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(Boolean), GetType(Boolean), GetType(Boolean)},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeAppearancePatch.ForTarget must expose the fill capability overload.")
            Return method.Invoke(patch, New Object() {supportsWidth, isModel3D, supportsFill})
        End Function

        Private Shared Function IsEmptyPatch(patch As Object) As Boolean
            Dim propertyInfo As PropertyInfo = patch.GetType().GetProperty(
                "IsEmpty", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null)
            Return CBool(propertyInfo.GetValue(patch, Nothing))
        End Function

        Private Shared Function TryValidate(patch As Object, ByRef reason As String) As Boolean
            Dim method As MethodInfo = patch.GetType().GetMethod(
                "TryValidate",
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(String).MakeByRefType()},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeAppearancePatch.TryValidate must expose the agreed pure validation contract.")
            Dim arguments As Object() = {reason}
            Dim result As Boolean = CBool(method.Invoke(patch, arguments))
            reason = TryCast(arguments(0), String)
            Return result
        End Function

        Private Shared Function Value(Of T)(instance As Object, propertyName As String) As T
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(
                propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null, propertyName & " must be public.")
            Return DirectCast(propertyInfo.GetValue(instance, Nothing), T)
        End Function

        Private Shared Sub AssertUnchangedExcept(source As Object,
                                                 result As Object,
                                                 ParamArray changedNames As String())
            For Each propertyInfo As PropertyInfo In source.GetType().GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                If Not changedNames.Contains(propertyInfo.Name) Then
                    Assert.That(propertyInfo.GetValue(result, Nothing),
                                [Is].EqualTo(propertyInfo.GetValue(source, Nothing)),
                                propertyInfo.Name & " was not specified by the sparse patch.")
                End If
            Next
        End Sub

        Private Shared Function Snapshot(values As Object) As Object()
            Return {
                Value(Of Integer)(values, "CadColorValue"),
                Value(Of String)(values, "LayerName"),
                Value(Of String)(values, "LinetypeName"),
                Value(Of Double)(values, "LinetypeScale"),
                Value(Of Integer)(values, "Lineweight"),
                Value(Of Double?)(values, "Width"),
                Value(Of Integer?)(values, "BodyArgb"),
                Value(Of Integer?)(values, "FillCadColorValue"),
                Value(Of String)(values, "HatchPatternName"),
                Value(Of Double?)(values, "HatchScale"),
                Value(Of Double?)(values, "HatchAngle")
            }
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
                        "Build CivilEnginStructures.vbproj before running bridge appearance tests.")
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
