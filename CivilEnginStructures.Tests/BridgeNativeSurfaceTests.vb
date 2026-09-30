Imports System.IO
Imports System.Reflection
Imports NUnit.Framework
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.Sfc
Imports Topomatic.Stg

Namespace Tests
    <TestFixture>
    Public Class BridgeNativeSurfaceTests
        Private Const RootTypeName As String = "CivilEnginStructures.InfrastradaBridgesModel"
        Private Const RuntimeTypeName As String = "CivilEnginStructures.BridgeModelRuntime"
        Private Const LayerTypeName As String = "CivilEnginStructures.BridgePlanCompoundLayer"

        <Test>
        Public Sub BridgeRootExposesTheNativeSurfaceModelContracts()
            Dim rootType As Type = MainType(RootTypeName)

            Assert.Multiple(
                Sub()
                    Assert.That(GetType(StateControllerObject).IsAssignableFrom(rootType), [Is].True)
                    Assert.That(GetType(IDisposable).IsAssignableFrom(rootType), [Is].True)
                    Assert.That(GetType(IDrawingContainer).IsAssignableFrom(rootType), [Is].True)
                    Assert.That(GetType(ITerrainModel).IsAssignableFrom(rootType), [Is].True)
                    Assert.That(GetType(ISurfaceContainer).IsAssignableFrom(rootType), [Is].True)
                    Assert.That(GetType(IStateElevationProviderFactory).IsAssignableFrom(rootType), [Is].True)
                    AssertProperty(rootType, "Arrangement", GetType(ArrangementModel))
                    AssertProperty(rootType, "Surface", GetType(Surface))
                    AssertProperty(rootType, "Drawing", GetType(Drawing))
                End Sub)
        End Sub

        <Test>
        Public Sub BridgeRootPublishesNativeStreamPersistenceWithoutExposingASecondDrawing()
            Dim rootType As Type = MainType(RootTypeName)
            Dim saveMethod As MethodInfo = PublicInstanceMethod(rootType, "NativeSaveToStream", GetType(Stream))
            Dim loadMethod As MethodInfo = PublicInstanceMethod(rootType, "NativeLoadFromStream", GetType(Stream))
            Dim saveFileMethod As MethodInfo = PublicInstanceMethod(rootType, "NativeSaveToFile", GetType(String))
            Dim loadFileMethod As MethodInfo = PublicInstanceMethod(rootType, "NativeLoadFromFile", GetType(String))
            Dim drawingProperties As PropertyInfo() = rootType.GetProperties(BindingFlags.Public Or BindingFlags.Instance).
                Where(Function(item) item.PropertyType Is GetType(Drawing)).ToArray()

            Assert.Multiple(
                Sub()
                    Assert.That(saveMethod, [Is].Not.Null)
                    If saveMethod IsNot Nothing Then Assert.That(saveMethod.ReturnType, [Is].EqualTo(GetType(Void)))
                    Assert.That(loadMethod, [Is].Not.Null)
                    If loadMethod IsNot Nothing Then Assert.That(loadMethod.ReturnType, [Is].EqualTo(GetType(Void)))
                    Assert.That(saveFileMethod, [Is].Not.Null)
                    Assert.That(loadFileMethod, [Is].Not.Null)
                    Assert.That(drawingProperties.Select(Function(item) item.Name),
                                [Is].EqualTo(New String() {"Drawing"}),
                                "The bridge root must publish one canonical Drawing: Surface.Situation.")
                End Sub)

            Dim disposedRoot As Object = System.Runtime.Serialization.FormatterServices.
                GetUninitializedObject(rootType)
            DirectCast(disposedRoot, IDisposable).Dispose()
            Dim existingPath As String = Path.Combine(
                Path.GetTempPath(),
                "infrastrada-disposed-save-" & Guid.NewGuid().ToString("N") & ".infrabridgex")
            Dim sentinel As Byte() = System.Text.Encoding.UTF8.GetBytes("existing-model-must-survive")
            Try
                File.WriteAllBytes(existingPath, sentinel)

                Assert.That(
                    Sub() InvokeInstanceUnwrapped(disposedRoot, "NativeSaveToFile", existingPath),
                    Throws.TypeOf(Of ObjectDisposedException)())
                Assert.That(File.ReadAllBytes(existingPath), [Is].EqualTo(sentinel),
                            "A failed save must validate before opening the existing file with FileMode.Create.")
            Finally
                If File.Exists(existingPath) Then File.Delete(existingPath)
            End Try
        End Sub

        <Test>
        Public Sub RuntimeCompatibilityHelpersHaveStableSignaturesAndNullFallbacks()
            Dim runtimeType As Type = MainType(RuntimeTypeName)
            Dim rootType As Type = MainType(RootTypeName)
            Dim getArrangement As MethodInfo = PublicSharedMethod(runtimeType, "GetArrangement", GetType(Object))
            Dim getDrawing As MethodInfo = PublicSharedMethod(runtimeType, "GetDrawing", GetType(ArrangementModel))
            Dim getRoot As MethodInfo = PublicSharedMethod(runtimeType, "GetRoot", GetType(ArrangementModel))
            Dim getModelObject As MethodInfo = PublicSharedMethod(runtimeType, "GetModelObject", GetType(ArrangementModel))

            Assert.Multiple(
                Sub()
                    Assert.That(getArrangement, [Is].Not.Null)
                    Assert.That(getArrangement.ReturnType, [Is].EqualTo(GetType(ArrangementModel)))
                    Assert.That(getDrawing, [Is].Not.Null)
                    Assert.That(getDrawing.ReturnType, [Is].EqualTo(GetType(Drawing)))
                    Assert.That(getRoot, [Is].Not.Null)
                    Assert.That(getRoot.ReturnType, [Is].EqualTo(rootType))
                    Assert.That(getModelObject, [Is].Not.Null)
                    Assert.That(getModelObject.ReturnType, [Is].EqualTo(GetType(Object)))
                End Sub)

            Assert.Multiple(
                Sub()
                    Assert.That(getArrangement.Invoke(Nothing, New Object() {Nothing}), [Is].Null)
                    Assert.That(getDrawing.Invoke(Nothing, New Object() {Nothing}), [Is].Null)
                    Assert.That(getRoot.Invoke(Nothing, New Object() {Nothing}), [Is].Null)
                    Assert.That(getModelObject.Invoke(Nothing, New Object() {Nothing}), [Is].Null)
                End Sub)
        End Sub

        <Test>
        Public Sub ElevationProviderDoesNotRepublishItsResolvedSurfacesAsReferences()
            Dim rootType As Type = MainType(RootTypeName)
            Dim providerType As Type = rootType.GetNestedType(
                "BridgeElevationProvider", BindingFlags.NonPublic)
            Assert.That(providerType, [Is].Not.Null)
            Dim constructor As ConstructorInfo = providerType.GetConstructor(
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing, New Type() {rootType}, Nothing)
            Assert.That(constructor, [Is].Not.Null)
            Dim provider As Object = constructor.Invoke(New Object() {Nothing})
            Dim surfacesField As FieldInfo = providerType.GetField(
                "_surfaces", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(surfacesField, [Is].Not.Null)
            Dim surfaces As System.Collections.IList = DirectCast(
                surfacesField.GetValue(provider), System.Collections.IList)
            surfaces.Add(System.Runtime.Serialization.FormatterServices.
                         GetUninitializedObject(GetType(Surface)))

            Dim references As System.Collections.IEnumerable = DirectCast(
                providerType.GetMethod("GetReferences", BindingFlags.Public Or BindingFlags.Instance).
                    Invoke(provider, Nothing),
                System.Collections.IEnumerable)

            Assert.That(references.Cast(Of Object)(), [Is].Empty,
                        "Resolved surfaces are elevation inputs, not nested model references.")
        End Sub

        <Test>
        Public Sub BridgePlanLayerExposesTheSameCanonicalSurfaceAndDrawing()
            Dim layerType As Type = MainType(LayerTypeName)
            Dim rootType As Type = MainType(RootTypeName)
            Dim rootBinding As PropertyInfo() = layerType.GetProperties(BindingFlags.Public Or BindingFlags.Instance).
                Where(Function(item) item.PropertyType Is rootType AndAlso item.CanRead AndAlso item.CanWrite).ToArray()

            Assert.Multiple(
                Sub()
                    Assert.That(GetType(IDrawingContainer).IsAssignableFrom(layerType), [Is].True)
                    Assert.That(GetType(ISurfaceContainer).IsAssignableFrom(layerType), [Is].True)
                    Assert.That(layerType.GetConstructor(New Type() {GetType(String)}), [Is].Not.Null)
                    AssertProperty(layerType, "Drawing", GetType(Drawing))
                    AssertProperty(layerType, "Surface", GetType(Surface))
                    Assert.That(rootBinding.Length, [Is].EqualTo(1),
                                "The plan layer needs one writable root binding so reload can rebind both native layers.")
                End Sub)
        End Sub

        <Test>
        Public Sub BridgeEditorLoadsTheNativeRootAndCreatesTheCompoundSurfaceLayer()
            Dim source As String = File.ReadAllText(Path.Combine(RepositoryRoot(), "ModelBridge", "BridgesModelEditor.vb"))

            Assert.Multiple(
                Sub()
                    Assert.That(source, Does.Contain("New InfrastradaBridgesModel"))
                    Assert.That(source, Does.Contain("BridgePlanCompoundLayer"))
                    Assert.That(source, Does.Not.Contain("Dim model As New ArrangementModel"),
                                "The editor must return the native bridge root, not the legacy backend object.")
                End Sub)
        End Sub

        <Test>
        Public Sub BridgeCreationDoesNotRequestANameForTheArrangementBackendProxy()
            Dim source As String = File.ReadAllText(
                Path.Combine(RepositoryRoot(), "UserForms", "Bridge", "FormPlacementBeams.vb"))
            Dim startIndex As Integer = source.IndexOf(
                "Private Sub Button5_Click", StringComparison.Ordinal)
            Assert.That(startIndex, [Is].GreaterThanOrEqualTo(0))
            Dim endIndex As Integer = source.IndexOf("End Sub", startIndex, StringComparison.Ordinal)
            Assert.That(endIndex, [Is].GreaterThan(startIndex))
            Dim methodBody As String = source.Substring(startIndex, endIndex - startIndex)

            Assert.Multiple(
                Sub()
                    Assert.That(methodBody, Does.Contain("createBridgeModel()"))
                    Assert.That(methodBody,
                                Does.Not.Contain("Plugins.Execute(""getname"", New Object() {arrangementProject})"),
                                "The backend ArrangementModel is not the registered project model object.")
                    Assert.That(methodBody, Does.Not.Match("\\bnameProject\\b"))
                End Sub)
        End Sub

        <Test>
        Public Sub MonolithSiteSelectorsUseTheSelectedBridgeDrawingInBothHandlers()
            Dim source As String = File.ReadAllText(
                Path.Combine(RepositoryRoot(), "UserForms", "Bridge", "FormCreateMonolitSitesBeams.vb"))
            Dim modelHandler As String = ExtractMethodSource(
                source, "Private Sub ComboBox1_SelectedIndexChanged")
            Dim bridgeHandler As String = ExtractMethodSource(
                source, "Private Sub ComboBox2_SelectedIndexChanged")

            For Each methodBody As String In New String() {modelHandler, bridgeHandler}
                Dim modelObjectIndex As Integer = methodBody.IndexOf(
                    "BridgeModelRuntime.GetModelObject(userSubObjectBearm)", StringComparison.Ordinal)
                Dim findModelIndex As Integer = methodBody.IndexOf(
                    "PluginCoreOps.FindModel", StringComparison.Ordinal)
                Dim activateIndex As Integer = methodBody.IndexOf(
                    "Plugins.Execute(""activate""", StringComparison.Ordinal)
                Dim drawingIndex As Integer = methodBody.IndexOf(
                    "BridgeModelRuntime.GetDrawing", StringComparison.Ordinal)
                Assert.Multiple(
                    Sub()
                        Assert.That(methodBody, Does.Contain("civilStructuresProject.getArrangementModelByIndex"))
                        Assert.That(methodBody, Does.Contain("BridgeModelRuntime.GetDrawing"))
                        Assert.That(methodBody, Does.Not.Contain(".arrx"))
                        Assert.That(methodBody, Does.Not.Contain("child.Model"))
                        Assert.That(methodBody, Does.Not.Contain("userSubObjectBearm.Drawing"))
                        Assert.That(modelObjectIndex, [Is].GreaterThanOrEqualTo(0))
                        Assert.That(findModelIndex, [Is].GreaterThanOrEqualTo(0))
                        Assert.That(activateIndex, [Is].GreaterThanOrEqualTo(0))
                        Assert.That(modelObjectIndex, [Is].LessThan(drawingIndex))
                        Assert.That(findModelIndex, [Is].LessThan(drawingIndex))
                        Assert.That(activateIndex, [Is].LessThan(drawingIndex),
                                    "The selected bridge project model must be active before its Drawing is used.")
                    End Sub)
            Next
        End Sub

        <Test>
        Public Sub ConeCreationTargetsTheBridgeRootSurfaceWithoutCreatingASiteModel()
            Dim root As String = RepositoryRoot()
            Dim formSource As String = File.ReadAllText(
                Path.Combine(root, "UserForms", "Bridge", "FormCreateConeLastPillars.vb"))
            Dim moduleSource As String = File.ReadAllText(Path.Combine(root, "RopExample1Module.vb"))

            Assert.Multiple(
                Sub()
                    Assert.That(formSource, Does.Contain("BridgeModelRuntime.GetRoot"))
                    Assert.That(formSource, Does.Contain(".Surface"))
                    Assert.That(formSource, Does.Not.Contain("SiteModel"))
                    Assert.That(formSource, Does.Not.Contain("New Object() {folder, ""site""}"))
                    Assert.That(moduleSource, Does.Not.Contain("FormCreateConePillar.CB_NameSites"))
                End Sub)
        End Sub

        <Test, Explicit("Requires the initialized Topomatic host for native Surface/Drawing construction.")>
        <Category("TopomaticIntegration")>
        Public Sub NativeBridgeRoundTripPreservesStructureLinesDrawingAndExternalSurfaceSettings()
            Dim original As Object = NewRoot()
            Try
                Dim surface As Surface = RootSurface(original)
                AddStructureLine(surface,
                                 New Vector3D(10.25, 20.5, 101.75),
                                 New Vector3D(30.75, 40.125, 99.5),
                                 "Конус К1")
                RootDrawing(original).ActiveSpace.AddLine(
                    New Vector3D(1.5, 2.5, 3.5), New Vector3D(4.5, 5.5, 6.5))
                SetExternalSurfaceSettings(RootArrangement(original),
                                           "roads\\A-101.roadx", "terrain\\earth.sfcx")

                Using stream As New MemoryStream()
                    InvokeInstance(original, "NativeSaveToStream", stream)
                    stream.Position = 0
                    Dim restored As Object = NewRoot()
                    Try
                        InvokeInstance(restored, "NativeLoadFromStream", stream)
                        Dim restoredSurface As Surface = RootSurface(restored)
                        Dim restoredSettings As Object = GetSettings(RootArrangement(restored))

                        Assert.Multiple(
                            Sub()
                                Assert.That(RootDrawing(restored), [Is].SameAs(restoredSurface.Situation))
                                Assert.That(restoredSurface.StructureLines.Count, [Is].EqualTo(1))
                                AssertLine(restoredSurface.StructureLines(0),
                                           New Vector3D(10.25, 20.5, 101.75),
                                           New Vector3D(30.75, 40.125, 99.5),
                                           "Конус К1")
                                Assert.That(RootDrawing(restored).ActiveSpace.Entities.Count, [Is].EqualTo(1))
                                AssertSettings(restoredSettings,
                                               "roads\\A-101.roadx", "terrain\\earth.sfcx")
                            End Sub)
                    Finally
                        DisposeRoot(restored)
                    End Try
                End Using
            Finally
                DisposeRoot(original)
            End Try
        End Sub

        <Test, Explicit("Requires the initialized Topomatic host for legacy Arrangement STG migration.")>
        <Category("TopomaticIntegration")>
        Public Sub LegacyArrangementMigrationPreservesHandlesGroupsXRecordsGeometryAndSettings()
            Dim legacy As New ArrangementModel()
            Dim first As DwgLine = legacy.Drawing.ActiveSpace.AddLine(
                New Vector3D(11, 12, 13), New Vector3D(21, 22, 23))
            Dim second As DwgLine = legacy.Drawing.ActiveSpace.AddLine(
                New Vector3D(-1.25, 2.75, 8.5), New Vector3D(7.125, 6.25, 5.5))
            SetBridgeXRecord(first, "bridge-migration", "axis-1")
            SetBridgeXRecord(second, "bridge-migration", "beam-1")
            Dim bridgeGroup As DwgGroup = legacy.Drawing.Groups.Add("INFRA_BRIDGE_legacy_migration_test")
            first.Group = bridgeGroup
            second.Group = bridgeGroup
            SetExternalSurfaceSettings(legacy, "roads\\design.roadx", "terrain\\ground.sfcx")
            Dim expectedHandles As String() = {first.Handle, second.Handle}

            Dim document As New StgDocument()
            SaveSettings(legacy, document.Body)
            legacy.SaveToStg(document.Body)
            Using stream As New MemoryStream()
                document.SaveToStreamAsBinary(stream)
                stream.Position = 0
                Dim migrated As Object = NewRoot()
                Try
                    InvokeInstance(migrated, "NativeLoadFromStream", stream)
                    Dim drawing As Drawing = RootDrawing(migrated)
                    Dim migratedEntities As DwgEntity() = drawing.ActiveSpace.Entities.
                        Cast(Of DwgEntity)().ToArray()

                    Assert.Multiple(
                        Sub()
                            Assert.That(drawing, [Is].SameAs(RootSurface(migrated).Situation))
                            Assert.That(migratedEntities.Length, [Is].EqualTo(2))
                            Assert.That(migratedEntities.Select(Function(item) item.Handle),
                                        [Is].EqualTo(expectedHandles),
                                        "Legacy Drawing.Assign migration must preserve handles.")
                            Assert.That(migratedEntities.Select(Function(item) item.Group.Name).Distinct(),
                                        [Is].EqualTo(New String() {"INFRA_BRIDGE_legacy_migration_test"}))
                            Assert.That(ReadBridgeIdentity(migratedEntities(0)),
                                        [Is].EqualTo("bridge-migration|axis-1"))
                            Assert.That(ReadBridgeIdentity(migratedEntities(1)),
                                        [Is].EqualTo("bridge-migration|beam-1"))
                            AssertDwgLine(DirectCast(migratedEntities(0), DwgLine),
                                          New Vector3D(11, 12, 13), New Vector3D(21, 22, 23))
                            AssertDwgLine(DirectCast(migratedEntities(1), DwgLine),
                                          New Vector3D(-1.25, 2.75, 8.5), New Vector3D(7.125, 6.25, 5.5))
                            AssertSettings(GetSettings(RootArrangement(migrated)),
                                           "roads\\design.roadx", "terrain\\ground.sfcx")
                        End Sub)
                Finally
                    DisposeRoot(migrated)
                End Try
            End Using
        End Sub

        <Test, Explicit("Requires the initialized Topomatic host for native model persistence.")>
        <Category("TopomaticIntegration")>
        Public Sub ReplacingAHardReferenceMarksTheRootDirtyAndPersistsTheNewPath()
            Const oldPath As String = "roads/old-design.roadx"
            Const newPath As String = "roads/new-design.roadx"
            Dim model As Object = NewRoot()
            Try
                Dim arrangement As ArrangementModel = RootArrangement(model)
                SetExternalSurfaceSettings(arrangement, oldPath, "terrain/earth.sfcx")
                DirectCast(model, StateControllerObject).Modified = False
                Dim editor As Object = Activator.CreateInstance(
                    MainType("CivilEnginStructures.BridgesModelEditor"))

                editor.GetType().GetMethod("ReplaceObjectHardReference").Invoke(
                    editor, New Object() {model, New URI(oldPath), New URI(newPath)})

                Assert.That(DirectCast(model, StateControllerObject).Modified, [Is].True,
                            "Changing root metadata must mark the registered project model dirty.")
                AssertSettings(GetSettings(arrangement), newPath, "terrain/earth.sfcx")

                Using stream As New MemoryStream()
                    InvokeInstance(model, "NativeSaveToStream", stream)
                    stream.Position = 0
                    Dim restored As Object = NewRoot()
                    Try
                        InvokeInstance(restored, "NativeLoadFromStream", stream)
                        AssertSettings(GetSettings(RootArrangement(restored)),
                                       newPath, "terrain/earth.sfcx")
                    Finally
                        DisposeRoot(restored)
                    End Try
                End Using
            Finally
                DisposeRoot(model)
            End Try
        End Sub

        <Test, Explicit("Requires the initialized Topomatic host for native Drawing construction.")>
        <Category("TopomaticIntegration")>
        Public Sub FailedNativeSaveDoesNotTruncateTheExistingModelFile()
            Dim model As Object = NewRoot()
            Dim path As String = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "infrastrada-native-save-" & Guid.NewGuid().ToString("N") & ".infrabridgex")
            Dim sentinel As Byte() = System.Text.Encoding.UTF8.GetBytes("existing-valid-model-sentinel")
            Try
                File.WriteAllBytes(path, sentinel)
                RootArrangement(model).Drawing.ActiveSpace.AddLine(
                    New Vector3D(1, 2, 3), New Vector3D(4, 5, 6))

                Assert.That(
                    Sub() InvokeInstanceUnwrapped(model, "NativeSaveToFile", path),
                    Throws.TypeOf(Of InvalidOperationException)())
                Assert.That(File.ReadAllBytes(path), [Is].EqualTo(sentinel),
                            "Validation must complete before an existing model file is opened with FileMode.Create.")
            Finally
                If File.Exists(path) Then File.Delete(path)
                DisposeRoot(model)
            End Try
        End Sub

        Private Shared Sub AssertProperty(ownerType As Type, propertyName As String, propertyType As Type)
            Dim item As PropertyInfo = ownerType.GetProperty(propertyName, BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(item, [Is].Not.Null, ownerType.FullName & "." & propertyName & " must be public.")
            If item IsNot Nothing Then Assert.That(item.PropertyType, [Is].EqualTo(propertyType))
        End Sub

        Private Shared Function ExtractMethodSource(source As String, signature As String) As String
            Dim startIndex As Integer = source.IndexOf(signature, StringComparison.Ordinal)
            Assert.That(startIndex, [Is].GreaterThanOrEqualTo(0), signature)
            Dim endIndex As Integer = source.IndexOf("End Sub", startIndex, StringComparison.Ordinal)
            Assert.That(endIndex, [Is].GreaterThan(startIndex), signature)
            Return source.Substring(startIndex, endIndex - startIndex)
        End Function

        Private Shared Function PublicInstanceMethod(ownerType As Type,
                                                     name As String,
                                                     ParamArray parameters As Type()) As MethodInfo
            Return ownerType.GetMethod(name, BindingFlags.Public Or BindingFlags.Instance,
                                       Nothing, parameters, Nothing)
        End Function

        Private Shared Function PublicSharedMethod(ownerType As Type,
                                                   name As String,
                                                   ParamArray parameters As Type()) As MethodInfo
            Return ownerType.GetMethod(name, BindingFlags.Public Or BindingFlags.Static,
                                       Nothing, parameters, Nothing)
        End Function

        Private Shared Function NewRoot() As Object
            Return Activator.CreateInstance(MainType(RootTypeName))
        End Function

        Private Shared Function RootArrangement(root As Object) As ArrangementModel
            Return DirectCast(root.GetType().GetProperty("Arrangement").GetValue(root, Nothing), ArrangementModel)
        End Function

        Private Shared Function RootSurface(root As Object) As Surface
            Return DirectCast(root.GetType().GetProperty("Surface").GetValue(root, Nothing), Surface)
        End Function

        Private Shared Function RootDrawing(root As Object) As Drawing
            Return DirectCast(root.GetType().GetProperty("Drawing").GetValue(root, Nothing), Drawing)
        End Function

        Private Shared Sub DisposeRoot(root As Object)
            Dim disposable As IDisposable = TryCast(root, IDisposable)
            If disposable IsNot Nothing Then disposable.Dispose()
        End Sub

        Private Shared Sub InvokeInstance(target As Object, methodName As String, ParamArray arguments As Object())
            target.GetType().GetMethod(methodName, BindingFlags.Public Or BindingFlags.Instance).
                Invoke(target, arguments)
        End Sub

        Private Shared Sub InvokeInstanceUnwrapped(target As Object,
                                                   methodName As String,
                                                   ParamArray arguments As Object())
            Try
                InvokeInstance(target, methodName, arguments)
            Catch ex As TargetInvocationException When ex.InnerException IsNot Nothing
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Sub AddStructureLine(surface As Surface,
                                            first As Vector3D,
                                            second As Vector3D,
                                            description As String)
            Dim firstIndex As Integer = surface.Points.Count
            surface.Points.Add(New SurfacePoint(first))
            Dim secondIndex As Integer = surface.Points.Count
            surface.Points.Add(New SurfacePoint(second))
            Dim line As New StructureLine() With {.Description = description}
            line.Add(firstIndex)
            line.Add(secondIndex)
            surface.StructureLines.Add(line)
        End Sub

        Private Shared Sub AssertLine(line As StructureLine,
                                     expectedFirst As Vector3D,
                                     expectedSecond As Vector3D,
                                     expectedDescription As String)
            Assert.Multiple(
                Sub()
                    AssertVector(line.GetPosition(0), expectedFirst)
                    AssertVector(line.GetPosition(1), expectedSecond)
                    Assert.That(line.Description, [Is].EqualTo(expectedDescription))
                End Sub)
        End Sub

        Private Shared Sub AssertDwgLine(line As DwgLine,
                                        expectedStart As Vector3D,
                                        expectedEnd As Vector3D)
            Assert.Multiple(
                Sub()
                    AssertVector(line.StartPoint, expectedStart)
                    AssertVector(line.EndPoint, expectedEnd)
                End Sub)
        End Sub

        Private Shared Sub AssertVector(actual As Vector3D, expected As Vector3D)
            Assert.Multiple(
                Sub()
                    Assert.That(actual.X, [Is].EqualTo(expected.X).Within(0.0000001))
                    Assert.That(actual.Y, [Is].EqualTo(expected.Y).Within(0.0000001))
                    Assert.That(actual.Z, [Is].EqualTo(expected.Z).Within(0.0000001))
                End Sub)
        End Sub

        Private Shared Sub SetBridgeXRecord(entity As DwgEntity, bridgeId As String, elementId As String)
            Dim assembly As Assembly = MainAssembly()
            Dim dataType As Type = assembly.GetType("CivilEnginStructures.StructureElement", True)
            Dim data As Object = Activator.CreateInstance(dataType)
            dataType.GetProperty("IdStructure").SetValue(data, bridgeId, Nothing)
            dataType.GetProperty("IdElement").SetValue(data, elementId, Nothing)
            Dim nameProperty As PropertyInfo = dataType.GetProperty("Name")
            nameProperty.SetValue(data, [Enum].Parse(nameProperty.PropertyType, "axisBridge"), Nothing)
            Dim tableType As Type = dataType.GetNestedType("tableXRecords")
            Dim table As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim funcType As Type = assembly.GetType("CivilEnginStructures.FuncXRecords", True)
            Dim setMethod As MethodInfo = funcType.GetMethod("setXRecords", BindingFlags.Public Or BindingFlags.Static)

            Assert.That(CBool(setMethod.Invoke(Nothing, New Object() {entity, table, data, String.Empty})), [Is].True)
        End Sub

        Private Shared Function ReadBridgeIdentity(entity As DwgEntity) As String
            Dim assembly As Assembly = MainAssembly()
            Dim dataType As Type = assembly.GetType("CivilEnginStructures.StructureElement", True)
            Dim data As Object = Activator.CreateInstance(dataType)
            Dim tableType As Type = dataType.GetNestedType("tableXRecords")
            Dim table As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim funcType As Type = assembly.GetType("CivilEnginStructures.FuncXRecords", True)
            Dim getMethod As MethodInfo = funcType.GetMethod("getXRecords", BindingFlags.Public Or BindingFlags.Static)
            Dim arguments As Object() = {entity, data, table}

            Assert.That(CBool(getMethod.Invoke(Nothing, arguments)), [Is].True)
            data = arguments(1)
            Return CStr(dataType.GetProperty("IdStructure").GetValue(data, Nothing)) & "|" &
                   CStr(dataType.GetProperty("IdElement").GetValue(data, Nothing))
        End Function

        Private Shared Function GetSettings(arrangement As ArrangementModel) As Object
            Dim storeType As Type = MainType("CivilEnginStructures.BridgeModelSettingsStore")
            Return storeType.GetMethod("GetSettings", BindingFlags.Public Or BindingFlags.Static).
                Invoke(Nothing, New Object() {arrangement})
        End Function

        Private Shared Sub SetExternalSurfaceSettings(arrangement As ArrangementModel,
                                                     projectPath As String,
                                                     earthPath As String)
            Dim settings As Object = GetSettings(arrangement)
            settings.GetType().GetProperty("ProjectSurfaceRelativePath").SetValue(settings, projectPath, Nothing)
            settings.GetType().GetProperty("EarthSurfaceRelativePath").SetValue(settings, earthPath, Nothing)
        End Sub

        Private Shared Sub AssertSettings(settings As Object,
                                          expectedProjectPath As String,
                                          expectedEarthPath As String)
            Assert.Multiple(
                Sub()
                    Assert.That(CStr(settings.GetType().GetProperty("ProjectSurfaceRelativePath").
                                     GetValue(settings, Nothing)), [Is].EqualTo(expectedProjectPath))
                    Assert.That(CStr(settings.GetType().GetProperty("EarthSurfaceRelativePath").
                                     GetValue(settings, Nothing)), [Is].EqualTo(expectedEarthPath))
                End Sub)
        End Sub

        Private Shared Sub SaveSettings(arrangement As ArrangementModel, body As StgNode)
            Dim storeType As Type = MainType("CivilEnginStructures.BridgeModelSettingsStore")
            storeType.GetMethod("Save", BindingFlags.Public Or BindingFlags.Static).
                Invoke(Nothing, New Object() {arrangement, body})
        End Sub

        Private Shared Function MainType(fullName As String) As Type
            Dim result As Type = MainAssembly().GetType(fullName, False)
            Assert.That(result, [Is].Not.Null, fullName & " must exist in the production assembly.")
            Return result
        End Function

        Private Shared Function MainAssembly() As Assembly
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
                        "Build CivilEnginStructures.vbproj before running the native bridge tests.")
            Return Assembly.LoadFrom(assemblyPath)
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
