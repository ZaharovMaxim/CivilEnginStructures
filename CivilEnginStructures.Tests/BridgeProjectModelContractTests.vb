Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json.Linq
Imports NUnit.Framework
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Stg

Namespace Tests
    <TestFixture>
    Public Class BridgeProjectModelContractTests
        Private Const BridgeModelType As String = "infrastrada_bridges"
        Private Const BridgeExtension As String = ".infrabridgex"

        <Test, Explicit("Requires the initialized Topomatic application environment (font/config paths).")>
        <Category("TopomaticIntegration")>
        Public Sub ArrangementModelNativeStgRoundTripPreservesDrawingAndSurfaceReferences()
            Dim original As New ArrangementModel()
            original.ProjectSurfacesRelativePaths.Add("surfaces/design.sfcx")
            original.ProjectSurfacesRelativePaths.Add("surfaces/earth.sfcx")

            Dim document As New StgDocument()
            original.SaveToStg(document.Body)

            Using stream As New MemoryStream()
                document.SaveToStreamAsBinary(stream)
                stream.Position = 0

                Dim loadedDocument As New StgDocument()
                loadedDocument.LoadFromStreamAsBinary(stream)
                Dim restored As New ArrangementModel()
                restored.LoadFromStg(loadedDocument.Body)

                Assert.Multiple(
                    Sub()
                        Assert.That(restored.Drawing, [Is].Not.Null)
                        Assert.That(restored.ProjectSurfacesRelativePaths,
                                    [Is].EqualTo(New String() {
                                        "surfaces/design.sfcx",
                                        "surfaces/earth.sfcx"
                                    }))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub ExistingProjectBridgePublicModelTypeRemainsArrangementModel()
            Dim projectBridge As Type = MainType("CivilEnginStructures.ProjectBridge")
            Dim modelProperty As PropertyInfo = projectBridge.GetProperty("BridgeModel")

            Assert.That(modelProperty, [Is].Not.Null)
            Assert.That(modelProperty.PropertyType, [Is].EqualTo(GetType(ArrangementModel)))
            Assert.That(modelProperty.CanRead, [Is].True)
            Assert.That(modelProperty.CanWrite, [Is].True)
        End Sub

        <Test>
        Public Sub IndependentBridgeEditorUsesPlanModelContractWithoutReplacingTutorialEditor()
            Dim tutorialEditor As Type = MainType("CivilEnginStructures.EditorBridgeModel")
            Dim editor As Type = MainType("CivilEnginStructures.BridgesModelEditor")

            Assert.Multiple(
                Sub()
                    Assert.That(tutorialEditor, [Is].Not.Null,
                                "The existing tutorial bridge/.bridgex editor must remain available.")
                    Assert.That(editor.BaseType.FullName,
                                [Is].EqualTo("Topomatic.ApplicationPlatform.Core.PlanModelEditor"))
                    Assert.That(editor.GetMethod("LoadFromFile").ReturnType,
                                [Is].EqualTo(GetType(Object)))
                    Assert.That(editor.GetMethod("SaveToFile"), [Is].Not.Null)
                    Assert.That(editor.GetMethod("GetHardReferences").DeclaringType,
                                [Is].EqualTo(editor))
                    Assert.That(editor.GetMethod("ReplaceObjectHardReference").DeclaringType,
                                [Is].EqualTo(editor))
                End Sub)
        End Sub

        <Test>
        Public Sub PluginInitializationRegistersOnlyTheIndependentBridgeModelType()
            Dim initializerType As Type = MainType("CivilEnginStructures.RopExample1.RopExample1Module")
            Dim initializer = DirectCast(Activator.CreateInstance(initializerType), PluginInitializator)
            Dim factory As New RecordingPluginFactory()

            initializer.Initialize(factory)

            Assert.Multiple(
                Sub()
                    Assert.That(factory.Editors.Count, [Is].EqualTo(1))
                    Assert.That(factory.Editors.ContainsKey("bridge"), [Is].False,
                                "The tutorial class remains in source but its model type must not be registered.")
                    Assert.That(factory.Editors.ContainsKey(BridgeModelType), [Is].True)
                End Sub)
            If factory.Editors.ContainsKey(BridgeModelType) Then
                Assert.That(factory.Editors(BridgeModelType).Extension, [Is].EqualTo(BridgeExtension))
            End If
        End Sub

        <Test>
        Public Sub BridgeFeatureSourceContractsKeepLegacyDiscoveryAndUseTheNewModel()
            Dim root As String = RepositoryRoot()
            Dim projectSource As String = File.ReadAllText(
                Path.Combine(root, "CivilEngineeringStructures", "ProjectCivilStructures.vb"))
            Dim formSource As String = File.ReadAllText(
                Path.Combine(root, "UserForms", "Bridge", "FormPlacementBeams.vb"))

            Assert.Multiple(
                Sub()
                    Assert.That(projectSource, Does.Contain(".arrx"),
                                "Legacy arrangement discovery must remain supported.")
                    Assert.That(projectSource, Does.Contain(BridgeExtension),
                                "The independent bridge model must be discovered.")
                    Assert.That(projectSource, Does.Contain("createArrangementModel"),
                                "The existing arrangement factory is a compatibility contract.")
                    Assert.That(projectSource, Does.Contain("createBridgeModel"),
                                "The independent bridge factory must be available.")
                    Assert.That(formSource, Does.Contain(".createBridgeModel()"),
                                "The create button must create the independent bridge model.")
                End Sub)
        End Sub

        <Test>
        Public Sub SettingsLoadFromMissingFieldsUsesEmptyDefaults()
            Dim settings As Object = NewSettings()
            Dim document As New StgDocument()

            InvokeSettings(settings, "LoadFromStg", document.Body)

            Assert.Multiple(
                Sub()
                    Assert.That(GetSetting(settings, "ProjectSurfaceRelativePath"), [Is].EqualTo(String.Empty))
                    Assert.That(GetSetting(settings, "EarthSurfaceRelativePath"), [Is].EqualTo(String.Empty))
                End Sub)
        End Sub

        <Test>
        Public Sub SettingsStgBinaryRoundTripPreservesUnicodeRelativePaths()
            Dim original As Object = NewSettings()
            SetSetting(original, "ProjectSurfaceRelativePath", "..\Поверхности\Проектная.sfcx")
            SetSetting(original, "EarthSurfaceRelativePath", "ЦММ\Существующая земля.sfcx")
            Dim document As New StgDocument()
            InvokeSettings(original, "SaveToStg", document.Body)

            Using stream As New MemoryStream()
                document.SaveToStreamAsBinary(stream)
                stream.Position = 0
                Dim loadedDocument As New StgDocument()
                loadedDocument.LoadFromStreamAsBinary(stream)
                Dim restored As Object = NewSettings()

                InvokeSettings(restored, "LoadFromStg", loadedDocument.Body)

                Assert.Multiple(
                    Sub()
                        Assert.That(GetSetting(restored, "ProjectSurfaceRelativePath"),
                                    [Is].EqualTo("..\Поверхности\Проектная.sfcx"))
                        Assert.That(GetSetting(restored, "EarthSurfaceRelativePath"),
                                    [Is].EqualTo("ЦММ\Существующая земля.sfcx"))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub ReplaceReferenceUpdatesBothSurfaceRolesWhenTheyShareAPath()
            Dim settings As Object = NewSettings()
            SetSetting(settings, "ProjectSurfaceRelativePath", "surfaces\common.sfcx")
            SetSetting(settings, "EarthSurfaceRelativePath", "surfaces\common.sfcx")

            InvokeSettings(settings, "ReplaceReference", "surfaces\common.sfcx", "surfaces\renamed.sfcx")

            Assert.Multiple(
                Sub()
                    Assert.That(GetSetting(settings, "ProjectSurfaceRelativePath"),
                                [Is].EqualTo("surfaces\renamed.sfcx"))
                    Assert.That(GetSetting(settings, "EarthSurfaceRelativePath"),
                                [Is].EqualTo("surfaces\renamed.sfcx"))
                End Sub)
        End Sub

        <Test>
        Public Sub ReplaceReferenceLeavesUnrelatedSurfaceRolesUnchanged()
            Dim settings As Object = NewSettings()
            SetSetting(settings, "ProjectSurfaceRelativePath", "surfaces\design.sfcx")
            SetSetting(settings, "EarthSurfaceRelativePath", "surfaces\earth.sfcx")

            InvokeSettings(settings, "ReplaceReference", "surfaces\other.sfcx", "surfaces\renamed.sfcx")

            Assert.Multiple(
                Sub()
                    Assert.That(GetSetting(settings, "ProjectSurfaceRelativePath"),
                                [Is].EqualTo("surfaces\design.sfcx"))
                    Assert.That(GetSetting(settings, "EarthSurfaceRelativePath"),
                                [Is].EqualTo("surfaces\earth.sfcx"))
                End Sub)
        End Sub

        <TestCase("models\Road\Shared.roadx", False)>
        <TestCase("models\Road\Shared.roadx", True)>
        <TestCase("models\Terrain\Shared.sfcx", False)>
        <TestCase("models\Terrain\Shared.sfcx", True)>
        Public Sub SurfaceChoiceUsesExactRelativePathForDuplicateNamesRegardlessOfOrdering(
            storedPath As String,
            reverseOrder As Boolean)

            Dim road As Object = NewSurfaceChoice("Shared", "RoadModel", "models\Road\Shared.roadx")
            Dim terrain As Object = NewSurfaceChoice("Shared", "TerrainModel", "models\Terrain\Shared.sfcx")
            Dim ordered As Object() = If(reverseOrder,
                                         New Object() {terrain, road},
                                         New Object() {road, terrain})
            Dim expected As Object = If(storedPath.EndsWith(".roadx", StringComparison.Ordinal), road, terrain)

            Dim selected As Object = FindSurfaceChoice(ordered, storedPath)

            Assert.That(selected, [Is].SameAs(expected))
        End Sub

        <Test>
        Public Sub SurfaceChoiceSupportsLegacyBareModelName()
            Dim road As Object = NewSurfaceChoice("Проектная поверхность", "RoadModel", "models\Road\Design.roadx")
            Dim terrain As Object = NewSurfaceChoice("Существующая земля", "TerrainModel", "models\Terrain\Earth.sfcx")

            Dim selected As Object = FindSurfaceChoice(New Object() {road, terrain}, "Существующая земля")

            Assert.That(selected, [Is].SameAs(terrain))
        End Sub

        <TestCase("models\Terrain\Missing.sfcx")>
        <TestCase("")>
        <TestCase(Nothing)>
        Public Sub SurfaceChoiceReturnsNothingWhenStoredValueDoesNotMatch(storedValue As String)
            Dim road As Object = NewSurfaceChoice("Проектная поверхность", "RoadModel", "models\Road\Design.roadx")

            Assert.That(FindSurfaceChoice(New Object() {road}, storedValue), [Is].Null)
        End Sub

        <Test>
        Public Sub PluginManifestExposesBridgeCreationSettingsCoreAndRibbonTasks()
            Dim manifestPath As String = Path.Combine(RepositoryRoot(), "InfrastradaBridgeTools.plugin")
            Dim manifest As JObject = JObject.Parse(File.ReadAllText(manifestPath))
            Dim assemblyEntry As String = CStr(manifest("assemblies")("InfrastradaRoadTools")("assembly"))
            Dim createAction As JObject = DirectCast(manifest("actions")("id_create_infrastrada_bridges"), JObject)
            Dim settingsAction As JObject = DirectCast(manifest("actions")("id_edit_infrastrada_bridge_settings"), JObject)
            Dim core As JObject = DirectCast(manifest("cores")(BridgeModelType), JObject)
            Dim createIcon As String = CStr(createAction("icon"))
            Dim ribbonItems As JArray = DirectCast(manifest("ribbon")("rbproj")("items"), JArray)

            Assert.Multiple(
                Sub()
                    Assert.That(assemblyEntry,
                                [Is].EqualTo("CivilEnginStructures.dll, CivilEnginStructures.RopExample1.RopExample1PluginHost"))
                    Assert.That(CStr(createAction("cmd")), Does.Contain(BridgeModelType))
                    Assert.That(CStr(settingsAction("cmd")), Does.Contain("edit_infrastrada_bridge_settings"))
                    Assert.That(createIcon, [Is].EqualTo("infrastrada_bridge_model"))
                    Assert.That(CStr(settingsAction("icon")), [Is].EqualTo(createIcon))
                    Assert.That(CStr(core("icon")), [Is].EqualTo(createIcon))
                    Assert.That(CStr(core("flags")), [Is].EqualTo("$(modelflags,%0)"))
                    Assert.That(CStr(manifest("core.items")(BridgeModelType)),
                                [Is].EqualTo("core." & BridgeModelType))
                    Assert.That(manifest("contexts")("ctx_mkitem")("items").
                                    Any(Function(item) item.Type = JTokenType.String AndAlso
                                                       CStr(item).Contains("id_create_infrastrada_bridges")),
                                [Is].True)
                    Assert.That(manifest("contexts")("core." & BridgeModelType)("items").
                                    Any(Function(item) item.Type = JTokenType.String AndAlso
                                                       CStr(item).Contains("id_edit_infrastrada_bridge_settings")),
                                [Is].True)
                    Assert.That(ribbonItems.Children(Of JObject)().
                                    Any(Function(item) CStr(item("group")) = "brep_infrastrada_bridges" AndAlso
                                                       CStr(item("flags")) = "$(infrastrada_bridges)"),
                                [Is].True)
                    Assert.That(ribbonItems.Children(Of JObject)().
                                    Any(Function(item) CStr(item("group")) = "task_infrastrada_bridges" AndAlso
                                                       CStr(item("flags")) = "$(infrastrada_bridges)"),
                                [Is].True)
                    Assert.That(manifest("ribbon")("rbproj.brep_infrastrada_bridges"), [Is].Not.Null)
                    Assert.That(manifest("ribbon")("rbproj.task_infrastrada_bridges"), [Is].Not.Null)
                End Sub)
        End Sub

        <TestCase("infrastrada_bridge_model_16dp_1x.png", 16)>
        <TestCase("infrastrada_bridge_model_32dp_1x.png", 32)>
        Public Sub BridgeModelPngIconsHaveExpectedDimensionsAndAlpha(fileName As String, expectedSize As Integer)
            Dim iconPath As String = Path.Combine(RepositoryRoot(), "Icons", fileName)
            Assert.That(File.Exists(iconPath), [Is].True, fileName & " must be included in the repository.")

            Using bitmap As New System.Drawing.Bitmap(iconPath)
                Assert.Multiple(
                    Sub()
                        Assert.That(bitmap.Width, [Is].EqualTo(expectedSize))
                        Assert.That(bitmap.Height, [Is].EqualTo(expectedSize))
                        Assert.That(System.Drawing.Image.IsAlphaPixelFormat(bitmap.PixelFormat), [Is].True)
                        Assert.That(HasTransparentPixel(bitmap), [Is].True)
                    End Sub)
            End Using
            Assert.That(File.Exists(Path.Combine(RepositoryRoot(), "Icons", "infrastrada_bridge_model.svg")),
                        [Is].True)
        End Sub

        Private Shared Function NewSurfaceChoice(modelName As String,
                                                 modelType As String,
                                                 relativePath As String) As Object
            Dim choiceType As Type = MainType("CivilEnginStructures.SurfaceModelChoice")
            Return Activator.CreateInstance(choiceType, New Object() {modelName, modelType, relativePath})
        End Function

        Private Shared Function FindSurfaceChoice(choices As Object(), storedValue As String) As Object
            Dim choiceType As Type = MainType("CivilEnginStructures.SurfaceModelChoice")
            Dim listType As Type = GetType(List(Of )).MakeGenericType(choiceType)
            Dim typedChoices As Object = Activator.CreateInstance(listType)
            Dim addMethod As MethodInfo = listType.GetMethod("Add")
            For Each choice As Object In choices
                addMethod.Invoke(typedChoices, New Object() {choice})
            Next
            Return choiceType.GetMethod("FindByStoredValue").
                Invoke(Nothing, New Object() {typedChoices, storedValue})
        End Function

        Private Shared Function HasTransparentPixel(bitmap As System.Drawing.Bitmap) As Boolean
            For y As Integer = 0 To bitmap.Height - 1
                For x As Integer = 0 To bitmap.Width - 1
                    If bitmap.GetPixel(x, y).A < Byte.MaxValue Then Return True
                Next
            Next
            Return False
        End Function

        Private Shared Function NewSettings() As Object
            Return Activator.CreateInstance(MainType("CivilEnginStructures.BridgeModelSettings"))
        End Function

        Private Shared Function GetSetting(settings As Object, propertyName As String) As String
            Return CStr(settings.GetType().GetProperty(propertyName).GetValue(settings, Nothing))
        End Function

        Private Shared Sub SetSetting(settings As Object, propertyName As String, value As String)
            settings.GetType().GetProperty(propertyName).SetValue(settings, value, Nothing)
        End Sub

        Private Shared Sub InvokeSettings(settings As Object, methodName As String, ParamArray arguments As Object())
            settings.GetType().GetMethod(methodName).Invoke(settings, arguments)
        End Sub

        Private Shared Function MainType(fullName As String) As Type
            Dim type As Type = MainAssembly().GetType(fullName, False)
            Assert.That(type, [Is].Not.Null, fullName & " must exist in the production assembly.")
            Return type
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
                        "Build CivilEnginStructures.vbproj before running the feature contract tests.")
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

        Private NotInheritable Class RecordingPluginFactory
            Inherits PluginFactory

            Friend ReadOnly Editors As New Dictionary(Of String, ModelEditorInfo)(StringComparer.Ordinal)

            Public Overrides Sub RegisterFunction(name As String, func As PluginFunction)
            End Sub

            Public Overrides Sub RegisterType(name As String, type As Type)
            End Sub

            Public Overrides Sub RegisterLoader(name As String, loader As String)
            End Sub

            Public Overrides Sub RegisterModelEditor(name As String, editor As ModelEditorInfo)
                Editors.Add(name, editor)
            End Sub

            Public Overrides Sub RegisterProjectSettings(name As String, settings As String)
            End Sub

            Public Overrides Sub RegisterTask(name As String, task As String)
            End Sub
        End Class
    End Class
End Namespace
