Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class BridgeRuntimeActivationTests
        Private Shared _mainAssembly As Assembly
        Private Shared _assemblyResolveHandler As ResolveEventHandler

        <OneTimeSetUp>
        Public Sub RegisterTopomaticSdkDependencyResolver()
            _assemblyResolveHandler = AddressOf ResolveTopomaticSdkAssembly
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, _assemblyResolveHandler
        End Sub

        <OneTimeTearDown>
        Public Sub UnregisterTopomaticSdkDependencyResolver()
            If _assemblyResolveHandler Is Nothing Then Return
            RemoveHandler AppDomain.CurrentDomain.AssemblyResolve, _assemblyResolveHandler
            _assemblyResolveHandler = Nothing
        End Sub

        <Test>
        Public Sub ConfigureComboSelectsUnassignedChoiceBeforeComboIsParented()
            Using combo As New ComboBox()
                Dim choices As IList = CreateChoices(
                    Choice("(не назначена)", String.Empty))

                Assert.DoesNotThrow(
                    Sub() ConfigureCombo(combo, choices, String.Empty),
                    "The settings dialog configures its ComboBox fields before adding them to the Form.")
                Assert.Multiple(
                    Sub()
                        Assert.That(combo.Items.Count, [Is].EqualTo(1))
                        Assert.That(combo.SelectedIndex, [Is].EqualTo(0))
                        Assert.That(SelectedRelativePath(combo), [Is].EqualTo(String.Empty))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub ConfigureComboSelectsAvailableSavedSurfaceBeforeComboIsParented()
            Using combo As New ComboBox()
                Dim choices As IList = CreateChoices(
                    Choice("(не назначена)", String.Empty),
                    Choice("Проектная [Поверхность] — surfaces/project.sfcx", "surfaces/project.sfcx"))

                Assert.DoesNotThrow(
                    Sub() ConfigureCombo(combo, choices, "SURFACES/project.sfcx"))
                Assert.Multiple(
                    Sub()
                        Assert.That(combo.Items.Count, [Is].EqualTo(2))
                        Assert.That(combo.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(SelectedRelativePath(combo), [Is].EqualTo("surfaces/project.sfcx"))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub ConfigureComboPreservesUnavailableSavedSurfaceBeforeComboIsParented()
            Using combo As New ComboBox()
                Dim missingPath As String = "surfaces/missing.sfcx"
                Dim choices As IList = CreateChoices(
                    Choice("(не назначена)", String.Empty))

                Assert.DoesNotThrow(
                    Sub() ConfigureCombo(combo, choices, missingPath))
                Assert.Multiple(
                    Sub()
                        Assert.That(combo.Items.Count, [Is].EqualTo(2))
                        Assert.That(combo.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(combo.GetItemText(combo.SelectedItem),
                                    [Is].EqualTo("(недоступна) " & missingPath))
                        Assert.That(SelectedRelativePath(combo), [Is].EqualTo(missingPath))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverFindsDirectBridgePlanLayer()
            Dim plan As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")

            Assert.That(ResolveBridgePlanLayer(plan), [Is].SameAs(plan))
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverTracksTheActiveBridgeAmongMultipleOpenModels()
            Dim models As New TestModelsLayer()
            Dim first As CadViewLayer = CreateBridgePlan("models/first.infrabridgex")
            Dim second As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")
            models.Add(first)
            models.Add(second)

            models.ActiveLayer = second
            Assert.That(ResolveBridgePlanLayer(models), [Is].SameAs(second),
                        "The resolver must not use the first visible bridge model.")

            models.ActiveLayer = first
            Assert.That(ResolveBridgePlanLayer(models), [Is].SameAs(first),
                        "Changing the active model must change the bridge command context.")
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverIgnoresVisibleBridgeWhenAnotherModelIsActive()
            Dim models As New TestModelsLayer()
            Dim bridge As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")
            Dim other As New TestNonBridgeLayer()
            models.Add(bridge)
            models.Add(other)
            models.ActiveLayer = other

            Assert.That(ResolveBridgePlanLayer(models), [Is].Null,
                        "A visible inactive bridge must not receive commands for another active model.")
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverFindsBridgeInsideActiveModelCompoundLayer()
            Dim models As New TestModelsLayer()
            Dim modelWrapper As New CompoundLayer(
                "Model 78", New Guid("{BF00D5ED-A7C4-4BE1-A315-54D479A9F9AC}"))
            Dim bridge As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")
            modelWrapper.Add(bridge)
            models.Add(modelWrapper)
            models.ActiveLayer = modelWrapper

            Assert.That(ResolveBridgePlanLayer(models), [Is].SameAs(bridge))
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverRejectsDisabledBridgeLayer()
            Dim bridge As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")
            bridge.Enable = False

            Assert.That(ResolveBridgePlanLayer(bridge), [Is].Null)
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverReturnsNothingForMissingRootLayer()
            Assert.That(ResolveBridgePlanLayer(Nothing), [Is].Null)
        End Sub

        <Test>
        Public Sub BridgePlanContextResolverReturnsNothingForMissingCadView()
            Assert.That(ResolveBridgePlan(Nothing), [Is].Null)
        End Sub

        <Test, Explicit("Requires the initialized Topomatic CAD host; standalone New CadView terminates the test host.")>
        <Category("TopomaticIntegration")>
        Public Sub BridgePlanContextResolverFindsPlanThroughActualCadView()
            Dim cadView As New CadView()
            Dim plan As CadViewLayer = CreateBridgePlan("models/78.infrabridgex")
            cadView.AddLayer(plan)

            Assert.That(ResolveBridgePlan(cadView), [Is].SameAs(plan))
        End Sub

        <Test>
        Public Sub BridgePlanLayerExposesItsCommandDrawingLayerAndExactModelPath()
            Dim modelPath As String = "models/78.infrabridgex"
            Dim plan As CadViewLayer = CreateBridgePlan(modelPath)
            Dim drawingLayerProperty As PropertyInfo = PlanType().GetProperty(
                "DrawingLayer",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            Dim modelPathProperty As PropertyInfo = PlanType().GetProperty(
                "ModelPathId",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

            Assert.Multiple(
                Sub()
                    Assert.That(drawingLayerProperty, [Is].Not.Null)
                    If drawingLayerProperty IsNot Nothing Then
                        Dim drawingLayer As Object = drawingLayerProperty.GetValue(plan, Nothing)
                        Assert.That(drawingLayer, [Is].Not.Null)
                        If drawingLayer IsNot Nothing Then
                            Assert.That(drawingLayer.GetType().FullName,
                                        [Is].EqualTo("CivilEnginStructures.BridgeDrawingLayer"))
                        End If
                    End If
                    Assert.That(modelPathProperty, [Is].Not.Null)
                    If modelPathProperty IsNot Nothing Then
                        Assert.That(CStr(modelPathProperty.GetValue(plan, Nothing)), [Is].EqualTo(modelPath))
                    End If
                End Sub)
        End Sub

        Private Shared Sub ConfigureCombo(combo As ComboBox,
                                          choices As IList,
                                          selectedPath As String)
            Dim method As MethodInfo = DialogType().GetMethod(
                "ConfigureCombo",
                BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            method.Invoke(Nothing, New Object() {combo, choices, selectedPath, 15, 30})
        End Sub

        Private Shared Function CreateChoices(ParamArray choices As Object()) As IList
            Dim listType As Type = GetType(List(Of )).MakeGenericType(ChoiceType())
            Dim result As IList = DirectCast(Activator.CreateInstance(listType), IList)
            For Each choice As Object In choices
                result.Add(choice)
            Next
            Return result
        End Function

        Private Shared Function Choice(displayName As String, relativePath As String) As Object
            Return Activator.CreateInstance(
                ChoiceType(),
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                args:=New Object() {displayName, relativePath},
                culture:=Nothing)
        End Function

        Private Shared Function SelectedRelativePath(combo As ComboBox) As String
            Assert.That(combo.SelectedItem, [Is].Not.Null)
            Return CStr(ChoiceType().GetProperty("RelativePath").GetValue(combo.SelectedItem, Nothing))
        End Function

        Private Shared Function ChoiceType() As Type
            Dim result As Type = DialogType().GetNestedType(
                "SurfaceReferenceChoice",
                BindingFlags.NonPublic)
            Assert.That(result, [Is].Not.Null)
            Return result
        End Function

        Private Shared Function DialogType() As Type
            Dim result As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgeModelSettingsDialog",
                throwOnError:=False)
            Assert.That(result, [Is].Not.Null)
            Return result
        End Function

        Private Shared Function CreateBridgePlan(modelPath As String) As CadViewLayer
            Return DirectCast(Activator.CreateInstance(PlanType(), New Object() {modelPath}), CadViewLayer)
        End Function

        Private Shared Function ResolveBridgePlan(cadView As CadView) As Object
            Dim resolverType As Type = BridgePlanResolverType()
            If resolverType Is Nothing Then Return Nothing
            Dim method As MethodInfo = resolverType.GetMethod(
                "Resolve",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {GetType(CadView)},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null)
            If method Is Nothing Then Return Nothing
            Return method.Invoke(Nothing, New Object() {cadView})
        End Function

        Private Shared Function ResolveBridgePlanLayer(rootLayer As CadViewLayer) As Object
            Dim resolverType As Type = BridgePlanResolverType()
            If resolverType Is Nothing Then Return Nothing
            Dim method As MethodInfo = resolverType.GetMethod(
                "ResolveLayer",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {GetType(CadViewLayer)},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "The CadView entry point and offline tests must share one active-layer resolver.")
            If method Is Nothing Then Return Nothing
            Return method.Invoke(Nothing, New Object() {rootLayer})
        End Function

        Private Shared Function BridgePlanResolverType() As Type
            Dim resolverType As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgePlanContextResolver",
                throwOnError:=False)
            Assert.That(resolverType, [Is].Not.Null,
                        "Bridge commands need an explicit resolver for the active compound bridge layer.")
            Return resolverType
        End Function

        Private Shared Function PlanType() As Type
            Dim result As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgePlanCompoundLayer",
                throwOnError:=False)
            Assert.That(result, [Is].Not.Null)
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
                        "Build CivilEnginStructures.vbproj before running activation tests.")
            _mainAssembly = Assembly.LoadFrom(assemblyPath)
            Return _mainAssembly
        End Function

        Private Shared Function ResolveTopomaticSdkAssembly(sender As Object,
                                                             args As ResolveEventArgs) As Assembly
            Dim assemblyName As New AssemblyName(args.Name)
            If Not assemblyName.Name.StartsWith("Topomatic.", StringComparison.OrdinalIgnoreCase) Then
                Return Nothing
            End If
            Dim candidate As String = Path.Combine(
                "C:\Program Files\Topomatic Robur Road 16.0",
                assemblyName.Name & ".dll")
            If Not File.Exists(candidate) Then Return Nothing
            Return Assembly.LoadFrom(candidate)
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

        Private NotInheritable Class TestModelsLayer
            Inherits MultiLayer

            Public Overrides ReadOnly Property LayerGuid As Guid
                Get
                    Return Consts.ModelsLayer
                End Get
            End Property

            Public Overrides ReadOnly Property Name As String
                Get
                    Return "Test models"
                End Get
            End Property
        End Class

        Private NotInheritable Class TestNonBridgeLayer
            Inherits CadViewLayer

            Private Shared ReadOnly TestGuid As New Guid("{E74E3E9D-8AC5-4276-A28F-81AEF7DE40C0}")
            Private ReadOnly _selectionSet As New DefaultSelectionSet(Me)

            Public Overrides ReadOnly Property LayerGuid As Guid
                Get
                    Return TestGuid
                End Get
            End Property

            Public Overrides ReadOnly Property SelectionSet As SelectionSet
                Get
                    Return _selectionSet
                End Get
            End Property

            Public Overrides ReadOnly Property Name As String
                Get
                    Return "Non-bridge model"
                End Get
            End Property

            Protected Overrides Sub OnPaint(pen As CadPen)
            End Sub

            Protected Overrides Function OnGetLimits(ByRef limits As BoundingBox2D) As Boolean
                limits = BoundingBox2D.Empty
                Return False
            End Function

            Protected Overrides Sub OnGetSnapObjects(e As ObjectSnapEventArgs)
            End Sub
        End Class
    End Class
End Namespace
