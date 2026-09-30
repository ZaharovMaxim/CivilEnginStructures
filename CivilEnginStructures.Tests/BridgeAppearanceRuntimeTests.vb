Imports System.Drawing
Imports System.IO
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime

Namespace Tests
    <TestFixture>
    Public Class BridgeAppearanceRuntimeTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub CloneElementForAppearanceCreatesIndependentStaticSolid()
            Dim shell As Shell = Tools.Polygon(
                New Vector3D() {
                    New Vector3D(0, 0, 0),
                    New Vector3D(2, 0, 0),
                    New Vector3D(0, 1, 0)
                })
            shell = Tools.Extrude(0.5, shell)
            Dim source As New StaticSolidElement(
                "Тестовое тело", "SmdxElement", New ImProperties(), shell, New ImDocuments())
            source.Color = Color.FromArgb(255, 20, 40, 60)

            Dim runtimeType As Type = MainType("CivilEnginStructures.BridgeAppearanceNativeRuntime")
            Dim cloneMethod As MethodInfo = runtimeType.GetMethod(
                "CloneElementForAppearance", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(cloneMethod, [Is].Not.Null)
            Dim clone As StaticSolidElement = DirectCast(
                cloneMethod.Invoke(Nothing, New Object() {source}), StaticSolidElement)

            Assert.Multiple(
                Sub()
                    Assert.That(clone, [Is].Not.SameAs(source))
                    Assert.That(clone.Color, [Is].EqualTo(source.Color))
                End Sub)

            clone.Color = Color.FromArgb(255, 180, 120, 30)

            Assert.Multiple(
                Sub()
                    Assert.That(clone.Color, [Is].Not.EqualTo(source.Color))
                    Assert.That(source.Color, [Is].EqualTo(Color.FromArgb(255, 20, 40, 60)),
                                "Changing the draft clone must not recolor the live source element.")
                End Sub)
        End Sub

        <Test>
        Public Sub UniformSavedBodyColorAppliesToAllNewSolidsWhenSolidCountChanges()
            Dim shell As Shell = Tools.Polygon(
                New Vector3D() {
                    New Vector3D(0, 0, 0),
                    New Vector3D(2, 0, 0),
                    New Vector3D(0, 1, 0)
                })
            shell = Tools.Extrude(0.5, shell)
            Dim source As New StaticSolidElement(
                "Перестроенное тело", "SmdxElement", New ImProperties(), shell, New ImDocuments())
            source.Color = Color.Blue
            Dim entity As New DwgModel3DElement() With {.Element = source}
            Dim savedUniformArgb As Integer = Color.Red.ToArgb()
            Dim runtimeType As Type = MainType("CivilEnginStructures.BridgeAppearanceNativeRuntime")
            Dim method As MethodInfo = runtimeType.GetMethod(
                "CreateBodyColorClone", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)

            Dim restored As StaticSolidElement = DirectCast(
                method.Invoke(Nothing, New Object() {entity, New Integer() {savedUniformArgb, savedUniformArgb}}),
                StaticSolidElement)

            Assert.Multiple(
                Sub()
                    Assert.That(restored, [Is].Not.Null.And.Not.SameAs(source))
                    Assert.That(restored.Color.ToArgb(), [Is].EqualTo(savedUniformArgb),
                                "A uniform saved color is independent of the old solid traversal count.")
                    Assert.That(source.Color, [Is].EqualTo(Color.Blue),
                                "Preparing a restore clone must leave the rebuilt source untouched.")
                End Sub)
        End Sub

        <Test>
        Public Sub Static3DPhongMeshIsCapturedAsPaintableBodyGeometry()
            Dim sourceMaterial As PhongMaterial = Nothing
            Dim sourceMesh As MeshGeometry3D = Nothing
            Dim source As Static3DElement = CreatePhongStaticElement(sourceMaterial, sourceMesh)
            Dim entity As New DwgModel3DElement() With {.Element = source}

            Dim colors As Integer() = CaptureBodyColors(entity)
            Dim supportsBody As Boolean = False
            Dim values As Object = CaptureValues(entity, supportsBody)

            Assert.Multiple(
                Sub()
                    Assert.That(colors, [Is].EqualTo(New Integer() {Color.FromArgb(255, 0, 128, 255).ToArgb()}))
                    Assert.That(supportsBody, [Is].True,
                                "A pile Static3DElement with a Phong material must enable body color editing.")
                    Assert.That(ReadProperty(Of Integer?)(values, "BodyArgb"),
                                [Is].EqualTo(New Integer?(Color.FromArgb(255, 0, 128, 255).ToArgb())))
                End Sub)
        End Sub

        <Test>
        Public Sub BodyColorCloneChangesOnlyPhongDiffuseAndPreservesMeshAndSource()
            Dim sourceMaterial As PhongMaterial = Nothing
            Dim sourceMesh As MeshGeometry3D = Nothing
            Dim source As Static3DElement = CreatePhongStaticElement(sourceMaterial, sourceMesh)
            Dim entity As New DwgModel3DElement() With {.Element = source}
            Dim sourceAmbient As Vector3F = sourceMaterial.Ambient
            Dim sourceDiffuse As Vector3F = sourceMaterial.Diffuse
            Dim sourceSpecular As Vector3F = sourceMaterial.Specular
            Dim sourceTransparency As Single = sourceMaterial.Transparency
            Dim sourceSpecularLevel As Single = sourceMaterial.SpecularLevel
            Dim sourceShininess As Single = sourceMaterial.Shininess
            Dim target As Color = Color.FromArgb(255, 30, 140, 220)

            Dim clone As Static3DElement = DirectCast(
                CreateBodyColorClone(entity, New Integer() {target.ToArgb()}), Static3DElement)
            Dim cloneModel As GeometryModel3D = clone.GetModel()
            Dim cloneMaterial As PhongMaterial = DirectCast(cloneModel.Materials("phong"), PhongMaterial)
            Dim cloneMesh As MeshGeometry3D = cloneModel.Meshes("sweep2")

            Assert.Multiple(
                Sub()
                    Assert.That(clone, [Is].Not.SameAs(source))
                    Assert.That(cloneModel, [Is].Not.SameAs(source.GetModel()))
                    Assert.That(cloneMaterial, [Is].Not.SameAs(sourceMaterial))
                    AssertVector(sourceMaterial.Diffuse, sourceDiffuse)
                    AssertVector(cloneMaterial.Diffuse,
                                 New Vector3F(target.R / 255.0F, target.G / 255.0F, target.B / 255.0F))
                    AssertVector(cloneMaterial.Ambient, sourceAmbient)
                    AssertVector(cloneMaterial.Specular, sourceSpecular)
                    Assert.That(cloneMaterial.Transparency, [Is].EqualTo(sourceTransparency))
                    Assert.That(cloneMaterial.SpecularLevel, [Is].EqualTo(sourceSpecularLevel))
                    Assert.That(cloneMaterial.Shininess, [Is].EqualTo(sourceShininess))
                    Assert.That(cloneMaterial.Flags, [Is].EqualTo(sourceMaterial.Flags))
                    Assert.That(cloneModel.Meshes.Keys, [Is].EquivalentTo(New String() {"sweep2"}))
                    Assert.That(cloneModel.Materials.Keys, [Is].EquivalentTo(New String() {"phong"}))
                    Assert.That(cloneMesh.Positions.Count, [Is].EqualTo(sourceMesh.Positions.Count))
                    Assert.That(cloneMesh.TriangleIndices.Count, [Is].EqualTo(sourceMesh.TriangleIndices.Count))
                    Assert.That(cloneMesh.Groups, Has.Count.EqualTo(1))
                    Assert.That(cloneMesh.Groups(0).Material, [Is].EqualTo("phong"))
                    Assert.That(cloneMesh.LocalMatrix, [Is].EqualTo(sourceMesh.LocalMatrix))
                End Sub)
        End Sub

        <Test>
        Public Sub ConstructedPileWrapperExposesNestedStatic3DPhongBody()
            Dim material As PhongMaterial = Nothing
            Dim mesh As MeshGeometry3D = Nothing
            Dim child As Static3DElement = CreatePhongStaticElement(material, mesh)
            Dim wrapper As New ConstructedModel3dElement()
            SetConstructedChild(wrapper, child)
            Assert.That(wrapper.Element, [Is].SameAs(child),
                        "The fixture must reproduce the observed pile wrapper chain.")
            Dim entity As New DwgModel3DElement() With {.Element = wrapper}

            Dim colors As Integer() = CaptureBodyColors(entity)
            Dim supportsBody As Boolean = False
            Dim values As Object = CaptureValues(entity, supportsBody)

            Assert.Multiple(
                Sub()
                    Assert.That(colors, [Is].EqualTo(New Integer() {Color.FromArgb(255, 0, 128, 255).ToArgb()}))
                    Assert.That(supportsBody, [Is].True)
                    Assert.That(ReadProperty(Of Integer?)(values, "BodyArgb"),
                                [Is].EqualTo(New Integer?(Color.FromArgb(255, 0, 128, 255).ToArgb())))
                End Sub)
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub CancellingAppearanceDraftDoesNotMutateInputOrRequestApply()
            Dim values As Object = CreateAppearanceValues()
            Dim item As Object = CreateAppearanceItem(values)
            Dim dialog As Form = CreateAppearanceDialog(item)
            Dim counter As New EventCounter()
            AddApplyHandler(dialog, counter)

            Try
                PrivateControl(Of CheckBox)(dialog, "_layerCheck").Checked = True
                PrivateControl(Of ComboBox)(dialog, "_layerValue").Text = "Черновой слой"
                PrivateControl(Of CheckBox)(dialog, "_widthCheck").Checked = True
                PrivateControl(Of NumericUpDown)(dialog, "_widthValue").Value = 3.75D

                dialog.DialogResult = DialogResult.Cancel

                Assert.Multiple(
                    Sub()
                        Assert.That(counter.Count, [Is].Zero)
                        Assert.That(dialog.DialogResult, [Is].EqualTo(DialogResult.Cancel))
                        Assert.That(item.GetType().GetProperty("Values").GetValue(item, Nothing), [Is].SameAs(values))
                        Assert.That(ReadProperty(Of String)(values, "LayerName"), [Is].EqualTo("Исходный слой"))
                        Assert.That(ReadProperty(Of Double?)(values, "Width"), [Is].EqualTo(New Double?(0.25)))
                    End Sub)
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub MixedCadCheckboxWithoutConcreteChoiceDoesNotRequestApply()
            Dim first As Object = CreateAppearanceItem(
                "cad-1", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                True, False, False)
            Dim second As Object = CreateAppearanceItem(
                "cad-2", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByBlock.ToCompressValue(), Color.Red.ToArgb()),
                True, False, False)
            Dim dialog As Form = CreateAppearanceDialog(first, second)
            Dim capture As New EventCounter()
            AddApplyHandler(dialog, capture)

            Try
                Assert.That(PrivateControl(Of ComboBox)(dialog, "_cadValue").SelectedIndex, [Is].Zero)
                PrivateControl(Of CheckBox)(dialog, "_cadCheck").Checked = True

                InvokeApplyChanges(dialog)

                Assert.That(capture.Count, [Is].Zero,
                            "Checking a mixed CAD color without choosing a concrete value must remain a no-op.")
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub ChoosingConcreteCadColorAfterMixedSelectionRequestsApply()
            Dim first As Object = CreateAppearanceItem(
                "cad-1", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                True, False, False)
            Dim second As Object = CreateAppearanceItem(
                "cad-2", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByBlock.ToCompressValue(), Color.Red.ToArgb()),
                True, False, False)
            Dim dialog As Form = CreateAppearanceDialog(first, second)
            Dim capture As New EventCounter()
            AddApplyHandler(dialog, capture)

            Try
                PrivateControl(Of ComboBox)(dialog, "_cadValue").SelectedIndex = 1

                InvokeApplyChanges(dialog)

                Dim patch As Object = ReadProperty(Of Object)(capture.Args, "Patch")
                Assert.Multiple(
                    Sub()
                        Assert.That(capture.Count, [Is].EqualTo(1))
                        Assert.That(ReadProperty(Of Integer?)(patch, "CadColorValue"),
                                    [Is].EqualTo(New Integer?(CadColor.ByLayer.ToCompressValue())))
                    End Sub)
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub MixedBodyCheckboxWithoutConcreteChoiceDoesNotRequestApply()
            Dim first As Object = CreateAppearanceItem(
                "body-1", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                False, True, True)
            Dim second As Object = CreateAppearanceItem(
                "body-2", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Blue.ToArgb()),
                False, True, True)
            Dim dialog As Form = CreateAppearanceDialog(first, second)
            Dim capture As New EventCounter()
            AddApplyHandler(dialog, capture)

            Try
                PrivateControl(Of CheckBox)(dialog, "_bodyCheck").Checked = True

                InvokeApplyChanges(dialog)

                Assert.That(capture.Count, [Is].Zero,
                            "Checking a mixed body color without choosing a concrete value must remain a no-op.")
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub ConcreteBodyColorDraftRequestsApply()
            Dim first As Object = CreateAppearanceItem(
                "body-1", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                False, True, True)
            Dim second As Object = CreateAppearanceItem(
                "body-2", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Blue.ToArgb()),
                False, True, True)
            Dim dialog As Form = CreateAppearanceDialog(first, second)
            Dim capture As New EventCounter()
            AddApplyHandler(dialog, capture)
            Dim selectedArgb As Integer = Color.FromArgb(255, 30, 140, 220).ToArgb()

            Try
                SetPrivateField(dialog, "_bodyArgbValue", selectedArgb)
                PrivateControl(Of CheckBox)(dialog, "_bodyCheck").Checked = True

                InvokeApplyChanges(dialog)

                Dim patch As Object = ReadProperty(Of Object)(capture.Args, "Patch")
                Assert.Multiple(
                    Sub()
                        Assert.That(capture.Count, [Is].EqualTo(1))
                        Assert.That(ReadProperty(Of Integer?)(patch, "BodyArgb"),
                                    [Is].EqualTo(New Integer?(selectedArgb)))
                    End Sub)
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub Unsupported3DOnlySelectionDisablesBodyEditor()
            Dim item As Object = CreateAppearanceItem(
                "unsupported-3d", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                False, False, True)
            Dim dialog As Form = CreateAppearanceDialog(item)

            Try
                Assert.Multiple(
                    Sub()
                        Assert.That(PrivateControl(Of CheckBox)(dialog, "_bodyCheck").Enabled, [Is].False)
                        Assert.That(PrivateControl(Of Label)(dialog, "_bodyCurrent").Text,
                                    [Is].EqualTo("Не применяется"))
                    End Sub)
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub MixedSupportedAndUnsupported3DSelectionKeepsBodyEditorEnabledForSupportedItems()
            Dim supported As Object = CreateAppearanceItem(
                "supported-3d", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Red.ToArgb()),
                False, True, True)
            Dim unsupported As Object = CreateAppearanceItem(
                "unsupported-3d", "bridge-1", "Мост 1",
                CreateAppearanceValues(CadColor.ByLayer.ToCompressValue(), Color.Blue.ToArgb()),
                False, False, True)
            Dim dialog As Form = CreateAppearanceDialog(supported, unsupported)
            Dim capture As New EventCounter()
            AddApplyHandler(dialog, capture)
            Dim selectedArgb As Integer = Color.FromArgb(255, 35, 145, 225).ToArgb()

            Try
                Assert.That(PrivateControl(Of CheckBox)(dialog, "_bodyCheck").Enabled, [Is].True)
                SetPrivateField(dialog, "_bodyArgbValue", selectedArgb)
                PrivateControl(Of CheckBox)(dialog, "_bodyCheck").Checked = True

                InvokeApplyChanges(dialog)

                Dim patch As Object = ReadProperty(Of Object)(capture.Args, "Patch")
                Dim selectionIds As IEnumerable = ReadProperty(Of IEnumerable)(capture.Args, "SelectionIds")
                Assert.Multiple(
                    Sub()
                        Assert.That(capture.Count, [Is].EqualTo(1))
                        Assert.That(ReadProperty(Of Integer?)(patch, "BodyArgb"),
                                    [Is].EqualTo(New Integer?(selectedArgb)))
                        Assert.That(selectionIds.Cast(Of String)(),
                                    [Is].EquivalentTo(New String() {"supported-3d", "unsupported-3d"}),
                                    "The runtime capability filter decides which selected 3D targets receive body color.")
                    End Sub)
            Finally
                dialog.Dispose()
            End Try
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub TreeGroupsEquivalentGuidFormatsButKeepsLegacyCaseDistinct()
            Dim values As Object = CreateAppearanceValues()
            Dim guidDialog As Form = CreateAppearanceDialog(
                CreateAppearanceItem(
                    "guid-1", "  {7F5F0F47-53A7-4C89-AB9F-66F8ED463628}  ", "Мост GUID",
                    values, False, False, False),
                CreateAppearanceItem(
                    "guid-2", "7f5f0f47-53a7-4c89-ab9f-66f8ed463628", "Мост GUID",
                    values, False, False, False))
            Dim legacyDialog As Form = CreateAppearanceDialog(
                CreateAppearanceItem(
                    "legacy-1", "Legacy Bridge 17", "Старый мост",
                    values, False, False, False),
                CreateAppearanceItem(
                    "legacy-2", "legacy bridge 17", "Старый мост",
                    values, False, False, False))

            Try
                Assert.Multiple(
                    Sub()
                        Assert.That(PrivateControl(Of TreeView)(guidDialog, "_tree").Nodes.Count,
                                    [Is].EqualTo(1),
                                    "Equivalent GUID formats must produce one bridge root.")
                        Assert.That(PrivateControl(Of TreeView)(legacyDialog, "_tree").Nodes.Count,
                                    [Is].EqualTo(2),
                                    "Legacy identifiers remain case-sensitive persisted identities.")
                    End Sub)
            Finally
                guidDialog.Dispose()
                legacyDialog.Dispose()
            End Try
        End Sub

        <Test>
        Public Sub SemanticKeyCanonicalizesBridgeIdAndDistinguishesRoleAndMetadata()
            Const bridgeGuid As String = "7f5f0f47-53a7-4c89-ab9f-66f8ed463628"
            Dim original As Object = CreateStructureElement(
                "  {7F5F0F47-53A7-4C89-AB9F-66F8ED463628}  ", 1, 18, 103,
                "{""numberProlet"":2,""numberRow"":4}")
            Dim sameLogicalElement As Object = CreateStructureElement(
                bridgeGuid, 1, 18, 103,
                "{""numberRow"":4,""numberProlet"":2}")
            Dim otherBridge As Object = CreateStructureElement(
                "a789218e-af2f-4b41-a433-c13a20ddb1cf", 1, 18, 103,
                "{""numberProlet"":2,""numberRow"":4}")
            Dim otherRole As Object = CreateStructureElement(
                bridgeGuid, 1, 18, 163,
                "{""numberProlet"":2,""numberRow"":4}")
            Dim otherRow As Object = CreateStructureElement(
                bridgeGuid, 1, 18, 103,
                "{""numberProlet"":2,""numberRow"":5}")

            Dim originalKey As String = SemanticKey(original)

            Assert.Multiple(
                Sub()
                    Assert.That(originalKey, [Is].Not.Null.And.Not.Empty)
                    Assert.That(SemanticKey(sameLogicalElement), [Is].EqualTo(originalKey),
                                "Equivalent persisted GUID formats must identify one bridge element.")
                    Assert.That(SemanticKey(otherBridge), [Is].Not.EqualTo(originalKey))
                    Assert.That(SemanticKey(otherRole), [Is].Not.EqualTo(originalKey))
                    Assert.That(SemanticKey(otherRow), [Is].Not.EqualTo(originalKey))
                End Sub)
        End Sub

        <Test>
        Public Sub SemanticKeySkipsMalformedAndUnsupportedMetadata()
            Dim malformed As Object = CreateStructureElement(
                "7f5f0f47-53a7-4c89-ab9f-66f8ed463628", 1, 18, 103, "{bad-json")
            Dim missingRow As Object = CreateStructureElement(
                "7f5f0f47-53a7-4c89-ab9f-66f8ed463628", 1, 18, 103,
                "{""numberProlet"":2}")
            Dim unsupported As Object = CreateStructureElement(
                "7f5f0f47-53a7-4c89-ab9f-66f8ed463628", 0, 0, 0, "{}")

            Assert.Multiple(
                Sub()
                    Assert.That(SemanticKey(malformed), [Is].Null)
                    Assert.That(SemanticKey(missingRow), [Is].Null)
                    Assert.That(SemanticKey(unsupported), [Is].Null)
                End Sub)
        End Sub

        <Test>
        Public Sub RebuildPatchPreservesLegacyExplicitTopLinetypeAcrossAuxiliaryLayerMigration()
            Dim saved As Object = CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(), "Старый слой верха", "DASHDOT", Color.Red.ToArgb())
            Dim current As Object = CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(), "ИССО-П_Балка (скрытый верх)", "Continuous", Color.Red.ToArgb())
            Dim target As Object = CreateNativeTargetForRebuild(current, "counterTopBeam")
            Dim scopeType As Type = MainType("CivilEnginStructures.BridgeAppearanceRebuildScope")
            Dim method As MethodInfo = scopeType.GetMethod(
                "ToPatch", BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)

            Dim patch As Object = method.Invoke(Nothing, New Object() {saved, target})

            Assert.Multiple(
                Sub()
                    Assert.That(ReadProperty(Of String)(patch, "LayerName"), [Is].Null,
                                "Legacy top contours stay on the current auxiliary layer after migration.")
                    Assert.That(ReadProperty(Of String)(patch, "LinetypeName"), [Is].EqualTo("DASHDOT"),
                                "An explicit legacy top linetype must survive the layer migration.")
                End Sub)
        End Sub

        <TestCase("counterSubFermentersTop")>
        <TestCase("counterSubFermentersBottom")>
        <TestCase("counterSubFermentersUTop")>
        <TestCase("counterSubFermentersUBottom")>
        Public Sub RebuildPatchKeepsSubFermenterContoursOnCurrentAuxiliaryLayer(objectTypeName As String)
            Dim savedCadColorValue As Integer = New CadColor(Color.FromArgb(25, 75, 125)).ToCompressValue()
            Dim saved As Object = CreateAppearanceValues(
                savedCadColorValue, "Старый слой подферменника", "DASHDOT", Color.Red.ToArgb())
            Dim current As Object = CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(), "ИССО-П_Подферменник (скрытый контур)",
                "Continuous", Color.Red.ToArgb())
            Dim target As Object = CreateNativeTargetForRebuild(current, objectTypeName)
            Dim scopeType As Type = MainType("CivilEnginStructures.BridgeAppearanceRebuildScope")
            Dim method As MethodInfo = scopeType.GetMethod(
                "ToPatch", BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)

            Dim patch As Object = method.Invoke(Nothing, New Object() {saved, target})

            Assert.Multiple(
                Sub()
                    Assert.That(ReadProperty(Of String)(patch, "LayerName"), [Is].Null,
                                "Rebuilt subfermenter contours must stay on the new auxiliary layer.")
                    Assert.That(ReadProperty(Of Integer?)(patch, "CadColorValue"),
                                [Is].EqualTo(New Integer?(savedCadColorValue)),
                                "The saved contour color must survive the layer migration.")
                    Assert.That(ReadProperty(Of String)(patch, "LinetypeName"), [Is].EqualTo("DASHDOT"),
                                "The saved contour linetype must survive the layer migration.")
                End Sub)
        End Sub

        <TestCase("axisSubFermenters")>
        <TestCase("modelSubFermenters")>
        <TestCase("modelUSubFermenters")>
        Public Sub RebuildPatchRestoresSavedLayerForSubFermenterAxisAndModels(objectTypeName As String)
            Dim saved As Object = CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(), "Старый слой подферменника", "DASHDOT", Color.Red.ToArgb())
            Dim current As Object = CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(), "Новый слой подферменника", "Continuous", Color.Red.ToArgb())
            Dim target As Object = CreateNativeTargetForRebuild(current, objectTypeName)
            Dim scopeType As Type = MainType("CivilEnginStructures.BridgeAppearanceRebuildScope")
            Dim method As MethodInfo = scopeType.GetMethod(
                "ToPatch", BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)

            Dim patch As Object = method.Invoke(Nothing, New Object() {saved, target})

            Assert.That(ReadProperty(Of String)(patch, "LayerName"),
                        [Is].EqualTo("Старый слой подферменника"),
                        "Axes and models keep the normal saved-layer restoration behavior.")
        End Sub

        <Test>
        Public Sub RequireStructureLineRejectsMissingSurfaceBeforeNativeWork()
            Dim structuresLinesType As Type = MainType("CivilEnginStructures.StructuresLines")
            Dim method As MethodInfo = structuresLinesType.GetMethods(BindingFlags.Public Or BindingFlags.Static).
                Single(Function(candidate) candidate.Name = "RequireStructureLineByPolyline3d")

            Dim thrown As TargetInvocationException = Assert.Throws(Of TargetInvocationException)(
                Sub()
                    method.Invoke(Nothing, New Object() {Nothing, Nothing, 0, String.Empty, 0.0})
                End Sub)

            Assert.Multiple(
                Sub()
                    Assert.That(thrown.InnerException, [Is].TypeOf(Of ArgumentNullException)())
                    Dim argumentFailure As ArgumentNullException = DirectCast(thrown.InnerException, ArgumentNullException)
                    Assert.That(argumentFailure.ParamName, [Is].EqualTo("acSurface"))
                    Assert.That(argumentFailure.Message, Does.Contain("поверхност"))
                End Sub)
        End Sub

        Private Shared Function CreateAppearanceValues() As Object
            Return CreateAppearanceValues(
                CadColor.ByLayer.ToCompressValue(),
                Color.FromArgb(255, 20, 40, 60).ToArgb())
        End Function

        Private Shared Function CreatePhongStaticElement(ByRef material As PhongMaterial,
                                                         ByRef mesh As MeshGeometry3D) As Static3DElement
            material = New PhongMaterial With {
                .Ambient = New Vector3F(0.05F, 0.1F, 0.2F),
                .Diffuse = New Vector3F(0.0F, 0.5F, 1.0F),
                .Specular = New Vector3F(0.85F, 0.75F, 0.65F),
                .SpecularLevel = 0.7F,
                .Shininess = 0.4F,
                .Transparency = 0.25F
            }
            mesh = New MeshGeometry3D()
            mesh.Positions.Add(New Vector3F(0.0F, 0.0F, 0.0F))
            mesh.Positions.Add(New Vector3F(2.0F, 0.0F, 0.0F))
            mesh.Positions.Add(New Vector3F(0.0F, 1.0F, 0.0F))
            mesh.Normals.Add(New Vector3F(0.0F, 0.0F, 1.0F))
            mesh.Normals.Add(New Vector3F(0.0F, 0.0F, 1.0F))
            mesh.Normals.Add(New Vector3F(0.0F, 0.0F, 1.0F))
            mesh.TriangleIndices.Add(New Topomatic.Visualization.Geometry.Face(0, 1, 2))
            Dim materialGroup As New MaterialGroup With {.Material = "phong"}
            materialGroup.AddIndex(0US)
            mesh.Groups.Add(materialGroup)

            Dim model As New GeometryModel3D()
            model.Meshes.Add("sweep2", mesh)
            model.Materials.Add("phong", material)
            Return New Static3DElement("Тестовая свая", "SmdxElement",
                                       New ImProperties(), model, New ImDocuments())
        End Function

        Private Shared Function CaptureBodyColors(entity As DwgEntity) As Integer()
            Dim method As MethodInfo = MainType("CivilEnginStructures.BridgeAppearanceNativeRuntime").GetMethod(
                "CaptureBodyColors", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            Return DirectCast(method.Invoke(Nothing, New Object() {entity}), Integer())
        End Function

        Private Shared Function CreateBodyColorClone(entity As DwgEntity,
                                                     colors As Integer()) As ImElement
            Dim method As MethodInfo = MainType("CivilEnginStructures.BridgeAppearanceNativeRuntime").GetMethod(
                "CreateBodyColorClone", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            Return DirectCast(method.Invoke(Nothing, New Object() {entity, colors}), ImElement)
        End Function

        Private Shared Function CaptureValues(entity As DwgEntity,
                                              ByRef supportsBody As Boolean) As Object
            Dim method As MethodInfo = MainType("CivilEnginStructures.BridgeAppearanceNativeRuntime").GetMethod(
                "CaptureValues", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            Dim arguments As Object() = {entity, supportsBody}
            Dim result As Object = method.Invoke(Nothing, arguments)
            supportsBody = CBool(arguments(1))
            Return result
        End Function

        Private Shared Sub SetConstructedChild(wrapper As ConstructedModel3dElement,
                                               child As ImElement)
            Dim childFields As FieldInfo() = wrapper.GetType().
                GetFields(BindingFlags.NonPublic Or BindingFlags.Instance).
                Where(Function(field) GetType(ImElement).IsAssignableFrom(field.FieldType)).
                ToArray()
            Assert.That(childFields, Has.Length.EqualTo(1),
                        "ConstructedModel3dElement must have one nested ImElement backing field.")
            childFields(0).SetValue(wrapper, child)
        End Sub

        Private Shared Sub AssertVector(actual As Vector3F, expected As Vector3F)
            Const tolerance As Single = 0.000001F
            Assert.Multiple(
                Sub()
                    Assert.That(actual.X, [Is].EqualTo(expected.X).Within(tolerance))
                    Assert.That(actual.Y, [Is].EqualTo(expected.Y).Within(tolerance))
                    Assert.That(actual.Z, [Is].EqualTo(expected.Z).Within(tolerance))
                End Sub)
        End Sub

        Private Shared Function CreateAppearanceValues(cadColorValue As Integer,
                                                       bodyArgb As Integer) As Object
            Return CreateAppearanceValues(cadColorValue, "Исходный слой", "Continuous", bodyArgb)
        End Function

        Private Shared Function CreateAppearanceValues(cadColorValue As Integer,
                                                       layerName As String,
                                                       linetypeName As String,
                                                       bodyArgb As Integer) As Object
            Dim valuesType As Type = MainType("CivilEnginStructures.BridgeAppearanceValues")
            Dim constructor As ConstructorInfo = valuesType.GetConstructor(
                New Type() {
                    GetType(Integer), GetType(String), GetType(String), GetType(Double),
                    GetType(Integer), GetType(Double?), GetType(Integer?)
                })
            Return constructor.Invoke(
                New Object() {
                    cadColorValue, layerName, linetypeName, 1.0,
                    25, New Double?(0.25), New Integer?(bodyArgb)
                })
        End Function

        Private Shared Function CreateNativeTargetForRebuild(currentValues As Object,
                                                             objectTypeName As String) As Object
            Dim item As Object = CreateAppearanceItem(
                "rebuild-target", "bridge-1", "Мост 1", currentValues, False, False, False)
            Dim dataType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim data As Object = Activator.CreateInstance(dataType)
            Dim nameProperty As PropertyInfo = dataType.GetProperty("Name")
            nameProperty.SetValue(data, [Enum].Parse(nameProperty.PropertyType, objectTypeName), Nothing)

            Dim targetType As Type = MainType("CivilEnginStructures.BridgeAppearanceNativeTarget")
            Dim target As Object = Activator.CreateInstance(targetType)
            targetType.GetProperty("Info").SetValue(target, item, Nothing)
            targetType.GetProperty("Data").SetValue(target, data, Nothing)
            Return target
        End Function

        Private Shared Function CreateAppearanceItem(values As Object) As Object
            Return CreateAppearanceItem(
                "selection-1", "bridge-1", "Мост 1", values, True, True, True)
        End Function

        Private Shared Function CreateAppearanceItem(selectionId As String,
                                                     bridgeId As String,
                                                     bridgeLabel As String,
                                                     values As Object,
                                                     supportsWidth As Boolean,
                                                     supportsBody As Boolean,
                                                     isModel3D As Boolean) As Object
            Dim itemType As Type = MainType("CivilEnginStructures.BridgeAppearanceItem")
            Dim constructor As ConstructorInfo = itemType.GetConstructor(
                New Type() {
                    GetType(String), GetType(String), GetType(String), GetType(String),
                    GetType(String), GetType(String), values.GetType(), GetType(Boolean),
                    GetType(Boolean), GetType(Boolean)
                })
            Assert.That(constructor, [Is].Not.Null)
            Return constructor.Invoke(
                New Object() {
                    selectionId, bridgeId, bridgeLabel, "Пролётное строение",
                    "Балка", selectionId, values, supportsWidth, supportsBody, isModel3D
                })
        End Function

        Private Shared Function CreateAppearanceDialog(ParamArray sourceItems As Object()) As Form
            Assert.That(sourceItems, [Is].Not.Null.And.Not.Empty)
            Dim itemType As Type = sourceItems(0).GetType()
            Dim listType As Type = GetType(List(Of )).MakeGenericType(itemType)
            Dim items As IList = DirectCast(Activator.CreateInstance(listType), IList)
            For Each item As Object In sourceItems
                items.Add(item)
            Next
            Dim enumerableType As Type = GetType(IEnumerable(Of )).MakeGenericType(itemType)
            Dim dialogType As Type = MainType("CivilEnginStructures.BridgeAppearanceDialog")
            Dim constructor As ConstructorInfo = dialogType.GetConstructor(
                New Type() {
                    enumerableType,
                    GetType(IEnumerable(Of String)),
                    GetType(IEnumerable(Of String))
                })
            Assert.That(constructor, [Is].Not.Null)
            Return DirectCast(
                constructor.Invoke(
                    New Object() {
                        items,
                        New String() {"Исходный слой", "Черновой слой"},
                        New String() {"Continuous"}
                    }),
                Form)
        End Function

        Private Shared Sub AddApplyHandler(dialog As Form, counter As EventCounter)
            Dim eventInfo As EventInfo = dialog.GetType().GetEvent("ApplyRequested")
            Assert.That(eventInfo, [Is].Not.Null)
            Dim invokeMethod As MethodInfo = eventInfo.EventHandlerType.GetMethod("Invoke")
            Dim parameters As ParameterExpression() = invokeMethod.GetParameters().
                Select(Function(parameter) Expression.Parameter(parameter.ParameterType, parameter.Name)).
                ToArray()
            Dim body As MethodCallExpression = Expression.Call(
                Expression.Constant(counter),
                GetType(EventCounter).GetMethod(NameOf(EventCounter.Capture)),
                Expression.Convert(parameters(1), GetType(Object)))
            Dim handler As [Delegate] = Expression.Lambda(
                eventInfo.EventHandlerType, body, parameters).Compile()
            eventInfo.AddEventHandler(dialog, handler)
        End Sub

        Private Shared Sub InvokeApplyChanges(dialog As Form)
            Dim method As MethodInfo = dialog.GetType().GetMethod(
                "ApplyChanges", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(method, [Is].Not.Null)
            method.Invoke(dialog, New Object() {False})
        End Sub

        Private Shared Sub SetPrivateField(instance As Object,
                                           fieldName As String,
                                           value As Object)
            Dim field As FieldInfo = instance.GetType().GetField(
                fieldName, BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(field, [Is].Not.Null)
            field.SetValue(instance, value)
        End Sub

        Private Shared Function PrivateControl(Of T As Control)(dialog As Form,
                                                                 fieldName As String) As T
            Dim field As FieldInfo = dialog.GetType().GetField(
                fieldName, BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(field, [Is].Not.Null)
            Return DirectCast(field.GetValue(dialog), T)
        End Function

        Private Shared Function CreateStructureElement(bridgeId As String,
                                                       bridgeClass As Integer,
                                                       structureClass As Integer,
                                                       objectType As Integer,
                                                       keyParameter As String) As Object
            Dim dataType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim data As Object = Activator.CreateInstance(dataType)
            SetEnumProperty(data, "ClassBridgeObject", bridgeClass)
            SetEnumProperty(data, "ClassObject", structureClass)
            SetEnumProperty(data, "Name", objectType)
            dataType.GetProperty("IdStructure").SetValue(data, bridgeId, Nothing)
            dataType.GetProperty("KeyParameter").SetValue(data, keyParameter, Nothing)
            Return data
        End Function

        Private Shared Sub SetEnumProperty(instance As Object,
                                           propertyName As String,
                                           numericValue As Integer)
            Dim propertyInfo As PropertyInfo = instance.GetType().GetProperty(propertyName)
            propertyInfo.SetValue(
                instance, [Enum].ToObject(propertyInfo.PropertyType, numericValue), Nothing)
        End Sub

        Private Shared Function SemanticKey(data As Object) As String
            Dim scopeType As Type = MainType("CivilEnginStructures.BridgeAppearanceRebuildScope")
            Dim method As MethodInfo = scopeType.GetMethod(
                "SemanticKey", BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            Return TryCast(method.Invoke(Nothing, New Object() {data}), String)
        End Function

        Private Shared Function ReadProperty(Of T)(instance As Object,
                                                   propertyName As String) As T
            Return DirectCast(instance.GetType().GetProperty(propertyName).GetValue(instance, Nothing), T)
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
                        "Build CivilEnginStructures.vbproj before running bridge appearance runtime tests.")
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

        Private NotInheritable Class EventCounter
            Public Property Count As Integer
            Public Property Args As Object

            Public Sub Capture(args As Object)
                Count += 1
                Me.Args = args
            End Sub
        End Class
    End Class
End Namespace
