Imports System.IO
Imports System.Reflection
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization.Runtime

Namespace Tests
    <TestFixture>
    Public Class BridgeDrawingLayerNativeIntegrationTests
        Private Shared _mainAssembly As Assembly

        <Test, Explicit("Requires an initialized Topomatic application; New Drawing() blocks outside the host.")>
        <Category("TopomaticIntegration")>
        Public Sub SameBridgeEntitiesFromDifferentBlocksResolveToOneContainerAfterDrawingChanges()
            Dim drawing As New Drawing()
            Dim first As DwgLine = AddLine(drawing.ActiveSpace, 0.0)
            SetBridgeId(first, "bridge-across-blocks")
            Dim layer As Object = CreateLayer(drawing)

            Dim firstResolved As Object = InvokeLayer(layer, "ResolveSelectionObject", first)

            Dim secondaryBlock As DwgBlock = drawing.Blocks.Add("BRIDGE_TEST_SECONDARY")
            Dim second As DwgLine = AddLine(secondaryBlock, 10.0)
            SetBridgeId(second, "bridge-across-blocks")
            Dim secondResolved As Object = InvokeLayer(layer, "ResolveSelectionObject", second)

            Assert.Multiple(
                Sub()
                    Assert.That(firstResolved, [Is].Not.SameAs(first),
                                "A tagged native entity must resolve to its bridge container.")
                    Assert.That(secondResolved, [Is].SameAs(firstResolved),
                                "Members with the same bridge id must share one container even across blocks.")
                End Sub)
        End Sub

        <Test, Explicit("Requires an initialized Topomatic application; New Drawing() blocks outside the host.")>
        <Category("TopomaticIntegration")>
        Public Sub PlanSelectionExcludesTaggedModelButKeepsUnrelatedModelSelectable()
            Dim drawing As New Drawing()
            Dim tagged As New DwgModel3DElement()
            Dim unrelated As New DwgModel3DElement()
            drawing.ActiveSpace.Add(tagged)
            drawing.ActiveSpace.Add(unrelated)
            SetBridgeId(tagged, "bridge-model")
            Dim layer As Object = CreateLayer(drawing)

            Dim taggedSelectable As Boolean = CBool(InvokeLayer(layer, "IsPlanSelectable", tagged))
            Dim unrelatedSelectable As Boolean = CBool(InvokeLayer(layer, "IsPlanSelectable", unrelated))

            Assert.Multiple(
                Sub()
                    Assert.That(taggedSelectable, [Is].False,
                                "A bridge 3D body is represented by bridge plan graphics and must not be picked twice.")
                    Assert.That(unrelatedSelectable, [Is].True,
                                "The bridge layer must not hide an unrelated native 3D body.")
                End Sub)
        End Sub

        <Test, Explicit("Requires an initialized Topomatic application; New Drawing() blocks outside the host.")>
        <Category("TopomaticIntegration")>
        Public Sub SynchronizeGroupsTaggedEntityStoredInANonActiveBlock()
            Dim drawing As New Drawing()
            Dim secondaryBlock As DwgBlock = drawing.Blocks.Add("BRIDGE_TEST_GROUP_BLOCK")
            Dim entity As DwgLine = AddLine(secondaryBlock, 20.0)
            SetBridgeId(entity, "bridge-in-custom-block")

            Dim changed As Integer = CInt(InvokeManager("Synchronize", drawing))

            Assert.Multiple(
                Sub()
                    Assert.That(changed, [Is].EqualTo(1))
                    Assert.That(entity.Group, [Is].Not.Null)
                    If entity.Group IsNot Nothing Then
                        Assert.That(entity.Group.Name,
                                    [Is].EqualTo(CStr(InvokeManager(
                                        "GetGroupName", "bridge-in-custom-block"))))
                    End If
                End Sub)
        End Sub

        Private Shared Function AddLine(block As DwgBlock, offset As Double) As DwgLine
            Return block.AddLine(
                New Vector3D(offset, 0.0, 0.0),
                New Vector3D(offset + 1.0, 0.0, 0.0))
        End Function

        Private Shared Function CreateLayer(drawing As Drawing) As Object
            Dim layerType As Type = MainType("CivilEnginStructures.BridgeDrawingLayer")
            Dim constructor As ConstructorInfo = layerType.GetConstructor(
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(String)},
                Nothing)
            Assert.That(constructor, [Is].Not.Null)
            Dim layer As Object = constructor.Invoke(New Object() {"Native bridge index integration"})
            layerType.GetProperty("Drawing", BindingFlags.Public Or BindingFlags.Instance).
                SetValue(layer, drawing, Nothing)
            Return layer
        End Function

        Private Shared Function InvokeLayer(layer As Object,
                                            methodName As String,
                                            argument As Object) As Object
            Dim method As MethodInfo = layer.GetType().GetMethod(
                methodName,
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(Object)},
                Nothing)
            Assert.That(method, [Is].Not.Null, methodName & " must remain available to the bridge selection layer.")
            Return method.Invoke(layer, New Object() {argument})
        End Function

        Private Shared Function InvokeManager(methodName As String,
                                              ParamArray arguments As Object()) As Object
            Dim method As MethodInfo = MainType("CivilEnginStructures.BridgeDrawingGroupManager").GetMethod(
                methodName, BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null, methodName & " must remain a public shared manager operation.")
            Return method.Invoke(Nothing, arguments)
        End Function

        Private Shared Sub SetBridgeId(entity As DwgEntity, idStructure As String)
            Dim structureType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim data As Object = Activator.CreateInstance(structureType)
            structureType.GetProperty("IdStructure").SetValue(data, idStructure, Nothing)
            Dim tableType As Type = structureType.GetNestedType("tableXRecords")
            Dim projectStructures As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim setXRecords As MethodInfo = MainType("CivilEnginStructures.FuncXRecords").GetMethod(
                "setXRecords", BindingFlags.Public Or BindingFlags.Static)

            Dim written As Boolean = CBool(setXRecords.Invoke(
                Nothing, New Object() {entity, projectStructures, data, String.Empty}))
            Assert.That(written, [Is].True, "The native entity must accept bridge metadata.")
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
                        "Build CivilEnginStructures.vbproj before running native bridge integration tests.")
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
