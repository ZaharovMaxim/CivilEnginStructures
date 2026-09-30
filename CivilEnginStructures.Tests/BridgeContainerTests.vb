Imports System.IO
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BridgeContainerTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub NewStateStartsWithNoOpenBridge()
            Dim state As Object = CreateState()

            Assert.Multiple(
                Sub()
                    Assert.That(OpenBridgeId(state), [Is].Null.Or.Empty)
                    Assert.That(IsBridgeOpen(state, "bridge-a"), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub EnterBridgeOpensEquivalentGuidFormatsAsOneBridge()
            Dim state As Object = CreateState()

            EnterBridge(state, "  {7F5F0F47-53A7-4C89-AB9F-66F8ED463628}  ")

            Assert.Multiple(
                Sub()
                    Assert.That(OpenBridgeId(state), [Is].Not.Null.And.Not.Empty)
                    Assert.That(IsBridgeOpen(state, "7f5f0f47-53a7-4c89-ab9f-66f8ed463628"), [Is].True)
                    Assert.That(IsBridgeOpen(state, "a789218e-af2f-4b41-a433-c13a20ddb1cf"), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub EmptyBridgeIdDoesNotOpenOrReplaceABridge()
            Dim state As Object = CreateState()

            EnterBridge(state, Nothing)
            Assert.That(OpenBridgeId(state), [Is].Null.Or.Empty)

            EnterBridge(state, "bridge-a")
            EnterBridge(state, "   ")

            Assert.Multiple(
                Sub()
                    Assert.That(IsBridgeOpen(state, "bridge-a"), [Is].True)
                    Assert.That(IsBridgeOpen(state, String.Empty), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub ExitBridgeClosesTheCurrentBridge()
            Dim state As Object = CreateState()
            EnterBridge(state, "bridge-a")

            ExitBridge(state)

            Assert.Multiple(
                Sub()
                    Assert.That(OpenBridgeId(state), [Is].Null.Or.Empty)
                    Assert.That(IsBridgeOpen(state, "bridge-a"), [Is].False)
                End Sub)
        End Sub

        <Test>
        Public Sub UntaggedAndMissingSelectionItemsPassThroughWithoutAContainer()
            Dim state As Object = CreateState()
            Dim item As New Object()

            Assert.Multiple(
                Sub()
                    Assert.That(ResolveSelectionObject(state, item, Nothing), [Is].SameAs(item))
                    Assert.That(ResolveSelectionObject(state, item, "   "), [Is].SameAs(item))
                    Assert.That(ResolveSelectionObject(state, Nothing, "bridge-a"), [Is].Null)
                End Sub)
        End Sub

        <Test>
        Public Sub ClosedBridgeLeavesResolveToOneCachedPresentationObject()
            Dim state As Object = CreateState()
            Dim firstLeaf As New Object()
            Dim secondLeaf As New Object()

            Dim firstResult As Object = ResolveSelectionObject(state, firstLeaf, "bridge-a")
            Dim secondResult As Object = ResolveSelectionObject(state, secondLeaf, "  bridge-a  ")

            Assert.Multiple(
                Sub()
                    Assert.That(firstResult, [Is].Not.SameAs(firstLeaf))
                    Assert.That(secondResult, [Is].SameAs(firstResult),
                                "Every leaf of a closed bridge must present the same container object.")
                    Assert.That(firstResult.GetType(), [Is].EqualTo(PresentationObjectType()))
                End Sub)
        End Sub

        <Test>
        Public Sub DifferentClosedBridgesUseDifferentCachedPresentationObjects()
            Dim state As Object = CreateState()

            Dim firstBridge As Object = ResolveSelectionObject(state, New Object(), "bridge-a")
            Dim secondBridge As Object = ResolveSelectionObject(state, New Object(), "bridge-b")
            Dim firstBridgeAgain As Object = ResolveSelectionObject(state, New Object(), "bridge-a")

            Assert.Multiple(
                Sub()
                    Assert.That(secondBridge, [Is].Not.SameAs(firstBridge))
                    Assert.That(firstBridgeAgain, [Is].SameAs(firstBridge))
                End Sub)
        End Sub

        <Test>
        Public Sub OpenBridgeReturnsItsRawLeavesWhileOtherBridgesStayContainers()
            Dim state As Object = CreateState()
            Dim openLeaf As New Object()
            Dim closedLeaf As New Object()
            EnterBridge(state, "bridge-a")

            Dim openResult As Object = ResolveSelectionObject(state, openLeaf, "bridge-a")
            Dim closedResult As Object = ResolveSelectionObject(state, closedLeaf, "bridge-b")

            Assert.Multiple(
                Sub()
                    Assert.That(openResult, [Is].SameAs(openLeaf))
                    Assert.That(closedResult, [Is].Not.SameAs(closedLeaf))
                    Assert.That(closedResult.GetType(), [Is].EqualTo(PresentationObjectType()))
                End Sub)
        End Sub

        <Test>
        Public Sub EnteringAnotherBridgeSwitchesWhichLeavesAreExposed()
            Dim state As Object = CreateState()
            Dim firstLeaf As New Object()
            Dim secondLeaf As New Object()
            EnterBridge(state, "bridge-a")

            EnterBridge(state, "bridge-b")

            Assert.Multiple(
                Sub()
                    Assert.That(IsBridgeOpen(state, "bridge-a"), [Is].False)
                    Assert.That(IsBridgeOpen(state, "bridge-b"), [Is].True)
                    Assert.That(ResolveSelectionObject(state, firstLeaf, "bridge-a"), [Is].Not.SameAs(firstLeaf))
                    Assert.That(ResolveSelectionObject(state, secondLeaf, "bridge-b"), [Is].SameAs(secondLeaf))
                End Sub)
        End Sub

        <Test>
        Public Sub LegacyIdsIgnoreOuterWhitespaceButPreserveCase()
            Dim state As Object = CreateState()
            EnterBridge(state, "  Legacy Bridge 17  ")

            Assert.Multiple(
                Sub()
                    Assert.That(IsBridgeOpen(state, "Legacy Bridge 17"), [Is].True)
                    Assert.That(IsBridgeOpen(state, "legacy bridge 17"), [Is].False,
                                "Legacy identifiers are exact persisted values; case must not merge bridges.")
                End Sub)
        End Sub

        <TestCase(True, True, False, TestName:="Plan_hides_tagged_bridge_3D_models")>
        <TestCase(True, False, True, TestName:="Plan_keeps_tagged_bridge_non_3D_entities")>
        <TestCase(False, True, True, TestName:="Plan_keeps_unrelated_3D_models")>
        <TestCase(False, False, True, TestName:="Plan_keeps_unrelated_non_3D_entities")>
        Public Sub PlanVisibilityHidesOnlyBridge3DModels(isBridgeEntity As Boolean,
                                                        isModel3d As Boolean,
                                                        expected As Boolean)
            Assert.That(ShouldDrawInPlan(isBridgeEntity, isModel3d), [Is].EqualTo(expected))
        End Sub

        Private Shared Function CreateState() As Object
            Return Activator.CreateInstance(StateType())
        End Function

        Private Shared Function OpenBridgeId(state As Object) As String
            Dim propertyInfo As PropertyInfo = state.GetType().GetProperty(
                "OpenBridgeId", BindingFlags.Public Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null,
                        "BridgeContainerState.OpenBridgeId must be Public ReadOnly.")
            Assert.That(propertyInfo.CanWrite, [Is].False,
                        "BridgeContainerState owns transitions through EnterBridge and ExitBridge.")
            Return TryCast(propertyInfo.GetValue(state, Nothing), String)
        End Function

        Private Shared Sub EnterBridge(state As Object, idStructure As String)
            InvokeState(state, "EnterBridge", New Type() {GetType(String)}, idStructure)
        End Sub

        Private Shared Sub ExitBridge(state As Object)
            InvokeState(state, "ExitBridge", Type.EmptyTypes)
        End Sub

        Private Shared Function IsBridgeOpen(state As Object, idStructure As String) As Boolean
            Return CBool(InvokeState(state, "IsBridgeOpen", New Type() {GetType(String)}, idStructure))
        End Function

        Private Shared Function ResolveSelectionObject(state As Object,
                                                       item As Object,
                                                       idStructure As String) As Object
            Return InvokeState(state,
                               "ResolveSelectionObject",
                               New Type() {GetType(Object), GetType(String)},
                               item,
                               idStructure)
        End Function

        Private Shared Function InvokeState(state As Object,
                                            methodName As String,
                                            parameterTypes As Type(),
                                            ParamArray arguments As Object()) As Object
            Dim method As MethodInfo = state.GetType().GetMethod(
                methodName,
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                parameterTypes,
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgeContainerState." & methodName & " must expose the pure container-state behavior.")
            Return method.Invoke(state, arguments)
        End Function

        Private Shared Function ShouldDrawInPlan(isBridgeEntity As Boolean,
                                                 isModel3d As Boolean) As Boolean
            Dim policyType As Type = MainType("CivilEnginStructures.BridgePlanVisibilityPolicy")
            Dim method As MethodInfo = policyType.GetMethod(
                "ShouldDrawInPlan",
                BindingFlags.Public Or BindingFlags.Static,
                Nothing,
                New Type() {GetType(Boolean), GetType(Boolean)},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "BridgePlanVisibilityPolicy.ShouldDrawInPlan must be a pure shared policy.")
            Return CBool(method.Invoke(Nothing, New Object() {isBridgeEntity, isModel3d}))
        End Function

        Private Shared Function StateType() As Type
            Return MainType("CivilEnginStructures.BridgeContainerState")
        End Function

        Private Shared Function PresentationObjectType() As Type
            Return MainType("CivilEnginStructures.BridgePresentationObject")
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
                        "Build CivilEnginStructures.vbproj before running bridge container tests.")
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
