Imports System.IO
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Stg

Namespace Tests
    <TestFixture>
    Public Class BridgeDrawingGroupTests
        Private Const GroupPrefix As String = "INFRA_BRIDGE_"
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub ManagerExposesTheLifecycleAndSynchronizationContract()
            Dim manager As Type = ManagerType()

            Assert.Multiple(
                Sub()
                    AssertSharedMethod(manager, "GetGroupName", GetType(String), GetType(String))
                    AssertSharedMethod(manager, "Attach", GetType(Void), GetType(Drawing))
                    AssertSharedMethod(manager, "Detach", GetType(Void), GetType(Drawing))
                    AssertSharedMethod(manager, "Synchronize", GetType(Integer), GetType(Drawing))
                    AssertSharedMethod(manager, "TryGroupEntity", Nothing, GetType(DwgEntity))
                End Sub)
        End Sub

        <TestCase("{7F5F0F47-53A7-4C89-AB9F-66F8ED463628}")>
        <TestCase("7f5f0f47-53a7-4c89-ab9f-66f8ed463628")>
        Public Sub GuidBridgeIdUsesOneCanonicalGroupName(idStructure As String)
            Dim actual As String = GetGroupName(idStructure)

            Assert.That(actual,
                        [Is].EqualTo(GroupPrefix & "7f5f0f4753a74c89ab9f66f8ed463628"))
        End Sub

        <TestCase("legacy bridge 17")>
        <TestCase("  legacy bridge 17  ")>
        <TestCase("Мост № 17")>
        Public Sub LegacyBridgeIdUsesAStableCollisionSafeGroupName(idStructure As String)
            Dim trimmedId As String = idStructure.Trim()
            Dim expected As String = GroupPrefix & "legacy_" & Sha256Hex(trimmedId)

            Assert.That(GetGroupName(idStructure), [Is].EqualTo(expected))
        End Sub

        <Test>
        Public Sub DifferentLegacyBridgeIdsDoNotShareAGroupName()
            Assert.That(GetGroupName("Bridge-A"),
                        [Is].Not.EqualTo(GetGroupName("bridge-a")),
                        "Legacy identifiers are exact persisted values; changing case must not merge bridges.")
        End Sub

        <TestCase(Nothing)>
        <TestCase("")>
        <TestCase("   ")>
        Public Sub MissingBridgeIdDoesNotCreateAGroupName(idStructure As String)
            Assert.That(GetGroupName(idStructure), [Is].EqualTo(String.Empty))
        End Sub

        <Test>
        Public Sub OwnedGroupNameAcceptsOnlyNamesGeneratedByTheManager()
            Assert.Multiple(
                Sub()
                    Assert.That(IsOwnedGroupName(GetGroupName(
                                    "7f5f0f47-53a7-4c89-ab9f-66f8ed463628")), [Is].True)
                    Assert.That(IsOwnedGroupName(GetGroupName("legacy bridge 17")), [Is].True)
                End Sub)
        End Sub

        <TestCase("INFRA_BRIDGE_CUSTOM")>
        <TestCase("INFRA_BRIDGE_7f5f0f4753a74c89ab9f66f8ed46362")>
        <TestCase("INFRA_BRIDGE_7f5f0f4753a74c89ab9f66f8ed4636280")>
        <TestCase("INFRA_BRIDGE_legacy_0123456789abcdef")>
        <TestCase("INFRA_BRIDGE_legacy_0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef_tail")>
        <TestCase("")>
        <TestCase(Nothing)>
        Public Sub OwnedGroupNameRejectsForeignOrMalformedNames(groupName As String)
            Assert.That(IsOwnedGroupName(groupName), [Is].False,
                        "A similar prefix does not make a user group owned by this plugin.")
        End Sub

        <Test>
        Public Sub InteractiveEntityPicksConsumeTheExactObjectReturnedByTopomatic()
            Dim root As String = RepositoryRoot()
            Dim sourceFiles As String() = {
                Path.Combine(root, "RopExample1Module.vb"),
                Path.Combine(root, "UserFormWPF", "PanelProjectBridge.vb")
            }
            Dim activeCalls As Integer = 0
            Dim unsafeCalls As New List(Of String)()

            For Each sourceFile As String In sourceFiles
                Dim lineNumber As Integer = 0
                For Each sourceLine As String In File.ReadLines(sourceFile)
                    lineNumber += 1
                    Dim trimmed As String = sourceLine.TrimStart()
                    If trimmed.StartsWith("'", StringComparison.Ordinal) Then Continue For

                    Dim callIndex As Integer = sourceLine.IndexOf(
                        ".SelectOneObjectAtScreen", StringComparison.Ordinal)
                    If callIndex < 0 Then Continue For

                    activeCalls += 1
                    If Not sourceLine.Substring(0, callIndex).Contains("=") Then
                        unsafeCalls.Add(Path.GetFileName(sourceFile) & ":" & lineNumber)
                    End If
                Next
            Next

            Assert.Multiple(
                Sub()
                    Assert.That(activeCalls, [Is].EqualTo(16),
                                "Keep every existing interactive element-pick workflow covered.")
                    Assert.That(unsafeCalls, [Is].Empty,
                                "A grouped SelectionSet expands to the whole bridge. The command must use " &
                                "the object returned by SelectOneObjectAtScreen for leaf editing: " &
                                String.Join(", ", unsafeCalls))
                End Sub)
        End Sub

        <Test, Explicit("Requires the initialized Topomatic application environment; New Drawing() blocks without it.")>
        <Category("TopomaticIntegration")>
        Public Sub SynchronizeSeparatesTwoBridgesAndPreservesUnmarkedEntitiesAndGroups()
            Dim drawing As New Drawing()
            Dim unrelated As DwgGroup = drawing.Groups.Add("USER_GROUP")
            Dim unmarked As DwgLine = AddLine(drawing, 0.0)
            unmarked.Group = unrelated
            Dim similarlyNamed As DwgGroup = drawing.Groups.Add("INFRA_BRIDGE_CUSTOM")
            Dim foreignPrefixedEntity As DwgLine = AddLine(drawing, 0.5)
            foreignPrefixedEntity.Group = similarlyNamed
            Dim first As DwgLine = AddBridgeLine(drawing, 1.0, "bridge-one")
            Dim second As DwgLine = AddBridgeLine(
                drawing, 2.0, "7f5f0f47-53a7-4c89-ab9f-66f8ed463628")

            Dim changed As Integer = Synchronize(drawing)

            Assert.Multiple(
                Sub()
                    Assert.That(changed, [Is].EqualTo(2))
                    Assert.That(first.Group, [Is].Not.Null)
                    Assert.That(second.Group, [Is].Not.Null)
                    Assert.That(first.Group, [Is].Not.SameAs(second.Group))
                    Assert.That(unmarked.Group, [Is].SameAs(unrelated))
                    Assert.That(foreignPrefixedEntity.Group, [Is].SameAs(similarlyNamed))
                    Assert.That(drawing.Groups.IsExists("USER_GROUP"), [Is].True)
                    Assert.That(drawing.Groups.IsExists("INFRA_BRIDGE_CUSTOM"), [Is].True)
                    Assert.That(Synchronize(drawing), [Is].EqualTo(0),
                                "Repeated synchronization must be idempotent.")
                End Sub)
        End Sub

        <Test, Explicit("Requires the initialized Topomatic application environment; New Drawing() blocks without it.")>
        <Category("TopomaticIntegration")>
        Public Sub AttachGroupsEntitiesWhenBridgeMetadataIsWrittenAndDetachStopsTracking()
            Dim drawing As New Drawing()
            InvokeManager("Attach", drawing)
            Dim tracked As DwgLine = AddBridgeLine(drawing, 1.0, "tracked-bridge")

            Assert.That(tracked.Group, [Is].Not.Null,
                        "Writing IdStructure after adding an entity must trigger grouping.")

            InvokeManager("Detach", drawing)
            Dim detached As DwgLine = AddBridgeLine(drawing, 2.0, "detached-bridge")

            Assert.That(detached.Group, [Is].Null)
        End Sub

        <Test, Explicit("Requires the initialized Topomatic application environment; New Drawing() blocks without it.")>
        <Category("TopomaticIntegration")>
        Public Sub NativeStgRoundTripPreservesBridgeGroupMembership()
            Dim original As New Drawing()
            Dim originalLine As DwgLine = AddBridgeLine(original, 1.0, "round-trip-bridge")
            Assert.That(Synchronize(original), [Is].EqualTo(1))
            Dim expectedGroupName As String = originalLine.Group.Name
            Dim document As New StgDocument()
            original.SaveToStg(document.Body)

            Using stream As New MemoryStream()
                document.SaveToStreamAsBinary(stream)
                stream.Position = 0
                Dim restoredDocument As New StgDocument()
                restoredDocument.LoadFromStreamAsBinary(stream)
                Dim restored As New Drawing()
                restored.LoadFromStg(restoredDocument.Body)

                Assert.Multiple(
                    Sub()
                        Assert.That(restored.Groups.IsExists(expectedGroupName), [Is].True)
                        Assert.That(restored.Groups(expectedGroupName).EntitysCount, [Is].EqualTo(1))
                    End Sub)
            End Using
        End Sub

        Private Shared Sub AssertSharedMethod(
            manager As Type,
            methodName As String,
            returnType As Type,
            ParamArray parameterTypes As Type())

            Dim method As MethodInfo = manager.GetMethod(
                methodName,
                BindingFlags.Public Or BindingFlags.Static,
                Nothing,
                parameterTypes,
                Nothing)
            Assert.That(method, [Is].Not.Null, methodName & " must be Public Shared.")
            If method IsNot Nothing AndAlso returnType IsNot Nothing Then
                Assert.That(method.ReturnType, [Is].EqualTo(returnType), methodName & " return type")
            End If
        End Sub

        Private Shared Function GetGroupName(idStructure As String) As String
            Return CStr(InvokeManager("GetGroupName", idStructure))
        End Function

        Private Shared Function Synchronize(drawing As Drawing) As Integer
            Return CInt(InvokeManager("Synchronize", drawing))
        End Function

        Private Shared Function IsOwnedGroupName(groupName As String) As Boolean
            Dim method As MethodInfo = ManagerType().GetMethod(
                "IsOwnedGroupName",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                Nothing,
                New Type() {GetType(String)},
                Nothing)
            Assert.That(method, [Is].Not.Null,
                        "Ownership cleanup needs one exact name validator shared by every code path.")
            If method Is Nothing Then Return False
            Return CBool(method.Invoke(Nothing, New Object() {groupName}))
        End Function

        Private Shared Function InvokeManager(methodName As String, ParamArray arguments As Object()) As Object
            Dim method As MethodInfo = ManagerType().GetMethod(
                methodName, BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null, methodName & " must be Public Shared.")
            Return method.Invoke(Nothing, arguments)
        End Function

        Private Shared Function ManagerType() As Type
            Dim manager As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgeDrawingGroupManager", throwOnError:=False)
            Assert.That(manager, [Is].Not.Null,
                        "The bridge model needs a drawing group manager for whole-bridge selection.")
            Return manager
        End Function

        Private Shared Function AddLine(drawing As Drawing, offset As Double) As DwgLine
            Return drawing.ActiveSpace.AddLine(
                New Vector3D(offset, 0.0, 0.0),
                New Vector3D(offset + 1.0, 0.0, 0.0))
        End Function

        Private Shared Function AddBridgeLine(
            drawing As Drawing,
            offset As Double,
            idStructure As String) As DwgLine

            Dim line As DwgLine = AddLine(drawing, offset)
            SetBridgeId(line, idStructure)
            Return line
        End Function

        Private Shared Sub SetBridgeId(entity As DwgEntity, idStructure As String)
            Dim structureType As Type = MainAssembly().GetType(
                "CivilEnginStructures.StructureElement", throwOnError:=True)
            Dim data As Object = Activator.CreateInstance(structureType)
            structureType.GetProperty("IdStructure").SetValue(data, idStructure)
            Dim tableType As Type = structureType.GetNestedType("tableXRecords")
            Dim projectStructures As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim xrecordsType As Type = MainAssembly().GetType(
                "CivilEnginStructures.FuncXRecords", throwOnError:=True)
            Dim setXRecords As MethodInfo = xrecordsType.GetMethod("setXRecords")

            Dim written As Boolean = CBool(setXRecords.Invoke(
                Nothing, New Object() {entity, projectStructures, data, String.Empty}))
            Assert.That(written, [Is].True, "The native entity must accept bridge metadata.")
        End Sub

        Private Shared Function Sha256Hex(value As String) As String
            Using algorithm As SHA256 = SHA256.Create()
                Dim bytes As Byte() = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value))
                Dim result As New StringBuilder(bytes.Length * 2)
                For Each item As Byte In bytes
                    result.Append(item.ToString("x2"))
                Next
                Return result.ToString()
            End Using
        End Function

        Private Shared Function MainAssembly() As Assembly
            If _mainAssembly IsNot Nothing Then Return _mainAssembly

            Dim root As String = RepositoryRoot()
            Dim candidates As String() = {
                Path.Combine(root, "bin", "Debug", "CivilEnginStructures.dll"),
                Path.Combine(root, "bin", "Release", "CivilEnginStructures.dll")
            }
            Dim assemblyPath As String = candidates.FirstOrDefault(Function(path) File.Exists(path))
            Assert.That(assemblyPath, [Is].Not.Null,
                        "Build CivilEnginStructures before running reflection contract tests.")
            _mainAssembly = Assembly.LoadFrom(assemblyPath)
            Return _mainAssembly
        End Function

        Private Shared Function RepositoryRoot() As String
            Return Path.GetFullPath(Path.Combine(TestContext.CurrentContext.TestDirectory, "..", "..", "..", ".."))
        End Function
    End Class
End Namespace
