Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports NUnit.Framework
Imports Topomatic.Dwg.Entities

Namespace Tests
    <TestFixture>
    Public Class BridgeTaggedEntityIndexTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub BuildGroupsTaggedMembersEvenWhenBridgeAxisIsMissing()
            Dim first As New DwgLine()
            Dim second As New DwgLine()
            SetMetadata(first, "  legacy-78  ", "axisBeam", String.Empty)
            SetMetadata(second, "legacy-78", "axisBeam", String.Empty)

            Dim index As IDictionary = BuildIndex(first, second)
            Dim record As Object = OnlyRecord(index)

            Assert.Multiple(
                Sub()
                    Assert.That(index.Count, [Is].EqualTo(1))
                    Assert.That(RecordString(record, "IdStructure"), [Is].EqualTo("legacy-78"))
                    Assert.That(RecordValue(record, "Bridge"), [Is].Null)
                    Assert.That(RecordEntities(record), [Is].EquivalentTo(New DwgEntity() {first, second}))
                End Sub)
        End Sub

        <Test>
        Public Sub BuildEnrichesExistingMembersWhenValidBridgeAxisIsPresent()
            Dim beam As New DwgLine()
            Dim contour As New DwgLine()
            Dim axis As New DwgLine()
            SetMetadata(beam, "bridge-78", "axisBeam", String.Empty)
            SetMetadata(contour, "bridge-78", "counterTopBeam", String.Empty)
            SetMetadata(axis, "bridge-78", "axisBridge", ValidBridgeJson("Мост 78"))

            Dim index As IDictionary = BuildIndex(beam, contour, axis)
            Dim record As Object = OnlyRecord(index)
            Dim bridge As Object = RecordValue(record, "Bridge")

            Assert.Multiple(
                Sub()
                    Assert.That(index.Count, [Is].EqualTo(1))
                    Assert.That(RecordEntities(record), [Is].EquivalentTo(New DwgEntity() {beam, contour, axis}))
                    Assert.That(bridge, [Is].Not.Null)
                    If bridge IsNot Nothing Then
                        Assert.That(CStr(bridge.GetType().GetProperty("NameBridge").GetValue(bridge, Nothing)),
                                    [Is].EqualTo("Мост 78"))
                    End If
                End Sub)
        End Sub

        <Test>
        Public Sub BuildMergesEquivalentGuidSpellingsUnderOneCanonicalKey()
            Dim guidValue As Guid = New Guid("{7F5F0F47-53A7-4C89-AB9F-66F8ED463628}")
            Dim firstId As String = "  " & guidValue.ToString("B").ToUpperInvariant() & "  "
            Dim secondId As String = guidValue.ToString("D").ToLowerInvariant()
            Dim first As New DwgLine()
            Dim second As New DwgLine()
            SetMetadata(first, firstId, "axisBeam", String.Empty)
            SetMetadata(second, secondId, "axisBeam", String.Empty)

            Dim index As IDictionary = BuildIndex(first, second)
            Dim record As Object = OnlyRecord(index)
            Dim expectedCanonicalKey As String = GetCanonicalKey(firstId)

            Assert.Multiple(
                Sub()
                    Assert.That(index.Count, [Is].EqualTo(1))
                    Assert.That(index.Contains(expectedCanonicalKey), [Is].True)
                    Assert.That(RecordString(record, "IdStructure"),
                                [Is].EqualTo(guidValue.ToString("B").ToUpperInvariant()))
                    Assert.That(RecordEntities(record).Count, [Is].EqualTo(2))
                End Sub)
        End Sub

        <Test>
        Public Sub BuildSkipsUntaggedAndBlankBridgeIdentifiers()
            Dim untagged As New DwgLine()
            Dim blank As New DwgLine()
            SetMetadata(blank, "   ", "axisBeam", String.Empty)

            Dim index As IDictionary = BuildIndex(untagged, blank)

            Assert.That(index, [Is].Empty)
        End Sub

        <Test>
        Public Sub BuildKeepsAllMembersWhenBridgeAxisJsonIsMalformed()
            Dim beam As New DwgLine()
            Dim axis As New DwgLine()
            SetMetadata(beam, "bridge-broken", "axisBeam", String.Empty)
            SetMetadata(axis, "bridge-broken", "axisBridge", "{not-json")

            Dim index As IDictionary = BuildIndex(beam, axis)
            Dim record As Object = OnlyRecord(index)

            Assert.Multiple(
                Sub()
                    Assert.That(index.Count, [Is].EqualTo(1))
                    Assert.That(RecordEntities(record), [Is].EquivalentTo(New DwgEntity() {beam, axis}))
                    Assert.That(RecordValue(record, "Bridge"), [Is].Null)
                End Sub)
        End Sub

        Private Shared Function BuildIndex(ParamArray entities As DwgEntity()) As IDictionary
            Dim indexType As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgeTaggedEntityIndex",
                throwOnError:=False)
            Assert.That(indexType, [Is].Not.Null,
                        "Container selection and appearance need one shared tagged-entity index.")
            If indexType Is Nothing Then Return New Hashtable()
            Dim method As MethodInfo = indexType.GetMethod(
                "Build",
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            If method Is Nothing Then Return New Hashtable()
            Return DirectCast(method.Invoke(Nothing, New Object() {entities.AsEnumerable()}), IDictionary)
        End Function

        Private Shared Function OnlyRecord(index As IDictionary) As Object
            Assert.That(index.Count, [Is].EqualTo(1))
            If index.Count <> 1 Then Return Nothing
            Return index.Values.Cast(Of Object)().Single()
        End Function

        Private Shared Function RecordValue(record As Object, propertyName As String) As Object
            If record Is Nothing Then Return Nothing
            Dim propertyInfo As PropertyInfo = record.GetType().GetProperty(
                propertyName,
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(propertyInfo, [Is].Not.Null, propertyName)
            If propertyInfo Is Nothing Then Return Nothing
            Return propertyInfo.GetValue(record, Nothing)
        End Function

        Private Shared Function RecordString(record As Object, propertyName As String) As String
            Return CStr(RecordValue(record, propertyName))
        End Function

        Private Shared Function RecordEntities(record As Object) As List(Of DwgEntity)
            Dim value As IEnumerable = TryCast(RecordValue(record, "Entities"), IEnumerable)
            Assert.That(value, [Is].Not.Null)
            If value Is Nothing Then Return New List(Of DwgEntity)()
            Return value.Cast(Of DwgEntity)().ToList()
        End Function

        Private Shared Sub SetMetadata(entity As DwgEntity,
                                       idStructure As String,
                                       objectName As String,
                                       keyParameter As String)
            Dim structureType As Type = MainAssembly().GetType(
                "CivilEnginStructures.StructureElement", throwOnError:=True)
            Dim data As Object = Activator.CreateInstance(structureType)
            structureType.GetProperty("IdStructure").SetValue(data, idStructure, Nothing)
            structureType.GetProperty("KeyParameter").SetValue(data, keyParameter, Nothing)
            Dim objectType As Type = structureType.GetNestedType("typeObject")
            structureType.GetProperty("Name").SetValue(
                data, [Enum].Parse(objectType, objectName), Nothing)

            Dim tableType As Type = structureType.GetNestedType("tableXRecords")
            Dim table As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim xrecordsType As Type = MainAssembly().GetType(
                "CivilEnginStructures.FuncXRecords", throwOnError:=True)
            Dim method As MethodInfo = xrecordsType.GetMethod("setXRecords")
            Dim written As Boolean = CBool(method.Invoke(
                Nothing, New Object() {entity, table, data, String.Empty}))
            Assert.That(written, [Is].True)
        End Sub

        Private Shared Function ValidBridgeJson(name As String) As String
            Dim bridgeType As Type = MainAssembly().GetType(
                "CivilEnginStructures.Bridges", throwOnError:=True)
            Dim bridge As Object = Activator.CreateInstance(bridgeType)
            bridgeType.GetProperty("NameBridge").SetValue(bridge, name, Nothing)
            Return JsonConvert.SerializeObject(bridge)
        End Function

        Private Shared Function GetCanonicalKey(idStructure As String) As String
            Dim managerType As Type = MainAssembly().GetType(
                "CivilEnginStructures.BridgeDrawingGroupManager", throwOnError:=True)
            Return CStr(managerType.GetMethod("GetGroupName").Invoke(
                Nothing, New Object() {idStructure}))
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
                        "Build CivilEnginStructures before running bridge index tests.")
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
