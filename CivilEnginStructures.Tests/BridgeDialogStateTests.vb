Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BridgeDialogStateTests
        Private Const NewStructureCaption As String = "Новое искусственное сооружение"

        <Test>
        Public Sub BridgeStructureChoiceExposesStableIdentityDisplayAndCreationState()
            Dim choiceType As Type = MainType("CivilEnginStructures.BridgeStructureChoice")

            Assert.Multiple(
                Sub()
                    Assert.That(choiceType.GetProperty("Id"), [Is].Not.Null)
                    Assert.That(choiceType.GetProperty("DisplayName"), [Is].Not.Null)
                    Assert.That(choiceType.GetProperty("IsNew"), [Is].Not.Null)
                    Assert.That(choiceType.GetMethod("CreateChoices", BindingFlags.Public Or BindingFlags.Static),
                                [Is].Not.Null)
                    Assert.That(choiceType.GetMethod("ResolveSavedName", BindingFlags.Public Or BindingFlags.Static),
                                [Is].Not.Null)
                End Sub)
        End Sub

        <Test>
        Public Sub CreateChoicesPreservesSourceOrderAndMarksOnlyExplicitTransientIdAsNew()
            Dim choices As IList = CreateChoices(
                New KeyValuePair(Of String, String)() {
                    Pair("bridge-b", "Второе сооружение"),
                    Pair("bridge-new", String.Empty),
                    Pair("bridge-a", "Первое сооружение")
                },
                "bridge-new")

            Assert.Multiple(
                Sub()
                    Assert.That(choices.Count, [Is].EqualTo(3))
                    Assert.That(ChoiceValue(choices(0), "Id"), [Is].EqualTo("bridge-b"))
                    Assert.That(ChoiceValue(choices(1), "Id"), [Is].EqualTo("bridge-new"))
                    Assert.That(ChoiceValue(choices(2), "Id"), [Is].EqualTo("bridge-a"))
                    Assert.That(ChoiceFlag(choices(0), "IsNew"), [Is].False)
                    Assert.That(ChoiceFlag(choices(1), "IsNew"), [Is].True)
                    Assert.That(ChoiceFlag(choices(2), "IsNew"), [Is].False)
                    Assert.That(ChoiceValue(choices(1), "DisplayName"), [Is].EqualTo(NewStructureCaption))
                End Sub)
        End Sub

        <Test>
        Public Sub ExistingBlankAndLegacySentinelNamesReceiveStableUniqueDisplayNames()
            Dim choices As IList = CreateChoices(
                New KeyValuePair(Of String, String)() {
                    Pair("blank", String.Empty),
                    Pair("reserved", "Искусственное сооружение 1"),
                    Pair("legacy", NewStructureCaption),
                    Pair("custom-a", "Мост через реку"),
                    Pair("custom-b", "Мост через реку"),
                    Pair("new", String.Empty)
                },
                "new")

            Assert.Multiple(
                Sub()
                    Assert.That(ChoiceValue(choices(0), "DisplayName"), [Is].EqualTo("Искусственное сооружение 2"))
                    Assert.That(ChoiceValue(choices(1), "DisplayName"), [Is].EqualTo("Искусственное сооружение 1"))
                    Assert.That(ChoiceValue(choices(2), "DisplayName"), [Is].EqualTo("Искусственное сооружение 3"))
                    Assert.That(ChoiceValue(choices(3), "DisplayName"), [Is].EqualTo("Мост через реку"))
                    Assert.That(ChoiceValue(choices(4), "DisplayName"), [Is].EqualTo("Мост через реку"))
                    Assert.That(ChoiceValue(choices(5), "DisplayName"), [Is].EqualTo(NewStructureCaption))
                End Sub)
        End Sub

        <Test>
        Public Sub ExistingSelectionResolvesBlankAndSentinelInputToItsAssignedDisplayName()
            Dim choices As IList = CreateChoices(
                New KeyValuePair(Of String, String)() {
                    Pair("existing", String.Empty),
                    Pair("new", String.Empty)
                },
                "new")

            Assert.Multiple(
                Sub()
                    Assert.That(ResolveSavedName(String.Empty, choices(0), choices),
                                [Is].EqualTo("Искусственное сооружение 1"))
                    Assert.That(ResolveSavedName(NewStructureCaption, choices(0), choices),
                                [Is].EqualTo("Искусственное сооружение 1"))
                End Sub)
        End Sub

        <Test>
        Public Sub NewSelectionResolvesBlankOrSentinelInputToNextAvailableGeneratedName()
            Dim choices As IList = CreateChoices(
                New KeyValuePair(Of String, String)() {
                    Pair("reserved", "Искусственное сооружение 1"),
                    Pair("existing", String.Empty),
                    Pair("new", String.Empty)
                },
                "new")

            Assert.Multiple(
                Sub()
                    Assert.That(ResolveSavedName(String.Empty, choices(2), choices),
                                [Is].EqualTo("Искусственное сооружение 3"))
                    Assert.That(ResolveSavedName(NewStructureCaption, choices(2), choices),
                                [Is].EqualTo("Искусственное сооружение 3"))
                End Sub)
        End Sub

        <Test>
        Public Sub ResolveSavedNamePreservesCustomUserInput()
            Dim choices As IList = CreateChoices(
                New KeyValuePair(Of String, String)() {Pair("new", String.Empty)},
                "new")

            Assert.That(ResolveSavedName("Мост через Волгу", choices(0), choices),
                        [Is].EqualTo("Мост через Волгу"))
        End Sub

        <Test>
        Public Sub SupportedSurfacePathFilterAcceptsOnlySurfaceOwningModelExtensions()
            Assert.Multiple(
                Sub()
                    Assert.That(IsSupportedSurfacePath("models\terrain.SFCX"), [Is].True)
                    Assert.That(IsSupportedSurfacePath("models\road.roadx"), [Is].True)
                    Assert.That(IsSupportedSurfacePath("models\site.site"), [Is].True)
                    Assert.That(IsSupportedSurfacePath("models\alignment.AlGx"), [Is].True)
                    Assert.That(IsSupportedSurfacePath("models\exchange.ifc"), [Is].False)
                    Assert.That(IsSupportedSurfacePath("models\bridge.infrabridgex"), [Is].False)
                    Assert.That(IsSupportedSurfacePath("models\legacy.arrx"), [Is].False)
                    Assert.That(IsSupportedSurfacePath("models\drawing.dwg"), [Is].False)
                    Assert.That(IsSupportedSurfacePath("project.rbprojx"), [Is].False)
                    Assert.That(IsSupportedSurfacePath(String.Empty), [Is].False)
                    Assert.That(IsSupportedSurfacePath(Nothing), [Is].False)
                End Sub)
        End Sub

        Private Shared Function Pair(id As String, name As String) As KeyValuePair(Of String, String)
            Return New KeyValuePair(Of String, String)(id, name)
        End Function

        Private Shared Function CreateChoices(names As KeyValuePair(Of String, String)(),
                                               newBridgeId As String) As IList
            Dim choiceType As Type = MainType("CivilEnginStructures.BridgeStructureChoice")
            Return DirectCast(choiceType.GetMethod("CreateChoices", BindingFlags.Public Or BindingFlags.Static).
                              Invoke(Nothing, New Object() {names, newBridgeId}), IList)
        End Function

        Private Shared Function ResolveSavedName(inputText As String,
                                                 selected As Object,
                                                 choices As IList) As String
            Dim choiceType As Type = MainType("CivilEnginStructures.BridgeStructureChoice")
            Return CStr(choiceType.GetMethod("ResolveSavedName", BindingFlags.Public Or BindingFlags.Static).
                        Invoke(Nothing, New Object() {inputText, selected, choices}))
        End Function

        Private Shared Function ChoiceValue(choice As Object, propertyName As String) As String
            Return CStr(choice.GetType().GetProperty(propertyName).GetValue(choice, Nothing))
        End Function

        Private Shared Function ChoiceFlag(choice As Object, propertyName As String) As Boolean
            Return CBool(choice.GetType().GetProperty(propertyName).GetValue(choice, Nothing))
        End Function

        Private Shared Function IsSupportedSurfacePath(path As String) As Boolean
            Dim choiceType As Type = MainType("CivilEnginStructures.SurfaceModelChoice")
            Dim method As MethodInfo = choiceType.GetMethod("IsSupportedSurfacePath",
                                                           BindingFlags.Public Or BindingFlags.Static)
            If method Is Nothing Then
                Assert.Fail("SurfaceModelChoice.IsSupportedSurfacePath must exist in the production assembly.")
                Return False
            End If
            Return CBool(method.Invoke(Nothing, New Object() {path}))
        End Function

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
                        "Build CivilEnginStructures.vbproj before running these contract tests.")
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
