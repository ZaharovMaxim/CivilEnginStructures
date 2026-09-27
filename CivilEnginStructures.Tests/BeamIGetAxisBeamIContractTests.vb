Imports System.IO
Imports System.Text.RegularExpressions
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class BeamIGetAxisBeamIContractTests
        Private Const BeamIRelativePath As String =
            "CivilEngineeringStructures\BridgeProject\Bridge\Beams\BeamI.vb"
        Private Const RegexOptionsValue As RegexOptions = RegexOptions.IgnoreCase Or
                                                           RegexOptions.CultureInvariant Or
                                                           RegexOptions.Singleline

        <Test>
        Public Sub GetAxisBeamIExposesRemovalSwitchDefaultingToFalse()
            Dim methodSource As String = ReadGetAxisBeamISource()
            Dim signaturePattern As String =
                "Public\s+Shared\s+Function\s+getAxisBeamI\s*\(\s*" &
                "ByVal\s+dictionaryObjectsBridge\s+As\s+" &
                "Dictionary\s*\(\s*Of\s+StructureElement\s*\.\s*typeObject\s*,\s*" &
                "List\s*\(\s*Of\s+StructureElement\s*\)\s*\)\s*,\s*" &
                "ByVal\s+numberProlet\s+As\s+Integer\s*,\s*" &
                "ByVal\s+numberRow\s+As\s+Integer\s*,\s*" &
                "Optional\s+(?:ByVal\s+)?removeDictionary\s+As\s+Boolean\s*=\s*False\s*\)\s+" &
                "As\s+StructureElement"

            Assert.That(Regex.IsMatch(methodSource, signaturePattern, RegexOptionsValue), [Is].True,
                        "Removal from the axisBeam list must be opt-in and disabled by default.")
        End Sub

        <Test>
        Public Sub GetAxisBeamIReturnsTheMatchedStructureElement()
            Dim methodSource As String = ReadGetAxisBeamISource()
            Dim matchAndResultPattern As String =
                "If\s+numberProlet\s*=\s*userBeam\s*\.\s*numberProlet\s+And\s+" &
                "numberRow\s*=\s*userBeam\s*\.\s*numberRow\s+Then\s*" &
                "(?<result>[A-Za-z_]\w*)\s*=\s*tempData"
            Dim resultMatch As Match = Regex.Match(methodSource,
                                                   matchAndResultPattern,
                                                   RegexOptionsValue)

            Assert.That(resultMatch.Success, [Is].True,
                        "The matching axis StructureElement must be saved as the function result.")
            If Not resultMatch.Success Then Return

            Dim resultName As String = Regex.Escape(resultMatch.Groups("result").Value)
            Assert.That(Regex.IsMatch(methodSource,
                                      "Return\s+" & resultName & "\s*$",
                                      RegexOptionsValue),
                        [Is].True,
                        "The same matched StructureElement must be returned after optional removal.")
        End Sub

        <Test>
        Public Sub GetAxisBeamIRemovesOnlyTheMatchedItemWhenRequested()
            Dim methodSource As String = ReadGetAxisBeamISource()
            Dim guardedRemovalPattern As String =
                "If\s+removeDictionary(?:\s*=\s*True)?\s+Then\s*" &
                "listAxisBeam\s*\.\s*(?:RemoveAt\s*\(\s*k\s*\)|" &
                "Remove\s*\(\s*tempData\s*\))\s*" &
                "End\s+If"
            Dim listMutationPattern As String =
                "listAxisBeam\s*\.\s*(?:RemoveAt|Remove|Clear)\s*\("

            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(methodSource,
                                              guardedRemovalPattern,
                                              RegexOptionsValue),
                                [Is].True,
                                "Only removeDictionary=True may remove the matched item from listAxisBeam.")
                    Assert.That(Regex.Matches(methodSource,
                                              listMutationPattern,
                                              RegexOptionsValue).Count,
                                [Is].EqualTo(1),
                                "getAxisBeamI must not perform any additional mutation of the axisBeam list.")
                End Sub)
        End Sub

        <Test>
        Public Sub GetAxisBeamIDoesNotTouchTheDrawingEntity()
            Dim methodSource As String = ReadGetAxisBeamISource()

            Assert.Multiple(
                Sub()
                    Assert.That(methodSource, Does.Not.Match("\bDWGEntity\b").IgnoreCase,
                                "Dictionary removal must not access StructureElement.DWGEntity.")
                    Assert.That(methodSource, Does.Not.Match("\bActiveSpace\b").IgnoreCase,
                                "Dictionary removal must not access the active drawing space.")
                    Assert.That(methodSource,
                                Does.Not.Match("Entities\s*\.\s*Remove\s*\(").IgnoreCase,
                                "Dictionary removal must not remove any entity from the drawing.")
                End Sub)
        End Sub

        <Test>
        Public Sub GetAxisBeamIPreservesExistingInvalidInputGuards()
            Dim methodSource As String = ReadGetAxisBeamISource()

            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(methodSource,
                                              "If\s+IsNothing\s*\(\s*dictionaryObjectsBridge\s*\)\s*=\s*True\s+Then\s+Return\s+Nothing",
                                              RegexOptionsValue),
                                [Is].True,
                                "A Nothing dictionary must still return Nothing without mutation.")
                    Assert.That(Regex.IsMatch(methodSource,
                                              "If\s+numberProlet\s*<\s*1\s+Then\s+Return\s+Nothing",
                                              RegexOptionsValue),
                                [Is].True,
                                "An invalid span number must still return Nothing without mutation.")
                    Assert.That(Regex.IsMatch(methodSource,
                                              "ContainsKey\s*\(\s*StructureElement\s*\.\s*typeObject\s*\.\s*axisBeam\s*\)",
                                              RegexOptionsValue),
                                [Is].True,
                                "The axisBeam list must be checked before it is read.")
                End Sub)
        End Sub

        Private Shared Function ReadGetAxisBeamISource() As String
            Dim source As String = ReadBeamISource()
            Dim methodMatch As Match = Regex.Match(
                source,
                "Public\s+Shared\s+Function\s+getAxisBeamI\b(?<method>.*?)End\s+Function",
                RegexOptionsValue)

            Assert.That(methodMatch.Success, [Is].True,
                        "Could not find BeamI.getAxisBeamI in the active checkout.")
            If Not methodMatch.Success Then Return String.Empty
            Return "Public Shared Function getAxisBeamI" & methodMatch.Groups("method").Value
        End Function

        Private Shared Function ReadBeamISource() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.WorkDirectory)

            While directory IsNot Nothing
                Dim candidate As String = Path.Combine(directory.FullName, BeamIRelativePath)
                If File.Exists(candidate) Then Return File.ReadAllText(candidate)
                directory = directory.Parent
            End While

            Assert.Fail("Could not locate " & BeamIRelativePath &
                        " above test work directory " & TestContext.CurrentContext.WorkDirectory & ".")
            Return String.Empty
        End Function
    End Class
End Namespace
