Imports System.IO
Imports System.Text.RegularExpressions
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class CalculationBeamsGetBeamsToRowContractTests
        Private Const CalculationBeamsRelativePath As String =
            "CivilEngineeringStructures\BridgeProject\Bridge\Beams\CalculationBeams.vb"
        Private Const RegexOptionsValue As RegexOptions = RegexOptions.IgnoreCase Or
                                                           RegexOptions.CultureInvariant Or
                                                           RegexOptions.Singleline

        <Test>
        Public Sub GetBeamsToRowAcceptsNestedBeamDictionary()
            Dim methodSource As String = ReadGetBeamsToRowSource()
            Dim signaturePattern As String =
                "Public\s+Shared\s+Function\s+getBeamsToRow\s*\(\s*" &
                "ByVal\s+dictionaryBeams\s+As\s+" &
                "Dictionary\s*\(\s*Of\s+Integer\s*,\s*" &
                "Dictionary\s*\(\s*Of\s+Integer\s*,\s*StructureElement\s*\)\s*\)\s*,\s*" &
                "ByVal\s+numberRow\s+As\s+Integer\s*\)\s+As\s+" &
                "List\s*\(\s*Of\s+StructureElement\s*\)"

            Assert.That(Regex.IsMatch(methodSource, signaturePattern, RegexOptionsValue), [Is].True,
                        "getBeamsToRow must accept the nested dictionary returned by Bridges.getBeams.")
        End Sub

        <Test>
        Public Sub GetBeamsToRowTraversesEveryInnerDictionary()
            Dim methodSource As String = ReadGetBeamsToRowSource()
            Dim outerLoop As Match = Regex.Match(
                methodSource,
                "For\s+Each\s+(?<inner>[A-Za-z_]\w*)\s+As\s+" &
                "Dictionary\s*\(\s*Of\s+Integer\s*,\s*StructureElement\s*\)\s+" &
                "In\s+dictionaryBeams\s*\.\s*Values",
                RegexOptionsValue)

            Assert.That(outerLoop.Success, [Is].True,
                        "The outer dictionary values must be traversed as inner beam dictionaries.")
            If Not outerLoop.Success Then Return

            Dim innerName As String = Regex.Escape(outerLoop.Groups("inner").Value)
            Dim innerLoopPattern As String =
                "For\s+Each\s+dataBeam\s+As\s+StructureElement\s+In\s+" &
                innerName & "\s*\.\s*Values"

            Assert.That(Regex.IsMatch(methodSource, innerLoopPattern, RegexOptionsValue), [Is].True,
                        "Every non-empty inner dictionary must contribute its StructureElement values.")
        End Sub

        <Test>
        Public Sub GetBeamsToRowAddsOnlyBeamsWhoseNumberRowMatches()
            Dim methodSource As String = ReadGetBeamsToRowSource()
            Dim matchingRowPattern As String =
                "If\s+userBeam\s*\.\s*numberRow\s*=\s*numberRow\s+Then\s*" &
                "result\s*\.\s*Add\s*\(\s*dataBeam\s*\)"

            Assert.That(Regex.IsMatch(methodSource, matchingRowPattern, RegexOptionsValue), [Is].True,
                        "A StructureElement may be added only after BeamI.numberRow matches numberRow.")
            Assert.That(Regex.Matches(methodSource,
                                      "result\s*\.\s*Add\s*\(",
                                      RegexOptionsValue).Count,
                        [Is].EqualTo(1),
                        "There must be no unfiltered path that adds a beam to the result.")
        End Sub

        <Test>
        Public Sub GetBeamsToRowSkipsNothingEmptyInnerDictionariesAndInvalidBeamData()
            Dim methodSource As String = ReadGetBeamsToRowSource()
            Dim outerLoop As Match = Regex.Match(
                methodSource,
                "For\s+Each\s+(?<inner>[A-Za-z_]\w*)\s+As\s+" &
                "Dictionary\s*\(\s*Of\s+Integer\s*,\s*StructureElement\s*\)\s+" &
                "In\s+dictionaryBeams\s*\.\s*Values",
                RegexOptionsValue)

            Assert.That(outerLoop.Success, [Is].True,
                        "An inner dictionary loop is required before null handling can be verified.")
            If Not outerLoop.Success Then Return

            Dim innerName As String = Regex.Escape(outerLoop.Groups("inner").Value)
            Dim nonNothingInnerPattern As String =
                "(?:IsNothing\s*\(\s*" & innerName & "\s*\)\s*=\s*False|" &
                "Not\s+IsNothing\s*\(\s*" & innerName & "\s*\)|" &
                innerName & "\s+IsNot\s+Nothing)"
            Dim invalidBeamPattern As String =
                "Dim\s+userBeam\s+As\s+BeamI\s*=\s*dataBeam\s*\.\s*getBeamI\s*\(\s*\)\s*" &
                "If\s+(?:IsNothing\s*\(\s*userBeam\s*\)\s*=\s*False|" &
                "Not\s+IsNothing\s*\(\s*userBeam\s*\)|userBeam\s+IsNot\s+Nothing)\s+Then"

            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(methodSource, nonNothingInnerPattern, RegexOptionsValue), [Is].True,
                                "Nothing inner dictionaries must be skipped before reading Values; empty ones are skipped by the inner For Each.")
                    Assert.That(Regex.IsMatch(methodSource,
                                              "(?:IsNothing\s*\(\s*dataBeam\s*\)\s*=\s*False|" &
                                              "Not\s+IsNothing\s*\(\s*dataBeam\s*\)|" &
                                              "dataBeam\s+IsNot\s+Nothing)",
                                              RegexOptionsValue), [Is].True,
                                "Nothing StructureElement values must be skipped.")
                    Assert.That(Regex.IsMatch(methodSource, invalidBeamPattern, RegexOptionsValue), [Is].True,
                                "Empty or invalid KeyParameter data makes getBeamI return Nothing and must be skipped.")
                End Sub)
        End Sub

        <Test>
        Public Sub GetBeamsToRowReturnsEmptyListWhenRequestedRowIsAbsent()
            Dim methodSource As String = ReadGetBeamsToRowSource()

            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(methodSource,
                                              "Dim\s+result\s+As\s+List\s*\(\s*Of\s+StructureElement\s*\)\s*=\s*" &
                                              "New\s+List\s*\(\s*Of\s+StructureElement\s*\)",
                                              RegexOptionsValue), [Is].True,
                                "The result must start as an empty list.")
                    Assert.That(Regex.IsMatch(methodSource,
                                              "Return\s+result\s*$",
                                              RegexOptionsValue), [Is].True,
                                "When no row matches, the unchanged empty result must be returned.")
                End Sub)
        End Sub

        Private Shared Function ReadGetBeamsToRowSource() As String
            Dim source As String = ReadCalculationBeamsSource()
            Dim methodMatch As Match = Regex.Match(
                source,
                "Public\s+Shared\s+Function\s+getBeamsToRow\b(?<method>.*?)End\s+Function",
                RegexOptionsValue)

            Assert.That(methodMatch.Success, [Is].True,
                        "Could not find CalculationBeams.getBeamsToRow in the active checkout.")
            If Not methodMatch.Success Then Return String.Empty
            Return "Public Shared Function getBeamsToRow" & methodMatch.Groups("method").Value
        End Function

        Private Shared Function ReadCalculationBeamsSource() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.WorkDirectory)

            While directory IsNot Nothing
                Dim candidate As String = Path.Combine(directory.FullName, CalculationBeamsRelativePath)
                If File.Exists(candidate) Then Return File.ReadAllText(candidate)
                directory = directory.Parent
            End While

            Assert.Fail("Could not locate " & CalculationBeamsRelativePath &
                        " above test work directory " & TestContext.CurrentContext.WorkDirectory & ".")
            Return String.Empty
        End Function
    End Class
End Namespace
