Imports System.Collections.Generic
Imports System.IO
Imports System.Text.RegularExpressions
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class FormCreateMiddlePillarsImageToolTipContractTests
        Private Const FormRelativePath As String = "UserForms\Bridge\FormCreateMiddlePillars.vb"
        Private Const RegexOptionsValue As RegexOptions = RegexOptions.IgnoreCase Or
                                                           RegexOptions.CultureInvariant Or
                                                           RegexOptions.Singleline

        <Test>
        Public Sub ConstructorCreatesImageToolTipAndPlacesItOverPictureBox1()
            Dim source As String = ReadFormSource()
            Dim ownerName As String = FindToolTipOwnerName(source)
            If ownerName Is Nothing Then Return

            Dim helperName As String = FindImageToolTipName(source, ownerName)
            If helperName Is Nothing Then Return

            Dim placementPattern As String =
                Regex.Escape(helperName) &
                "\s*\.\s*SetPlacementBounds\s*\(\s*New\s+Rectangle\s*\(\s*" &
                "TabControl2\s*\.\s*DisplayRectangle\s*\.\s*Left\s*\+\s*PictureBox1\s*\.\s*Left\s*,\s*" &
                "TabControl2\s*\.\s*DisplayRectangle\s*\.\s*Top\s*\+\s*PictureBox1\s*\.\s*Top\s*,\s*" &
                "PictureBox1\s*\.\s*Width\s*,\s*PictureBox1\s*\.\s*Height\s*\)\s*\)"

            Assert.That(Regex.IsMatch(source, placementPattern, RegexOptionsValue), [Is].True,
                        "The image tooltip must be placed over PictureBox1 in TabControl2 coordinates.")
        End Sub

        <Test>
        Public Sub ConstructorRegistersExactSixMiddlePillarImages()
            Dim source As String = ReadFormSource()
            Dim ownerName As String = FindToolTipOwnerName(source)
            If ownerName Is Nothing Then Return

            Dim helperName As String = FindImageToolTipName(source, ownerName)
            If helperName Is Nothing Then Return

            Dim expected As New Dictionary(Of String, String) From {
                {"TabPage1", "\FileResources\ImageObject\Bridge\Pillars\RigelPillar.PNG"},
                {"TabPage9", "\FileResources\ImageObject\Bridge\Pillars\SubFermenterPillar.PNG"},
                {"TabPage2", "\FileResources\ImageObject\Bridge\Pillars\RackPillar.PNG"},
                {"TabPage4", "\FileResources\ImageObject\Bridge\Pillars\GrillagePillar.PNG"},
                {"TabPage3", "\FileResources\ImageObject\Bridge\Pillars\PreparationPillar.PNG"},
                {"TabPage5", "\FileResources\ImageObject\Bridge\Pillars\PilePillar.PNG"}
            }
            Dim registrationPattern As String =
                Regex.Escape(helperName) &
                "\s*\.\s*Register\s*\(\s*(?<tab>TabPage\d+)\s*,\s*generalDir\s*&\s*""(?<path>[^""]+)""\s*\)"
            Dim matches As MatchCollection = Regex.Matches(source, registrationPattern, RegexOptionsValue)
            Dim actual As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)

            For Each registration As Match In matches
                actual(registration.Groups("tab").Value) = registration.Groups("path").Value
            Next

            Assert.Multiple(
                Sub()
                    Assert.That(matches.Count, [Is].EqualTo(expected.Count),
                                "Exactly six middle-pillar parameter tabs must be registered.")
                    For Each mapping As KeyValuePair(Of String, String) In expected
                        Assert.That(actual.ContainsKey(mapping.Key), [Is].True,
                                    mapping.Key & " must be registered.")
                        If actual.ContainsKey(mapping.Key) Then
                            Assert.That(actual(mapping.Key), [Is].EqualTo(mapping.Value).IgnoreCase,
                                        mapping.Key & " must use its exact pillar PNG path.")
                        End If
                    Next
                End Sub)
        End Sub

        <Test>
        Public Sub FormClosedAndDisposedReleaseImageToolTipAndOwnedToolTip()
            Dim source As String = ReadFormSource()
            Dim ownerName As String = FindToolTipOwnerName(source)
            If ownerName Is Nothing Then Return

            Dim helperName As String = FindImageToolTipName(source, ownerName)
            If helperName Is Nothing Then Return

            Dim closedHandler As String = FindEventHandlerName(source, "FormClosed")
            Dim disposedHandler As String = FindEventHandlerName(source, "Disposed")
            If closedHandler Is Nothing OrElse disposedHandler Is Nothing Then Return

            Dim closedBody As String = FindMethodBody(source, closedHandler)
            Dim cleanupMatch As Match = Regex.Match(closedBody,
                                                    "(?<cleanup>[A-Za-z_]\w*Dispose[A-Za-z_]\w*)\s*\(\s*\)",
                                                    RegexOptionsValue)
            Assert.That(cleanupMatch.Success, [Is].True,
                        "The FormClosed handler must call a shared cleanup method.")
            If Not cleanupMatch.Success Then Return

            Dim cleanupName As String = cleanupMatch.Groups("cleanup").Value
            Dim disposedBody As String = FindMethodBody(source, disposedHandler)
            Assert.That(Regex.IsMatch(disposedBody,
                                      Regex.Escape(cleanupName) & "\s*\(\s*\)",
                                      RegexOptionsValue), [Is].True,
                        "The Disposed handler must call the same cleanup method.")

            Dim cleanupBody As String = FindMethodBody(source, cleanupName)
            Assert.Multiple(
                Sub()
                    Assert.That(Regex.IsMatch(cleanupBody,
                                              Regex.Escape(helperName) & "\s*\.\s*Dispose\s*\(\s*\)",
                                              RegexOptionsValue), [Is].True,
                                "Cleanup must dispose TabPageImageToolTip.")
                    Assert.That(Regex.IsMatch(cleanupBody,
                                              Regex.Escape(helperName) & "\s*=\s*Nothing",
                                              RegexOptionsValue), [Is].True,
                                "Cleanup must clear the TabPageImageToolTip reference.")
                    Assert.That(Regex.IsMatch(cleanupBody,
                                              Regex.Escape(ownerName) & "\s*\.\s*Dispose\s*\(\s*\)",
                                              RegexOptionsValue), [Is].True,
                                "Cleanup must dispose the form-owned ToolTip.")
                    Assert.That(Regex.IsMatch(cleanupBody,
                                              Regex.Escape(ownerName) & "\s*=\s*Nothing",
                                              RegexOptionsValue), [Is].True,
                                "Cleanup must clear the form-owned ToolTip reference.")
                End Sub)
        End Sub

        Private Shared Function FindToolTipOwnerName(source As String) As String
            Dim ownerMatch As Match = Regex.Match(
                source,
                "(?<owner>[A-Za-z_]\w*)\s*=\s*New\s+(?:System\s*\.\s*Windows\s*\.\s*Forms\s*\.\s*)?ToolTip\s*(?:\(\s*\))?",
                RegexOptionsValue)
            Assert.That(ownerMatch.Success, [Is].True,
                        "FormCreateMiddlePillars must create a ToolTip owned by the form.")
            If Not ownerMatch.Success Then Return Nothing
            Return ownerMatch.Groups("owner").Value
        End Function

        Private Shared Function FindImageToolTipName(source As String, ownerName As String) As String
            Dim helperMatch As Match = Regex.Match(
                source,
                "(?<helper>[A-Za-z_]\w*)\s*=\s*New\s+TabPageImageToolTip\s*\(\s*" &
                Regex.Escape(ownerName) & "\s*,\s*TabControl1\s*,\s*TabControl2\s*\)",
                RegexOptionsValue)
            Assert.That(helperMatch.Success, [Is].True,
                        "TabPageImageToolTip must observe TabControl1 and TabControl2.")
            If Not helperMatch.Success Then Return Nothing
            Return helperMatch.Groups("helper").Value
        End Function

        Private Shared Function FindEventHandlerName(source As String, eventName As String) As String
            Dim handlerMatch As Match = Regex.Match(
                source,
                "AddHandler\s+Me\s*\.\s*" & Regex.Escape(eventName) &
                "\s*,\s*AddressOf\s+(?<handler>[A-Za-z_]\w*)",
                RegexOptionsValue)
            Assert.That(handlerMatch.Success, [Is].True,
                        eventName & " must be connected to image-tooltip cleanup.")
            If Not handlerMatch.Success Then Return Nothing
            Return handlerMatch.Groups("handler").Value
        End Function

        Private Shared Function FindMethodBody(source As String, methodName As String) As String
            Dim methodMatch As Match = Regex.Match(
                source,
                "(?:Private|Friend|Public)\s+Sub\s+" & Regex.Escape(methodName) &
                "\s*\([^)]*\)\s*(?<body>.*?)\bEnd\s+Sub\b",
                RegexOptionsValue)
            Assert.That(methodMatch.Success, [Is].True,
                        "Could not find method body for " & methodName & ".")
            If Not methodMatch.Success Then Return String.Empty
            Return methodMatch.Groups("body").Value
        End Function

        Private Shared Function ReadFormSource() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.WorkDirectory)

            While directory IsNot Nothing
                Dim candidate As String = Path.Combine(directory.FullName, FormRelativePath)
                If File.Exists(candidate) Then Return File.ReadAllText(candidate)
                directory = directory.Parent
            End While

            Assert.Fail("Could not locate " & FormRelativePath &
                        " above test work directory " & TestContext.CurrentContext.WorkDirectory & ".")
            Return String.Empty
        End Function
    End Class
End Namespace
