Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class FormCreateMiddlePillarsImageToolTipContractTests
        Private Const FormRelativePath As String = "UserForms\Bridge\FormCreateMiddlePillars.vb"
        Private Const RegexOptionsValue As RegexOptions = RegexOptions.IgnoreCase Or
                                                           RegexOptions.CultureInvariant Or
                                                           RegexOptions.Singleline

        <Test>
        Public Sub ConstructorRegistersExactSixExplicitParameterSchemesWithoutHoverToolTip()
            Dim source As String = ReadFormSource()
            Dim expected As New Dictionary(Of String, String) From {
                {"TabPage1", "\FileResources\ImageObject\Bridge\Pillars\RigelPillar.PNG"},
                {"TabPage9", "\FileResources\ImageObject\Bridge\Pillars\SubFermenterPillar.PNG"},
                {"TabPage2", "\FileResources\ImageObject\Bridge\Pillars\RackPillar.PNG"},
                {"TabPage4", "\FileResources\ImageObject\Bridge\Pillars\GrillagePillar.PNG"},
                {"TabPage3", "\FileResources\ImageObject\Bridge\Pillars\PreparationPillar.PNG"},
                {"TabPage5", "\FileResources\ImageObject\Bridge\Pillars\PilePillar.PNG"}
            }
            Dim pattern As String =
                "RegisterParameterScheme\s*\(\s*(?<tab>TabPage\d+)\s*,\s*generalDir\s*&\s*""(?<path>[^""]+)""\s*\)"
            Dim matches As MatchCollection = Regex.Matches(source, pattern, RegexOptionsValue)
            Dim actual As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each registration As Match In matches
                actual(registration.Groups("tab").Value) = registration.Groups("path").Value
            Next

            Assert.Multiple(
                Sub()
                    Assert.That(matches.Count, [Is].EqualTo(expected.Count))
                    For Each mapping As KeyValuePair(Of String, String) In expected
                        Assert.That(actual.ContainsKey(mapping.Key), [Is].True, mapping.Key)
                        If actual.ContainsKey(mapping.Key) Then
                            Assert.That(actual(mapping.Key), [Is].EqualTo(mapping.Value).IgnoreCase, mapping.Key)
                        End If
                    Next
                    Assert.That(source, Does.Not.Contain("New TabPageImageToolTip"),
                                "Parameter help must open only from the explicit button, without a hover overlay.")
                End Sub)
        End Sub

        <Test>
        Public Sub SchemeButtonOpensContainedPanelChangesWithTabAndCanClose()
            Dim firstPath As String = Path.Combine(TestContext.CurrentContext.WorkDirectory,
                                                   Guid.NewGuid().ToString("N") & "-first.png")
            Dim secondPath As String = Path.Combine(TestContext.CurrentContext.WorkDirectory,
                                                    Guid.NewGuid().ToString("N") & "-second.png")
            SaveImage(firstPath, New Size(31, 17), Color.Red)
            SaveImage(secondPath, New Size(43, 19), Color.Blue)
            Try
                Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                    Invoke(form, "ApplyModernAppearance")
                    Invoke(form, "RegisterParameterScheme", form.TabPage1, firstPath)
                    Invoke(form, "RegisterParameterScheme", form.TabPage9, secondPath)
                    form.ShowInTaskbar = False
                    form.StartPosition = FormStartPosition.Manual
                    form.Location = New Point(-30000, -30000)
                    form.Show()
                    Application.DoEvents()

                    Dim button As Button = DirectCast(RequireControl(form, "ModernParameterSchemeButton"), Button)
                    Dim panel As Panel = DirectCast(RequireControl(form, "ModernParameterSchemePanel"), Panel)
                    Dim closeButton As Button = DirectCast(RequireControl(form, "ModernParameterSchemeClose"), Button)
                    Dim picture As PictureBox = DirectCast(RequireControl(form, "ModernParameterSchemeImage"), PictureBox)
                    Assert.That(panel.Parent, [Is].SameAs(form))
                    Assert.That(panel.Visible, [Is].False)

                    form.TabControl1.SelectedTab = form.TabPage1
                    button.PerformClick()
                    Application.DoEvents()
                    Assert.Multiple(
                        Sub()
                            Assert.That(panel.Visible, [Is].True)
                            Assert.That(picture.Image, [Is].Not.Null)
                            Assert.That(picture.Image.Size, [Is].EqualTo(New Size(31, 17)))
                            Assert.That(form.TabControl1.Bounds.Contains(panel.Bounds), [Is].True,
                                        "The root-level panel must stay inside the parameter tab bounds.")
                            Assert.That(panel.Bounds.IntersectsWith(form.TabControl2.Bounds), [Is].False,
                                        "Parameter help must not cover the preview area.")
                        End Sub)

                    form.TabControl1.SelectedTab = form.TabPage9
                    Application.DoEvents()
                    Assert.Multiple(
                        Sub()
                            Assert.That(panel.Visible, [Is].False)
                            Assert.That(picture.Image, [Is].Null,
                                        "Changing the parameter tab must close and release the old scheme.")
                        End Sub)

                    button.PerformClick()
                    Application.DoEvents()
                    Assert.Multiple(
                        Sub()
                            Assert.That(panel.Visible, [Is].True)
                            Assert.That(picture.Image, [Is].Not.Null)
                            Assert.That(picture.Image.Size, [Is].EqualTo(New Size(43, 19)))
                        End Sub)

                    closeButton.PerformClick()
                    Application.DoEvents()
                    Assert.Multiple(
                        Sub()
                            Assert.That(panel.Visible, [Is].False)
                            Assert.That(picture.Image, [Is].Null)
                        End Sub)
                End Using
            Finally
                If File.Exists(firstPath) Then File.Delete(firstPath)
                If File.Exists(secondPath) Then File.Delete(secondPath)
            End Try
        End Sub

        Private Shared Sub SaveImage(path As String, size As Size, color As Color)
            Using bitmap As New Bitmap(size.Width, size.Height), graphics As Graphics = Graphics.FromImage(bitmap)
                graphics.Clear(color)
                bitmap.Save(path)
            End Using
        End Sub

        Private Shared Sub Invoke(target As Object, methodName As String, ParamArray arguments As Object())
            Dim method As MethodInfo = target.GetType().GetMethods(
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic).
                FirstOrDefault(Function(candidate) candidate.Name = methodName AndAlso
                                                   candidate.GetParameters().Length = arguments.Length)
            If method Is Nothing Then Throw New AssertionException(methodName & " is missing.")
            Try
                method.Invoke(target, arguments)
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Function RequireControl(parent As Control, name As String) As Control
            Dim result As Control = Descendants(parent).
                FirstOrDefault(Function(control) String.Equals(control.Name, name, StringComparison.Ordinal))
            If result Is Nothing Then Throw New AssertionException(name & " control is missing.")
            Return result
        End Function

        Private Shared Function Descendants(parent As Control) As IEnumerable(Of Control)
            Dim result As New List(Of Control)()
            For Each child As Control In parent.Controls
                result.Add(child)
                result.AddRange(Descendants(child))
            Next
            Return result
        End Function

        Private Shared Function ReadFormSource() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.WorkDirectory)
            While directory IsNot Nothing
                Dim candidate As String = Path.Combine(directory.FullName, FormRelativePath)
                If File.Exists(candidate) Then Return File.ReadAllText(candidate)
                directory = directory.Parent
            End While
            Assert.Fail("Could not locate " & FormRelativePath & ".")
            Return String.Empty
        End Function
    End Class
End Namespace
