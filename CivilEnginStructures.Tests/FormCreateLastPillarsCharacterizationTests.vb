Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class FormCreateLastPillarsCharacterizationTests
        Private Const FormRelativePath As String = "UserForms\Bridge\FormCreateLastPillars.vb"

        <Test, Category("Legacy")>
        Public Sub DesignerInitializesFunctionalDefaultsAndAllElevenGrids()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                AssertNumeric(form.NUpD_ElevationLand, -99999999999D, 9999999999D, 0D, False)
                AssertNumeric(form.NUpD_CountRack, 1D, 10000D, 3D, True)
                AssertNumeric(form.NUpD_OffsetColumnsPile, -100000000000D, 10000000000D, 0D, False)
                AssertNumeric(form.NUpD_OffsetRowsPile, -100000000D, 100000000D, 0D, False)
                AssertNumeric(form.NUpD_CountColumnsPile, 1D, 10000000D, 1D, False)
                AssertNumeric(form.NUpD_CountRowsPile, 1D, 1000000D, 1D, False)

                For Each scale As NumericUpDown In New NumericUpDown() {
                    form.NUpD_ScaleFront, form.NUpD_ScaleLeft, form.NUpD_ScaleRight}
                    Assert.That(scale.Value, [Is].EqualTo(1D), scale.Name)
                    Assert.That(scale.Minimum, [Is].EqualTo(0.01D), scale.Name)
                Next

                Dim listOnly As ComboBox() = {
                    form.CBox_ListNamesArrProject, form.CBox_ListNamesBridge,
                    form.CB_ProjectSurface, form.CB_EgSurface, form.CB_NameAlignment,
                    form.CBox_ListNamesTemplateXML, form.ComboBox11,
                    form.CB_NozzleTLC, form.CB_CabinetWallTLC,
                    form.CB_LeftHandTLC, form.CB_RightHandTLC,
                    form.CB_LeftPostcardTLC, form.CB_RightPostcardTLC,
                    form.CB_RackTLC, form.CB_GrillageTLC, form.CB_PreparationTLC, form.CB_PileTLC}
                For Each combo As ComboBox In listOnly
                    Assert.That(combo.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList), combo.Name)
                Next
                Assert.That(form.CB_NumberPillar.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDown))

                Assert.That(form.ChB_CalculateVerticalLineLeftHand.Checked, [Is].True)
                Assert.That(form.ChB_CalculateVerticalLineRightHand.Checked, [Is].True)
                Assert.That(form.ChB_LeftPostcsrdFixedLenght.Checked, [Is].True)
                Assert.That(form.ChB_RightPostcsrdFixedLenght.Checked, [Is].True)
                Assert.That(form.ChB_EgeParallel.Checked, [Is].True)
                For Each omitOption As CheckBox In New CheckBox() {
                    form.ChB_CreateRack, form.ChB_CreateGrillage, form.ChB_CreatePreparation,
                    form.ChB_InsertPileInRack, form.ChB_PileExpand}
                    Assert.That(omitOption.Checked, [Is].False, omitOption.Name)
                Next
                Assert.That(form.Button8.Enabled, [Is].False)

                AssertGrid(form.DGV_Nozzle, "Column1", "Column2")
                AssertGrid(form.DGV_SubFermenter, "DataGridViewTextBoxColumn19", "DataGridViewTextBoxColumn20")
                AssertGrid(form.DGV_CabinetWall, "DataGridViewTextBoxColumn1", "DataGridViewTextBoxColumn2")
                AssertGrid(form.DGV_LeftHand, "DataGridViewTextBoxColumn3", "DataGridViewTextBoxColumn4")
                AssertGrid(form.DGV_RightHand, "DataGridViewTextBoxColumn5", "DataGridViewTextBoxColumn6")
                AssertGrid(form.DGV_LeftPostcard, "DataGridViewTextBoxColumn15", "DataGridViewTextBoxColumn16")
                AssertGrid(form.DGV_RightPostcard, "DataGridViewTextBoxColumn17", "DataGridViewTextBoxColumn18")
                AssertGrid(form.DGV_Rack, "DataGridViewTextBoxColumn7", "DataGridViewTextBoxColumn8")
                AssertGrid(form.DGV_Grillage, "DataGridViewTextBoxColumn9", "DataGridViewTextBoxColumn10")
                AssertGrid(form.DGV_Preparation, "DataGridViewTextBoxColumn11", "DataGridViewTextBoxColumn12")
                AssertGrid(form.DGV_Piles, "DataGridViewTextBoxColumn13", "DataGridViewTextBoxColumn14")
            End Using
        End Sub

        <Test, Category("Legacy")>
        Public Sub DesignerKeepsNestedSerializationHierarchyTagsAndPreviewIdentity()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Assert.That(form.TabControl1.Parent, [Is].SameAs(form))
                Assert.That(form.TabControl4.Parent, [Is].SameAs(form))
                Assert.That(form.TabControl2.Parent, [Is].SameAs(form.TabPage7))
                Assert.That(form.TabControl3.Parent, [Is].SameAs(form.TabPage6))
                AssertTabOrder(form.TabControl1, "TabPage1", "TabPage16", "TabPage8", "TabPage7",
                               "TabPage6", "TabPage2", "TabPage4", "TabPage3", "TabPage5")
                AssertTabOrder(form.TabControl2, "TabPage9", "TabPage10")
                AssertTabOrder(form.TabControl3, "TabPage11", "TabPage12")
                AssertTabOrder(form.TabControl4, "TabPage13", "TabPage14")

                AssertTags(form)
                Assert.That(form.PictureBox1.Parent, [Is].SameAs(form.TabPage13))
                Assert.That(form.PictureBox2.Parent, [Is].SameAs(form.TabPage14))
                Assert.That(form.PictureBox5.Parent, [Is].SameAs(form.TabPage14))
                For Each editor As Control In LegacyPreviewEditors(form)
                    Assert.That(editor.Parent, [Is].AnyOf(form.TabPage13, form.TabPage14), editor.Name)
                Next
                For Each control As Control In SerializedControls(form)
                    Assert.That(TypeOf control.Parent Is TabPage OrElse control.Parent Is form.GroupBox1,
                                [Is].True, control.Name & " must remain directly under its legacy serialization parent.")
                Next
            End Using
        End Sub

        <Test, Category("Modern")>
        Public Sub ModernContractExposesTheExpectedMethodsAndUsesExplicitNinePageHelp()
            Dim formType As Type = GetType(FormCreateLastPillars)
            Assert.That(FindMethod(formType, "ApplyModernAppearance", Type.EmptyTypes), [Is].Not.Null)
            Assert.That(FindMethod(formType, "RegisterParameterScheme",
                                   New Type() {GetType(TabPage), GetType(String)}), [Is].Not.Null)
            Assert.That(FindMethod(formType, "SetPileLayoutPreview",
                                   New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,))}), [Is].Not.Null)
            Assert.That(FindMethod(formType, "SetPileLayoutPreview",
                                   New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,)),
                                               GetType(Double(,)), GetType(Double())}), [Is].Not.Null)

            Dim expected As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"TabPage1", "\FileResources\ImageObject\Bridge\Pillars\NozzlePillar.PNG"},
                {"TabPage16", "\FileResources\ImageObject\Bridge\Pillars\SubFermenterPillar.PNG"},
                {"TabPage8", "\FileResources\ImageObject\Bridge\Pillars\CabinetWallPillar.PNG"},
                {"TabPage7", "\FileResources\ImageObject\Bridge\Pillars\HandPillar.PNG"},
                {"TabPage6", "\FileResources\ImageObject\Bridge\Pillars\PostcardPillar.PNG"},
                {"TabPage2", "\FileResources\ImageObject\Bridge\Pillars\RackPillar.PNG"},
                {"TabPage4", "\FileResources\ImageObject\Bridge\Pillars\GrillagePillar.PNG"},
                {"TabPage3", "\FileResources\ImageObject\Bridge\Pillars\PreparationPillar.PNG"},
                {"TabPage5", "\FileResources\ImageObject\Bridge\Pillars\PilePillar.PNG"}
            }
            Dim source As String = ReadFormSource()
            Dim matches As MatchCollection = Regex.Matches(
                source,
                "RegisterParameterScheme\s*\(\s*(?<tab>TabPage\d+)\s*,\s*generalDir\s*&\s*""(?<path>[^""]+)""\s*\)",
                RegexOptions.IgnoreCase Or RegexOptions.Multiline)
            Dim actual As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each registration As Match In matches
                actual(registration.Groups("tab").Value) = registration.Groups("path").Value
            Next
            Assert.That(matches.Count, [Is].EqualTo(expected.Count))
            For Each pair As KeyValuePair(Of String, String) In expected
                Assert.That(actual.ContainsKey(pair.Key), [Is].True, pair.Key)
                If actual.ContainsKey(pair.Key) Then Assert.That(actual(pair.Key), [Is].EqualTo(pair.Value).IgnoreCase)
            Next
            Assert.That(source, Does.Not.Contain("New TabPageImageToolTip"),
                        "LastPillars parameter help must open from the explicit button only.")
        End Sub

        <Test, Category("Modern")>
        Public Sub AppearanceAppliedTwicePreservesSeededSerializationReferencesParentsAndTags()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Dim appearance As MethodInfo = RequireMethod("ApplyModernAppearance", Type.EmptyTypes)
                SeedFunctionalState(form)
                Dim controls As Control() = Descendants(form).ToArray()
                Dim parents As Dictionary(Of Control, Control) = controls.ToDictionary(Function(item) item,
                                                                                        Function(item) item.Parent)
                Dim tags As Dictionary(Of Control, Object) = controls.ToDictionary(Function(item) item,
                                                                                     Function(item) item.Tag)
                Dim serializedBefore As String() = CaptureSerializationState(form)
                Dim references As Dictionary(Of String, Control) = controls.Where(Function(item) item.Name.Length > 0).
                    ToDictionary(Function(item) item.Name, Function(item) item, StringComparer.Ordinal)
                Dim images As New Dictionary(Of PictureBox, Image) From {
                    {form.PictureBox1, form.PictureBox1.Image},
                    {form.PictureBox2, form.PictureBox2.Image},
                    {form.PictureBox5, form.PictureBox5.Image}}

                Invoke(appearance, form)
                Invoke(appearance, form)

                Assert.That(CaptureSerializationState(form), [Is].EqualTo(serializedBefore))
                For Each pair As KeyValuePair(Of Control, Control) In parents
                    Assert.That(pair.Key.Parent, [Is].SameAs(pair.Value), pair.Key.Name & " parent")
                    Assert.That(pair.Key.Tag, [Is].EqualTo(tags(pair.Key)), pair.Key.Name & " tag")
                Next
                For Each pair As KeyValuePair(Of String, Control) In references
                    Assert.That(FindControl(form, pair.Key), [Is].SameAs(pair.Value), pair.Key & " reference")
                Next
                For Each pair As KeyValuePair(Of PictureBox, Image) In images
                    Assert.That(pair.Key.Image, [Is].SameAs(pair.Value), pair.Key.Name & " image")
                Next
                AssertTags(form)
                AssertModernCaptions(form)
            End Using
        End Sub

        <Test, Category("Modern")>
        Public Sub AppearanceCanBeAppliedTwiceWithoutFurtherVisualChanges()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Dim appearance As MethodInfo = RequireMethod("ApplyModernAppearance", Type.EmptyTypes)
                Invoke(appearance, form)
                Dim afterFirst As String() = CaptureVisualState(form)
                Invoke(appearance, form)
                Assert.That(CaptureVisualState(form), [Is].EqualTo(afterFirst))
            End Using
        End Sub

        <TestCase(False, Category:="Modern", TestName:="LastPillarsModernLayoutWorksAtPreferredSize")>
        <TestCase(True, Category:="Modern", TestName:="LastPillarsModernLayoutWorksAtMinimumSize")>
        Public Sub ModernLayoutKeepsParametersLeftAndPreviewRightWithoutOverlap(useMinimumSize As Boolean)
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Invoke(RequireMethod("ApplyModernAppearance", Type.EmptyTypes), form)
                form.ShowInTaskbar = False
                form.StartPosition = FormStartPosition.Manual
                form.Location = New Point(-30000, -30000)
                Dim requested As Size = If(useMinimumSize, form.MinimumSize, form.Size)
                Assert.That(requested.Width, [Is].GreaterThan(0))
                Assert.That(requested.Height, [Is].GreaterThan(0))
                form.Size = requested
                form.Show()
                Application.DoEvents()
                PerformLayoutTree(form)

                AssertLeftRightLayout(form)
                Assert.That(form.TabControl1.Multiline, [Is].True)
                For index As Integer = 0 To form.TabControl1.TabCount - 1
                    Assert.That(form.TabControl1.GetTabRect(index).Width, [Is].GreaterThan(0),
                                form.TabControl1.TabPages(index).Name & " tab width")
                Next
                AssertTabOrder(form.TabControl1, "TabPage1", "TabPage16", "TabPage8", "TabPage7",
                               "TabPage6", "TabPage2", "TabPage4", "TabPage3", "TabPage5")
                Assert.That(form.TabControl4.TabCount, [Is].EqualTo(3))
                For Each grid As DataGridView In AllGrids(form)
                    Assert.That(grid.Width, [Is].GreaterThan(100), grid.Name & " width")
                    Assert.That(grid.Height, [Is].GreaterThan(80), grid.Name & " height")
                Next

                form.Size = New Size(Math.Max(form.MinimumSize.Width, requested.Width + 83),
                                     Math.Max(form.MinimumSize.Height, requested.Height + 47))
                Application.DoEvents()
                PerformLayoutTree(form)
                AssertLeftRightLayout(form)
            End Using
        End Sub

        <Test, Category("Modern")>
        Public Sub WheelZoomReplacesButDoesNotRemoveLegacyPreviewEditors()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Dim before As Dictionary(Of Control, String) = LegacyPreviewEditors(form).
                    ToDictionary(Function(item) item,
                                 Function(item) item.Parent.Name & "|" & ControlValue(item))
                Invoke(RequireMethod("ApplyModernAppearance", Type.EmptyTypes), form)

                For Each pair As KeyValuePair(Of Control, String) In before
                    Assert.That(pair.Key.Visible, [Is].False, pair.Key.Name)
                    Assert.That(pair.Key.Parent.Name & "|" & ControlValue(pair.Key), [Is].EqualTo(pair.Value), pair.Key.Name)
                Next
                For Each picture As PictureBox In New PictureBox() {form.PictureBox1, form.PictureBox2, form.PictureBox5}
                    Assert.That(picture.SizeMode, [Is].EqualTo(PictureBoxSizeMode.Zoom), picture.Name)
                Next
                Dim zoomFields As FieldInfo() = GetType(FormCreateLastPillars).GetFields(
                    BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic).
                    Where(Function(field) field.FieldType Is GetType(PreviewImageZoom)).ToArray()
                Assert.That(zoomFields.Length, [Is].EqualTo(1), "Exactly one shared wheel-zoom controller is expected.")
                Assert.That(zoomFields(0).GetValue(form), [Is].Not.Null)
            End Using
        End Sub

        <Test, Category("Modern")>
        Public Sub PileLayoutTabStartsBlankAndBothSetOverloadsRenderWithoutChangingOriginalViews()
            Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                Invoke(RequireMethod("ApplyModernAppearance", Type.EmptyTypes), form)
                Dim layoutTab As TabPage = DirectCast(RequireControl(form, "ModernPileLayoutTab"), TabPage)
                Dim preview As PictureBox = DirectCast(RequireControl(form, "ModernPileLayoutPreview"), PictureBox)
                Assert.That(TabOrder(form.TabControl4).Take(2), [Is].EqualTo(New String() {"TabPage13", "TabPage14"}))
                Assert.That(layoutTab.Text, [Is].EqualTo("Раскладка свай"))
                Assert.That(layoutTab.Parent, [Is].SameAs(form.TabControl4))
                Assert.That(preview.Parent, [Is].SameAs(layoutTab))
                Assert.That(preview.Image, [Is].Null)

                form.TabControl4.SelectedTab = form.TabPage14
                Dim originalParents As Control() = {form.PictureBox1.Parent, form.PictureBox2.Parent, form.PictureBox5.Parent}
                Dim corners As Double(,) = {
                    {0.0R, 0.0R, 0.5R}, {0.0R, 1.5R, 0.5R},
                    {4.1R, 1.5R, 0.5R}, {4.1R, 0.0R, 0.5R}}
                Dim piles As Double(,) = {
                    {0.25R, 0.25R, 1.0R, 0.25R, 0.25R, 12.0R, 0.5R, 0.0R},
                    {0.85R, 0.25R, 1.0R, 0.9R, 0.25R, 12.0R, 0.0R, 0.45R}}
                Dim edges As Double(,) = {{0.0R, 0.75R, 0.5R, 4.1R, 0.75R, 0.5R}}
                Invoke(RequireMethod("SetPileLayoutPreview",
                                     New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,))}),
                       form, corners, piles, edges)
                Assert.That(preview.Image, [Is].Not.Null)

                Dim axisData As Double(,) = {
                    {1.0R, 1.0R, 0.25R, 0.25R},
                    {1.0R, 2.0R, 0.85R, 0.25R}}
                Invoke(RequireMethod("SetPileLayoutPreview",
                                     New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,)),
                                                 GetType(Double(,)), GetType(Double())}),
                       form, corners, piles, edges, axisData, New Double() {4.1R, 1.5R})
                Assert.That(preview.Image, [Is].Not.Null)
                Assert.That(form.TabControl4.SelectedTab, [Is].SameAs(form.TabPage14))
                Assert.That(form.PictureBox1.Parent, [Is].SameAs(originalParents(0)))
                Assert.That(form.PictureBox2.Parent, [Is].SameAs(originalParents(1)))
                Assert.That(form.PictureBox5.Parent, [Is].SameAs(originalParents(2)))

                Invoke(RequireMethod("SetPileLayoutPreview",
                                     New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,))}),
                       form, Nothing, Nothing, Nothing)
                Assert.That(preview.Image, [Is].Null)
            End Using
        End Sub

        <Test, Category("Modern")>
        Public Sub ExplicitSchemeHelpStaysInsideTheLeftParameterAreaAndClosesOnTabChange()
            Dim firstPath As String = Path.Combine(TestContext.CurrentContext.WorkDirectory,
                                                   Guid.NewGuid().ToString("N") & "-last-first.png")
            Dim secondPath As String = Path.Combine(TestContext.CurrentContext.WorkDirectory,
                                                    Guid.NewGuid().ToString("N") & "-last-second.png")
            SaveImage(firstPath, New Size(31, 17), Color.Red)
            SaveImage(secondPath, New Size(43, 19), Color.Blue)
            Try
                Using form As FormCreateLastPillars = FormCreateLastPillars.CreateDesignerOnly()
                    Invoke(RequireMethod("ApplyModernAppearance", Type.EmptyTypes), form)
                    Dim register As MethodInfo = RequireMethod(
                        "RegisterParameterScheme", New Type() {GetType(TabPage), GetType(String)})
                    Invoke(register, form, form.TabPage1, firstPath)
                    Invoke(register, form, form.TabPage16, secondPath)
                    form.ShowInTaskbar = False
                    form.StartPosition = FormStartPosition.Manual
                    form.Location = New Point(-30000, -30000)
                    form.Show()
                    Application.DoEvents()

                    Dim button As Button = DirectCast(RequireControl(form, "ModernParameterSchemeButton"), Button)
                    Dim panel As Panel = DirectCast(RequireControl(form, "ModernParameterSchemePanel"), Panel)
                    Dim picture As PictureBox = DirectCast(RequireControl(form, "ModernParameterSchemeImage"), PictureBox)
                    form.TabControl1.SelectedTab = form.TabPage1
                    button.PerformClick()
                    Application.DoEvents()
                    Assert.That(panel.Visible, [Is].True)
                    Assert.That(picture.Image, [Is].Not.Null)
                    Assert.That(picture.Image.Size, [Is].EqualTo(New Size(31, 17)))
                    Assert.That(ScreenBounds(form.TabControl1).Contains(ScreenBounds(panel)), [Is].True)
                    Assert.That(ScreenBounds(panel).IntersectsWith(ScreenBounds(form.TabControl4)), [Is].False)

                    form.TabControl1.SelectedTab = form.TabPage16
                    Application.DoEvents()
                    Assert.That(panel.Visible, [Is].False)
                    Assert.That(picture.Image, [Is].Null)
                    button.PerformClick()
                    Application.DoEvents()
                    Assert.That(picture.Image.Size, [Is].EqualTo(New Size(43, 19)))
                End Using
            Finally
                If File.Exists(firstPath) Then File.Delete(firstPath)
                If File.Exists(secondPath) Then File.Delete(secondPath)
            End Try
        End Sub

        Private Shared Sub AssertNumeric(control As NumericUpDown, minimum As Decimal,
                                         maximum As Decimal, value As Decimal, expectedReadOnly As Boolean)
            Assert.That(control.Minimum, [Is].EqualTo(minimum), control.Name & " minimum")
            Assert.That(control.Maximum, [Is].EqualTo(maximum), control.Name & " maximum")
            Assert.That(control.Value, [Is].EqualTo(value), control.Name & " value")
            Assert.That(control.ReadOnly, [Is].EqualTo(expectedReadOnly), control.Name & " read-only")
        End Sub

        Private Shared Sub AssertGrid(grid As DataGridView, ParamArray names As String())
            Assert.That(grid.Columns.Count, [Is].EqualTo(names.Length), grid.Name & " column count")
            For index As Integer = 0 To names.Length - 1
                Assert.That(grid.Columns(index).Name, [Is].EqualTo(names(index)), grid.Name & " column " & index)
                Assert.That(grid.Columns(index), [Is].TypeOf(Of DataGridViewTextBoxColumn)())
            Next
        End Sub

        Private Shared Sub AssertTags(form As FormCreateLastPillars)
            Assert.That(form.TabControl1.Tag, [Is].EqualTo("SlopingWings"))
            Assert.That(form.TabControl2.Tag, [Is].EqualTo("SlopingWingsRight"))
            Dim expected As New Dictionary(Of TabPage, String) From {
                {form.TabPage1, "Nozzle"}, {form.TabPage8, "CabinetWall"},
                {form.TabPage7, "SlopingWings"}, {form.TabPage9, "SlopingWingsLeft"},
                {form.TabPage10, "SlopingWingsRight"}, {form.TabPage6, "Postcards"},
                {form.TabPage11, "PostcardsLeft"}, {form.TabPage12, "PostcardsRight"},
                {form.TabPage2, "Racks"}, {form.TabPage4, "Grillage"},
                {form.TabPage3, "Preparation"}, {form.TabPage5, "Piles"}}
            For Each pair As KeyValuePair(Of TabPage, String) In expected
                Assert.That(pair.Key.Tag, [Is].EqualTo(pair.Value), pair.Key.Name)
            Next
            Assert.That(form.TabPage16.Tag, [Is].Null)
            Assert.That(form.TabPage13.Tag, [Is].Null)
            Assert.That(form.TabPage14.Tag, [Is].Null)
        End Sub

        Private Shared Sub AssertModernCaptions(form As FormCreateLastPillars)
            Dim expected As String() = {
                "Насадка", "Подферменник", "Шкафная стенка", "Откосные крылья", "Открылки",
                "Стойки", "Ростверк", "Подготовка", "Сваи"}
            Assert.That(form.TabControl1.TabPages.Cast(Of TabPage)().Select(Function(page) page.Text),
                        [Is].EqualTo(expected))
            For Each leftPage As TabPage In New TabPage() {form.TabPage9, form.TabPage11}
                Assert.That(leftPage.Text, Does.StartWith("Лев"), leftPage.Name)
            Next
            For Each rightPage As TabPage In New TabPage() {form.TabPage10, form.TabPage12}
                Assert.That(rightPage.Text, Does.StartWith("Прав"), rightPage.Name)
            Next
        End Sub

        Private Shared Sub AssertTabOrder(tabs As TabControl, ParamArray expected As String())
            Assert.That(TabOrder(tabs), [Is].EqualTo(expected), tabs.Name)
        End Sub

        Private Shared Function TabOrder(tabs As TabControl) As String()
            Return tabs.TabPages.Cast(Of TabPage)().Select(Function(page) page.Name).ToArray()
        End Function

        Private Shared Function SerializedControls(form As FormCreateLastPillars) As Control()
            Return Descendants(form).Where(
                Function(control) control.Name.Length > 0 AndAlso
                                  (TypeOf control Is ComboBox OrElse TypeOf control Is NumericUpDown OrElse
                                   TypeOf control Is TextBox OrElse TypeOf control Is CheckBox OrElse
                                   TypeOf control Is DataGridView)).ToArray()
        End Function

        Private Shared Function LegacyPreviewEditors(form As FormCreateLastPillars) As Control()
            Return New Control() {
                form.Label30, form.Label31, form.NUpD_ScaleFront, form.Label32,
                form.NUpD_dxFront, form.NUpD_dyFront,
                form.Label36, form.Label37, form.NUpD_ScaleRight, form.Label38,
                form.NUpD_dxRight, form.NUpD_dyRight,
                form.Label33, form.Label34, form.NUpD_ScaleLeft, form.Label35,
                form.NUpD_dxLeft, form.NUpD_dyLeft}
        End Function

        Private Shared Function AllGrids(form As FormCreateLastPillars) As DataGridView()
            Return New DataGridView() {
                form.DGV_Nozzle, form.DGV_SubFermenter, form.DGV_CabinetWall,
                form.DGV_LeftHand, form.DGV_RightHand, form.DGV_LeftPostcard, form.DGV_RightPostcard,
                form.DGV_Rack, form.DGV_Grillage, form.DGV_Preparation, form.DGV_Piles}
        End Function

        Private Shared Sub SeedFunctionalState(form As FormCreateLastPillars)
            form.NUpD_ElevationLand.Value = -12.34D
            form.NUpD_CountRack.Value = 7D
            form.NUpD_OffsetColumnsPile.Value = -0.75D
            form.NUpD_OffsetRowsPile.Value = 1.25D
            form.NUpD_CountColumnsPile.Value = 4D
            form.NUpD_CountRowsPile.Value = 3D
            form.NUpD_ScaleFront.Value = 2.5D
            form.NUpD_dxLeft.Value = -17D
            form.NUpD_dyRight.Value = 23D
            form.TxtB_PileCollDiagram.Text = "1 2 1"
            form.TxtB_PileRowDiagram.Text = "2 1"
            form.ChB_ElevationLand.Checked = True
            form.ChB_CreateRack.Checked = True
            form.ChB_CreateGrillage.Checked = True
            form.ChB_CreatePreparation.Checked = True
            form.ChB_InsertPileInRack.Checked = True
            form.ChB_PileExpand.Checked = True
            form.ChB_CalculateVerticalLineLeftHand.Checked = False
            form.ChB_RightPostcsrdFixedLenght.Checked = False

            Dim combos As ComboBox() = {
                form.CBox_ListNamesArrProject, form.CBox_ListNamesBridge, form.CB_ProjectSurface,
                form.CB_EgSurface, form.CB_NameAlignment, form.CBox_ListNamesTemplateXML, form.ComboBox11,
                form.CB_NozzleTLC, form.CB_CabinetWallTLC, form.CB_LeftHandTLC, form.CB_RightHandTLC,
                form.CB_LeftPostcardTLC, form.CB_RightPostcardTLC, form.CB_RackTLC,
                form.CB_GrillageTLC, form.CB_PreparationTLC, form.CB_PileTLC}
            For index As Integer = 0 To combos.Length - 1
                combos(index).Items.Add("seed-" & index.ToString(CultureInfo.InvariantCulture))
                combos(index).SelectedIndex = 0
            Next
            form.CB_NumberPillar.Text = "42"

            For index As Integer = 0 To AllGrids(form).Length - 1
                Dim grid As DataGridView = AllGrids(form)(index)
                grid.Rows.Add("left-" & index.ToString(CultureInfo.InvariantCulture),
                              "right-" & index.ToString(CultureInfo.InvariantCulture))
            Next
        End Sub

        Private Shared Function CaptureSerializationState(form As FormCreateLastPillars) As String()
            Return SerializedControls(form).
                OrderBy(Function(control) control.Name, StringComparer.Ordinal).
                Select(Function(control) control.Name & "|" & control.Parent.Name & "|" &
                                         ValueText(control.Tag) & "|" & ControlValue(control)).ToArray()
        End Function

        Private Shared Function ControlValue(control As Control) As String
            If TypeOf control Is NumericUpDown Then
                Dim numeric As NumericUpDown = DirectCast(control, NumericUpDown)
                Return DecimalText(numeric.Value) & ";" & DecimalText(numeric.Minimum) & ";" &
                       DecimalText(numeric.Maximum) & ";" & numeric.ReadOnly.ToString()
            End If
            If TypeOf control Is CheckBox Then Return DirectCast(control, CheckBox).Checked.ToString()
            If TypeOf control Is ComboBox Then
                Dim combo As ComboBox = DirectCast(control, ComboBox)
                Return combo.SelectedIndex.ToString(CultureInfo.InvariantCulture) & ";" & combo.Text
            End If
            If TypeOf control Is DataGridView Then
                Dim grid As DataGridView = DirectCast(control, DataGridView)
                Dim rows = grid.Rows.Cast(Of DataGridViewRow)().Where(Function(row) Not row.IsNewRow).
                    Select(Function(row) String.Join(",", row.Cells.Cast(Of DataGridViewCell)().
                        Select(Function(cell) ValueText(cell.Value))))
                Return String.Join(";", rows)
            End If
            Return control.Text
        End Function

        Private Shared Function CaptureVisualState(form As FormCreateLastPillars) As String()
            Return Descendants(form).OrderBy(Function(control) control.Name, StringComparer.Ordinal).
                Select(Function(control) control.Name & "|" & control.GetType().Name & "|" &
                                         If(control.Parent Is Nothing, "", control.Parent.Name) & "|" &
                                         control.Bounds.ToString() & "|" & control.Visible.ToString() & "|" &
                                         control.Dock.ToString() & "|" & control.Anchor.ToString() & "|" & control.Text).
                ToArray()
        End Function

        Private Shared Sub AssertLeftRightLayout(form As FormCreateLastPillars)
            Dim left As Rectangle = ScreenBounds(form.TabControl1)
            Dim right As Rectangle = ScreenBounds(form.TabControl4)
            Assert.That(left.Width, [Is].GreaterThan(350))
            Assert.That(right.Width, [Is].GreaterThan(350))
            Assert.That(left.Right, [Is].LessThanOrEqualTo(right.Left))
            Assert.That(left.IntersectsWith(right), [Is].False)
            Assert.That(form.TabControl1.Left, [Is].LessThan(form.TabControl4.Left))
        End Sub

        Private Shared Function ScreenBounds(control As Control) As Rectangle
            Return New Rectangle(control.PointToScreen(Point.Empty), control.Size)
        End Function

        Private Shared Sub SaveImage(path As String, size As Size, color As Color)
            Using bitmap As New Bitmap(size.Width, size.Height), graphics As Graphics = Graphics.FromImage(bitmap)
                graphics.Clear(color)
                bitmap.Save(path)
            End Using
        End Sub

        Private Shared Function FindMethod(formType As Type, name As String, parameterTypes As Type()) As MethodInfo
            Return formType.GetMethod(name,
                                      BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                      Nothing, parameterTypes, Nothing)
        End Function

        Private Shared Function RequireMethod(name As String, parameterTypes As Type()) As MethodInfo
            Dim method As MethodInfo = FindMethod(GetType(FormCreateLastPillars), name, parameterTypes)
            Assert.That(method, [Is].Not.Null, name & "(" & parameterTypes.Length & " parameters) is missing.")
            Return method
        End Function

        Private Shared Sub Invoke(method As MethodInfo, target As Object, ParamArray arguments As Object())
            Try
                method.Invoke(target, arguments)
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Function Descendants(parent As Control) As IEnumerable(Of Control)
            Dim result As New List(Of Control)()
            For Each child As Control In parent.Controls
                result.Add(child)
                result.AddRange(Descendants(child))
            Next
            Return result
        End Function

        Private Shared Function FindControl(parent As Control, name As String) As Control
            Return Descendants(parent).FirstOrDefault(Function(control) control.Name = name)
        End Function

        Private Shared Function RequireControl(parent As Control, name As String) As Control
            Dim control As Control = FindControl(parent, name)
            Assert.That(control, [Is].Not.Null, name & " is missing.")
            Return control
        End Function

        Private Shared Sub PerformLayoutTree(parent As Control)
            parent.PerformLayout()
            For Each child As Control In parent.Controls
                PerformLayoutTree(child)
            Next
        End Sub

        Private Shared Function DecimalText(value As Decimal) As String
            Return value.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function ValueText(value As Object) As String
            If value Is Nothing Then Return "<null>"
            Return Convert.ToString(value, CultureInfo.InvariantCulture)
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
