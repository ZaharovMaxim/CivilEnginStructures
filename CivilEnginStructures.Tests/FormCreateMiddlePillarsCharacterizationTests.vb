Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class FormCreateMiddlePillarsCharacterizationTests
        <Test>
        Public Sub DesignerInitializesFunctionalDefaults()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                Assert.Multiple(
                    Sub()
                        AssertNumeric(form.NUpD_ElevationLand, -99999999999D, 9999999999D, 0D, False, True)
                        AssertNumeric(form.NUpD_CountRack, 1D, 10000D, 3D, True, True)
                        AssertNumeric(form.NUpD_OffsetColumnsPile, -100000000000D, 10000000000D, 0D, False, True)
                        AssertNumeric(form.NUpD_OffsetRowsPile, -100000000D, 100000000D, 0D, False, True)
                        AssertNumeric(form.NUpD_CountColumnsPile, 1D, 10000000D, 1D, True, True)
                        AssertNumeric(form.NUpD_CountRowsPile, 1D, 1000000D, 1D, True, True)

                        AssertPreviewNumeric(form.NUpD_ScaleFront, 0.01D, 100000000D, 1D)
                        AssertPreviewNumeric(form.NUpD_dxFront, -100000000D, 100000000D, 0D)
                        AssertPreviewNumeric(form.NUpD_dyFront, -100000000D, 1000000000D, 0D)
                        AssertPreviewNumeric(form.NUpD_ScaleLeft, 0.01D, 100000000D, 1D)
                        AssertPreviewNumeric(form.NUpD_dxLeft, -100000000D, 100000000D, 0D)
                        AssertPreviewNumeric(form.NUpD_dyLeft, -100000000D, 1000000000D, 0D)
                        AssertPreviewNumeric(form.NUpD_ScaleRight, 0.01D, 100000000D, 1D)
                        AssertPreviewNumeric(form.NUpD_dxRight, -100000000D, 100000000D, 0D)
                        AssertPreviewNumeric(form.NUpD_dyRight, -100000000D, 1000000000D, 0D)
                    End Sub)

                Dim listOnly As ComboBox() = {
                    form.CBox_ListNamesArrProject, form.CBox_ListNamesBridge,
                    form.CB_ProjectSurface, form.CB_EgSurface, form.CB_NameAlignment,
                    form.CBox_ListNamesTemplateXML, form.ComboBox11,
                    form.CB_RigelTLC, form.CB_RackTLC, form.CB_GrillageTLC,
                    form.CB_PreparationTLC, form.CB_PileTLC
                }
                For Each combo As ComboBox In listOnly
                    Assert.That(combo.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList), combo.Name)
                Next
                Assert.That(form.CB_NumberPillar.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDown))

                For Each checkBox As CheckBox In AllCheckBoxes(form)
                    Assert.Multiple(
                        Sub()
                            Assert.That(checkBox.Checked, [Is].False, checkBox.Name & " checked")
                            Assert.That(checkBox.Enabled, [Is].True, checkBox.Name & " enabled")
                        End Sub)
                Next

                AssertGrid(form.DGV_Rigel, "Column1", "Column2")
                AssertGrid(form.DGV_SubFermenter, "DataGridViewTextBoxColumn1", "DataGridViewTextBoxColumn2")
                AssertGrid(form.DGV_Rack, "DataGridViewTextBoxColumn7", "DataGridViewTextBoxColumn8")
                AssertGrid(form.DGV_Grillage, "DataGridViewTextBoxColumn9", "DataGridViewTextBoxColumn10")
                AssertGrid(form.DGV_Preparation, "DataGridViewTextBoxColumn11", "DataGridViewTextBoxColumn12")
                AssertGrid(form.DGV_Piles, "DataGridViewTextBoxColumn13", "DataGridViewTextBoxColumn14")
            End Using
        End Sub

        <Test>
        Public Sub DesignerKeepsSerializationControlsInTheLegacyDirectHierarchy()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                Assert.That(form.TabControl1.Parent, [Is].SameAs(form))
                Assert.That(form.TabControl2.Parent, [Is].SameAs(form))
                AssertTabOrder(form.TabControl1, "TabPage1", "TabPage9", "TabPage2", "TabPage4", "TabPage3", "TabPage5")
                AssertTabOrder(form.TabControl2, "TabPage6", "TabPage7")

                For Each tabs As TabControl In New TabControl() {form.TabControl1, form.TabControl2}
                    For Each page As TabPage In tabs.TabPages
                        For Each control As Control In page.Controls
                            If IsSerializedControl(control) Then
                                Assert.That(control.Parent, [Is].SameAs(page),
                                            control.Name & " must remain a direct child of " & page.Name & ".")
                            End If
                        Next
                    Next
                Next
            End Using
        End Sub

        <Test>
        Public Sub AppearancePreservesDesignerIdentitySeededStateAndSerializationTraversal()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                Dim appearance As MethodInfo = RequireAppearanceMethod()
                SeedFunctionalState(form)

                Dim references As Dictionary(Of String, Object) = CaptureDesignerReferences(form)
                Dim parents As Dictionary(Of Control, TabPage) = CaptureSerializedParents(form)
                Dim images As Dictionary(Of String, Image) = CapturePictureImages(form)
                Dim before As String() = CaptureSerializationTraversal(form)
                Dim componentOrder As String() = TabOrder(form.TabControl1)
                Dim previewOrder As String() = TabOrder(form.TabControl2)

                InvokeAppearance(appearance, form)

                Assert.That(CaptureSerializationTraversal(form), [Is].EqualTo(before),
                            "Appearance must preserve the values consumed by the legacy XML traversal.")
                Assert.That(TabOrder(form.TabControl1), [Is].EqualTo(componentOrder))
                Assert.That(TabOrder(form.TabControl2).Take(previewOrder.Length), [Is].EqualTo(previewOrder),
                            "Appearance must preserve the two designer preview pages and their order.")
                Assert.That(TabOrder(form.TabControl2).Last(), [Is].EqualTo("ModernPileLayoutTab"))
                Assert.That(form.TabControl1.Parent, [Is].SameAs(form))
                Assert.That(form.TabControl2.Parent, [Is].SameAs(form))
                For Each pair As KeyValuePair(Of Control, TabPage) In parents
                    Assert.That(pair.Key.Parent, [Is].SameAs(pair.Value), pair.Key.Name & " parent")
                Next
                AssertDesignerReferences(form, references)
                For Each picture As KeyValuePair(Of String, Image) In images
                    Assert.That(FindControl(form, picture.Key), [Is].TypeOf(Of PictureBox)())
                    Assert.That(DirectCast(FindControl(form, picture.Key), PictureBox).Image,
                                [Is].SameAs(picture.Value), picture.Key & " image")
                Next
            End Using
        End Sub

        <Test>
        Public Sub AppearanceCanBeAppliedTwiceWithoutFurtherVisualChanges()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                Dim appearance As MethodInfo = RequireAppearanceMethod()
                InvokeAppearance(appearance, form)
                Dim afterFirst As String() = CaptureVisualState(form)
                InvokeAppearance(appearance, form)
                Assert.That(CaptureVisualState(form), [Is].EqualTo(afterFirst))
            End Using
        End Sub

        <TestCase(False, TestName:="AppearanceLayoutWorksAtPreferredSize")>
        <TestCase(True, TestName:="AppearanceLayoutWorksAtMinimumSize")>
        Public Sub AppearanceLayoutKeepsAllTabsFieldsAndFooterReachable(useMinimumSize As Boolean)
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                InvokeAppearance(RequireAppearanceMethod(), form)
                form.ShowInTaskbar = False
                form.StartPosition = FormStartPosition.Manual
                form.Location = New Point(-30000, -30000)
                Dim targetSize As Size = If(useMinimumSize, form.MinimumSize, form.Size)
                Assert.That(targetSize.Width, [Is].GreaterThan(0))
                Assert.That(targetSize.Height, [Is].GreaterThan(0))

                form.Size = targetSize
                form.Show()
                Application.DoEvents()
                PerformLayoutTree(form)

                AssertFirstShownLayout(form)
                AssertCompactProjectGroupLayout(form)
                AssertPreviewLayout(form)

                form.Size = New Size(Math.Max(targetSize.Width + 73, form.MinimumSize.Width),
                                     Math.Max(targetSize.Height + 41, form.MinimumSize.Height))
                PerformLayoutTree(form)
                form.Size = targetSize
                PerformLayoutTree(form)
                form.Size = New Size(Math.Max(targetSize.Width + 29, form.MinimumSize.Width),
                                     Math.Max(targetSize.Height + 19, form.MinimumSize.Height))
                form.Size = targetSize
                PerformLayoutTree(form)

                Assert.That(form.TabControl1.TabCount, [Is].EqualTo(6))
                Assert.That(form.TabControl2.TabCount, [Is].EqualTo(3))
                ValidateTabs(form.TabControl1)
                ValidateTabs(form.TabControl2)
                AssertPilePageLayout(form)
                For Each action As Button In New Button() {form.Button2, form.Button7, form.Button8}
                    AssertFullyInsideAncestor(action, action.Parent)
                    AssertFullyInsideAncestor(action, form)
                Next
            End Using
        End Sub

        <Test>
        Public Sub AppearanceSmallViewportUsesRootScrollingToKeepFooterAccessible()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                InvokeAppearance(RequireAppearanceMethod(), form)
                form.ShowInTaskbar = False
                form.StartPosition = FormStartPosition.Manual
                form.Location = New Point(-30000, -30000)
                form.MinimumSize = Size.Empty
                form.Size = New Size(900, 520)

                form.Show()
                Application.DoEvents()
                PerformLayoutTree(form)

                Assert.That(form.TabControl1.Parent, [Is].SameAs(form))
                Assert.That(form.TabControl2.Parent, [Is].SameAs(form))
                Assert.That(form.AutoScroll, [Is].True,
                            "A viewport below the logical canvas must expose root scrolling.")
                Dim scaledCanvas As New Size(ScaleForDpi(form, 760), ScaleForDpi(form, 600))
                Assert.That(form.AutoScrollMinSize.Width, [Is].GreaterThanOrEqualTo(scaledCanvas.Width))
                Assert.That(form.AutoScrollMinSize.Height, [Is].GreaterThanOrEqualTo(scaledCanvas.Height))

                Dim footer As Control = FindControl(form, "ModernFooter")
                Assert.That(footer, [Is].Not.Null)
                form.ScrollControlIntoView(footer)
                Application.DoEvents()
                Assert.That(-form.AutoScrollPosition.Y, [Is].GreaterThan(0),
                            "Showing the footer in the small viewport must move the root scroll position.")
                For Each action As Button In New Button() {form.Button2, form.Button7, form.Button8}
                    Assert.That(action.Parent, [Is].SameAs(footer), action.Name & " footer parent")
                    AssertFullyInsideAncestor(action, footer)
                    AssertFullyInsideAncestor(action, form)
                Next
            End Using
        End Sub

        <Test>
        Public Sub PileScrollPositionSurvivesWidthResizeAndClampsAfterHeightResize()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                InvokeAppearance(RequireAppearanceMethod(), form)
                form.ShowInTaskbar = False
                form.StartPosition = FormStartPosition.Manual
                form.Location = New Point(-30000, -30000)
                Dim initialSize As Size = form.MinimumSize
                form.Size = initialSize
                form.TabControl1.SelectedTab = form.TabPage5

                form.Show()
                Application.DoEvents()
                PerformLayoutTree(form)
                form.TabPage5.ScrollControlIntoView(form.DGV_Piles)
                Application.DoEvents()

                Dim initialOffset As Integer = -form.TabPage5.AutoScrollPosition.Y
                Assert.That(initialOffset, [Is].GreaterThan(0),
                            "The pile grid must require a positive vertical scroll offset at minimum height.")
                AssertFullyInsideAncestor(form.DGV_Piles, form.TabPage5)
                Dim stateBeforeResize As String() = CaptureSerializationTraversal(form)

                form.Size = New Size(initialSize.Width + 73, initialSize.Height)
                Application.DoEvents()
                Dim widthResizeOffset As Integer = -form.TabPage5.AutoScrollPosition.Y
                Assert.That(widthResizeOffset, [Is].EqualTo(initialOffset),
                            "A width-only resize must preserve the pile page vertical scroll offset.")
                AssertFullyInsideAncestor(form.DGV_Piles, form.TabPage5)
                Assert.That(CaptureSerializationTraversal(form), [Is].EqualTo(stateBeforeResize),
                            "A width-only resize must not reset functional state.")

                form.Size = New Size(initialSize.Width + 73, initialSize.Height + 41)
                Application.DoEvents()
                Dim heightResizeOffset As Integer = -form.TabPage5.AutoScrollPosition.Y
                Dim maximumOffset As Integer = MaximumVerticalScrollOffset(form.TabPage5)
                Assert.That(heightResizeOffset, [Is].EqualTo(Math.Min(widthResizeOffset, maximumOffset)),
                            "A height resize must retain the offset or clamp it to the new scroll range.")
                Assert.That(CaptureSerializationTraversal(form), [Is].EqualTo(stateBeforeResize),
                            "Resizing must not reset functional state.")
            End Using
        End Sub

        <Test>
        Public Sub PileLayoutUsesItsOwnTabWithoutChangingOrSelectingTheOriginalFrontPreview()
            Using form As FormCreateMiddlePillars = FormCreateMiddlePillars.CreateDesignerOnly()
                InvokeAppearance(RequireAppearanceMethod(), form)
                form.ShowInTaskbar = False
                form.StartPosition = FormStartPosition.Manual
                form.Location = New Point(-30000, -30000)
                form.Show()
                Application.DoEvents()

                If form.DGV_Rigel.ColumnCount > 0 AndAlso form.DGV_Rigel.RowCount > 0 Then
                    form.DGV_Rigel.CurrentCell = form.DGV_Rigel.Rows(0).Cells(0)
                End If
                Dim serializedBefore As String() = CaptureSerializationTraversal(form)
                Dim layoutTab As TabPage = TryCast(FindControl(form, "ModernPileLayoutTab"), TabPage)
                Dim layoutPreview As PictureBox = TryCast(FindControl(form, "ModernPileLayoutPreview"), PictureBox)
                Assert.That(layoutTab, [Is].Not.Null)
                Assert.That(layoutPreview, [Is].Not.Null)
                Assert.That(layoutTab.Parent, [Is].SameAs(form.TabControl2))
                Assert.That(layoutTab.Text, [Is].EqualTo("Раскладка свай"))
                Assert.That(layoutPreview.Parent, [Is].SameAs(layoutTab))
                Assert.That(layoutPreview.Image, [Is].Null,
                            "Before a valid calculation the pile layout canvas must stay blank.")
                Assert.That(form.TabControl1.TabCount, [Is].EqualTo(6))
                Assert.That(form.TabControl2.TabCount, [Is].EqualTo(3))

                form.TabControl2.SelectedTab = form.TabPage7
                Dim frontBoundsBefore As Rectangle = form.PictureBox1.Bounds
                form.TabControl1.SelectedTab = form.TabPage5
                PerformLayoutTree(form)

                Assert.Multiple(
                    Sub()
                        Assert.That(form.TabControl2.SelectedTab, [Is].SameAs(form.TabPage7),
                                    "Opening pile parameters must not change the user's preview selection.")
                        Assert.That(form.PictureBox1.Parent, [Is].SameAs(form.TabPage6))
                        Assert.That(form.PictureBox1.Bounds, [Is].EqualTo(frontBoundsBefore),
                                    "The original front preview must retain its full-page layout.")
                    End Sub)

                InvokeSetPileLayoutPreview(form,
                                           New Double(,) {{0.0R, 0.0R, 0.5R}, {0.0R, 1.5R, 0.5R},
                                                           {4.1R, 1.5R, 0.5R}, {4.1R, 0.0R, 0.5R}},
                                           New Double(,) {{0.25R, 0.25R, 1.0R, 0.25R, 0.25R, 6.0R, 0.5R, 0.0R},
                                                           {0.85R, 0.25R, 1.0R, 0.9R, 0.25R, 6.0R, 0.0R, 0.45R}},
                                           New Double(,) {{0.0R, 0.75R, 0.5R, 4.1R, 0.75R, 0.5R}})
                Assert.That(layoutPreview.Image, [Is].Not.Null)

                form.TabControl2.SelectedTab = layoutTab
                form.Size = New Size(form.Width + 17, form.Height + 11)
                PerformLayoutTree(form)
                Assert.That(form.TabControl2.SelectedTab, [Is].SameAs(layoutTab))
                AssertFullyInsideAncestor(layoutPreview, layoutTab)

                form.TabControl1.SelectedTab = form.TabPage1
                form.TabControl2.SelectedTab = form.TabPage6
                PerformLayoutTree(form)
                Assert.Multiple(
                    Sub()
                        AssertFullyInsideAncestor(form.PictureBox1, form.TabPage6)
                        Assert.That(form.PictureBox1.Bounds,
                                    [Is].EqualTo(New Rectangle(ScaleForDpi(form, 10), ScaleForDpi(form, 10),
                                                               Math.Max(ScaleForDpi(form, 100), form.TabPage6.ClientSize.Width - ScaleForDpi(form, 20)),
                                                               Math.Max(ScaleForDpi(form, 100), form.TabPage6.ClientSize.Height - ScaleForDpi(form, 20)))),
                                    "The original front preview must always fill its page.")
                        Assert.That(CaptureSerializationTraversal(form), [Is].EqualTo(serializedBefore),
                                    "The dynamic layout tab must not add serialized inputs or change legacy state.")
                    End Sub)

                InvokeSetPileLayoutPreview(form, Nothing, Nothing, Nothing)
                Assert.That(layoutPreview.Image, [Is].Null)
            End Using
        End Sub

        Private Shared Sub AssertNumeric(control As NumericUpDown, minimum As Decimal, maximum As Decimal,
                                         value As Decimal, expectedReadOnly As Boolean, expectedEnabled As Boolean)
            Assert.That(control.Minimum, [Is].EqualTo(minimum), control.Name & " minimum")
            Assert.That(control.Maximum, [Is].EqualTo(maximum), control.Name & " maximum")
            Assert.That(control.Value, [Is].EqualTo(value), control.Name & " value")
            Assert.That(control.ReadOnly, [Is].EqualTo(expectedReadOnly), control.Name & " read-only")
            Assert.That(control.Enabled, [Is].EqualTo(expectedEnabled), control.Name & " enabled")
        End Sub

        Private Shared Sub AssertPreviewNumeric(control As NumericUpDown, minimum As Decimal,
                                                maximum As Decimal, value As Decimal)
            AssertNumeric(control, minimum, maximum, value, False, True)
        End Sub

        Private Shared Sub AssertGrid(grid As DataGridView, ParamArray names As String())
            Assert.That(grid.ReadOnly, [Is].False, grid.Name & " read-only")
            Assert.That(grid.Columns.Count, [Is].EqualTo(names.Length), grid.Name & " column count")
            For index As Integer = 0 To names.Length - 1
                Dim currentIndex As Integer = index
                Dim column As DataGridViewColumn = grid.Columns(currentIndex)
                Assert.Multiple(
                    Sub()
                        Assert.That(column.Name, [Is].EqualTo(names(currentIndex)), grid.Name & " column name")
                        Assert.That(column, [Is].TypeOf(Of DataGridViewTextBoxColumn)(), grid.Name & " column type")
                        Assert.That(column.ReadOnly, [Is].False, grid.Name & " column read-only")
                        Assert.That(column.Tag, [Is].Null, grid.Name & " column tag")
                    End Sub)
            Next
        End Sub

        Private Shared Function AllCheckBoxes(form As FormCreateMiddlePillars) As CheckBox()
            Return New CheckBox() {form.ChB_ElevationLand, form.ChB_ProjectSurfaceInAlignment,
                                  form.CBox_SingleSubFermentes, form.ChB_fixedHeightRack,
                                  form.ChB_EgeParallel, form.ChB_InsertPileInRack, form.ChB_PileExpand}
        End Function

        Private Shared Sub AssertTabOrder(tabs As TabControl, ParamArray expected As String())
            Assert.That(TabOrder(tabs), [Is].EqualTo(expected), tabs.Name & " tab order")
        End Sub

        Private Shared Function TabOrder(tabs As TabControl) As String()
            Return tabs.TabPages.Cast(Of TabPage)().Select(Function(page) page.Name).ToArray()
        End Function

        Private Shared Function IsSerializedControl(control As Control) As Boolean
            Return TypeOf control Is DataGridView OrElse TypeOf control Is ComboBox OrElse
                   TypeOf control Is CheckBox OrElse TypeOf control Is NumericUpDown OrElse
                   TypeOf control Is TextBox OrElse TypeOf control Is TabControl
        End Function

        Private Shared Function RequireAppearanceMethod() As MethodInfo
            Dim method As MethodInfo = GetType(FormCreateMiddlePillars).GetMethod(
                "ApplyModernAppearance", BindingFlags.Instance Or BindingFlags.NonPublic)
            If method Is Nothing Then
                Assert.Ignore("ApplyModernAppearance has not been implemented yet; baseline characterization remains active.")
            End If
            Return method
        End Function

        Private Shared Sub InvokeAppearance(method As MethodInfo, form As FormCreateMiddlePillars)
            Try
                method.Invoke(form, Nothing)
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Sub InvokeSetPileLayoutPreview(form As FormCreateMiddlePillars,
                                                      corners As Double(,), piles As Double(,), edges As Double(,))
            Dim method As MethodInfo = GetType(FormCreateMiddlePillars).GetMethod(
                "SetPileLayoutPreview",
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                Nothing,
                New Type() {GetType(Double(,)), GetType(Double(,)), GetType(Double(,))},
                Nothing)
            If method Is Nothing Then Throw New AssertionException("SetPileLayoutPreview(Double(,), Double(,), Double(,)) is missing.")
            Try
                method.Invoke(form, New Object() {corners, piles, edges})
            Catch ex As TargetInvocationException
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Sub SeedFunctionalState(form As FormCreateMiddlePillars)
            Dim combos As ComboBox() = Descendants(form).OfType(Of ComboBox)().ToArray()
            For index As Integer = 0 To combos.Length - 1
                combos(index).Items.Add("seed-" & index.ToString(CultureInfo.InvariantCulture))
                combos(index).SelectedIndex = 0
            Next
            form.CB_NumberPillar.Text = "pillar-seed"

            For index As Integer = 0 To AllCheckBoxes(form).Length - 1
                AllCheckBoxes(form)(index).Checked = (index Mod 2 = 0)
            Next
            form.NUpD_ElevationLand.Value = -1234.567D
            form.NUpD_CountRack.Value = 7D
            form.NUpD_OffsetColumnsPile.Value = 3456D
            form.NUpD_OffsetRowsPile.Value = -789D
            form.NUpD_CountColumnsPile.Value = 4D
            form.NUpD_CountRowsPile.Value = 5D
            Dim preview As NumericUpDown() = {form.NUpD_ScaleFront, form.NUpD_dxFront, form.NUpD_dyFront,
                                              form.NUpD_ScaleLeft, form.NUpD_dxLeft, form.NUpD_dyLeft,
                                              form.NUpD_ScaleRight, form.NUpD_dxRight, form.NUpD_dyRight}
            For index As Integer = 0 To preview.Length - 1
                preview(index).Value = If(index Mod 3 = 0, 2.5D, 100D + index)
            Next
            form.TxtB_PileCollDiagram.Text = "1-2-1"
            form.TxtB_PileRowDiagram.Text = "2-3-2"

            Dim grids As DataGridView() = {form.DGV_Rigel, form.DGV_SubFermenter, form.DGV_Rack,
                                           form.DGV_Grillage, form.DGV_Preparation, form.DGV_Piles}
            For index As Integer = 0 To grids.Length - 1
                Dim grid As DataGridView = grids(index)
                grid.AllowUserToAddRows = False
                grid.Tag = "grid-tag-" & index.ToString(CultureInfo.InvariantCulture)
                grid.Columns(0).Tag = "property-tag-" & index.ToString(CultureInfo.InvariantCulture)
                grid.Columns(0).ReadOnly = True
                If grid Is form.DGV_SubFermenter OrElse grid Is form.DGV_Rack OrElse grid Is form.DGV_Piles Then
                    Dim dynamicColumn As DataGridViewColumn = DirectCast(grid.Columns(1).Clone(), DataGridViewColumn)
                    dynamicColumn.Name = grid.Name & "Dynamic"
                    dynamicColumn.HeaderText = "Dynamic"
                    dynamicColumn.Tag = "dynamic-tag-" & index.ToString(CultureInfo.InvariantCulture)
                    grid.Columns.Add(dynamicColumn)
                End If
                Dim values(grid.Columns.Count - 1) As Object
                values(0) = "property-" & index.ToString(CultureInfo.InvariantCulture)
                For columnIndex As Integer = 1 To values.Length - 1
                    values(columnIndex) = "value-" & index.ToString(CultureInfo.InvariantCulture) & "-" & columnIndex.ToString(CultureInfo.InvariantCulture)
                Next
                Dim rowIndex As Integer = grid.Rows.Add(values)
                grid.Rows(rowIndex).Tag = "row-tag-" & index.ToString(CultureInfo.InvariantCulture)
                grid.ClearSelection()
                grid.Rows(rowIndex).Selected = True
                grid.CurrentCell = grid.Rows(rowIndex).Cells(Math.Min(1, grid.Columns.Count - 1))
            Next
        End Sub

        Private Shared Function CaptureSerializationTraversal(form As FormCreateMiddlePillars) As String()
            Dim result As New List(Of String)()
            For Each rootControl As Control In form.Controls
                Dim tabs As TabControl = TryCast(rootControl, TabControl)
                If tabs Is Nothing Then Continue For
                For Each page As TabPage In tabs.TabPages
                    For Each control As Control In page.Controls
                        If Not IsSerializedControl(control) Then Continue For
                        result.Add(ControlState(control))
                        Dim grid As DataGridView = TryCast(control, DataGridView)
                        If grid Is Nothing Then Continue For
                        For Each column As DataGridViewColumn In grid.Columns
                            result.Add("column|" & grid.Name & "|" & column.Name & "|" & column.GetType().Name & "|" &
                                       column.ReadOnly.ToString() & "|" & ValueText(column.Tag))
                        Next
                        For Each row As DataGridViewRow In grid.Rows
                            If row.IsNewRow Then Continue For
                            Dim values As String = String.Join(",", row.Cells.Cast(Of DataGridViewCell)().Select(Function(cell) ValueText(cell.Value)))
                            result.Add("row|" & grid.Name & "|" & row.Index.ToString(CultureInfo.InvariantCulture) & "|" &
                                       ValueText(row.Tag) & "|" & row.Selected.ToString() & "|" & values)
                        Next
                        result.Add("current|" & grid.Name & "|" & CellAddress(grid.CurrentCell))
                    Next
                Next
            Next
            Return result.ToArray()
        End Function

        Private Shared Function ControlState(control As Control) As String
            Dim prefix As String = "control|" & control.Parent.Name & "|" & control.Name & "|" & control.GetType().Name & "|" & ValueText(control.Tag) & "|"
            If TypeOf control Is NumericUpDown Then
                Dim number As NumericUpDown = DirectCast(control, NumericUpDown)
                Return prefix & DecimalText(number.Minimum) & "," & DecimalText(number.Maximum) & "," & DecimalText(number.Value) &
                       "," & number.ReadOnly.ToString() & "," & number.Enabled.ToString()
            End If
            If TypeOf control Is ComboBox Then
                Dim combo As ComboBox = DirectCast(control, ComboBox)
                Return prefix & combo.DropDownStyle.ToString() & "," & combo.SelectedIndex.ToString(CultureInfo.InvariantCulture) & "," & combo.Text
            End If
            If TypeOf control Is CheckBox Then Return prefix & DirectCast(control, CheckBox).Checked.ToString()
            If TypeOf control Is TextBox Then Return prefix & DirectCast(control, TextBox).Text
            If TypeOf control Is DataGridView Then Return prefix & DirectCast(control, DataGridView).ReadOnly.ToString()
            Return prefix
        End Function

        Private Shared Function CellAddress(cell As DataGridViewCell) As String
            If cell Is Nothing Then Return "nothing"
            Return cell.RowIndex.ToString(CultureInfo.InvariantCulture) & "," & cell.ColumnIndex.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function DecimalText(value As Decimal) As String
            Return value.ToString(CultureInfo.InvariantCulture)
        End Function

        Private Shared Function ValueText(value As Object) As String
            If value Is Nothing Then Return "<null>"
            Return Convert.ToString(value, CultureInfo.InvariantCulture).Replace("|", "||")
        End Function

        Private Shared Function CaptureSerializedParents(form As FormCreateMiddlePillars) As Dictionary(Of Control, TabPage)
            Dim result As New Dictionary(Of Control, TabPage)()
            For Each tabs As TabControl In New TabControl() {form.TabControl1, form.TabControl2}
                For Each page As TabPage In tabs.TabPages
                    For Each control As Control In page.Controls
                        If IsSerializedControl(control) Then result.Add(control, page)
                    Next
                Next
            Next
            Return result
        End Function

        Private Shared Function CaptureDesignerReferences(form As FormCreateMiddlePillars) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            For Each field As FieldInfo In GetType(FormCreateMiddlePillars).GetFields(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                Dim value As Object = field.GetValue(form)
                If TypeOf value Is Control OrElse TypeOf value Is DataGridViewColumn Then result.Add(field.Name, value)
            Next
            Return result
        End Function

        Private Shared Sub AssertDesignerReferences(form As FormCreateMiddlePillars, expected As Dictionary(Of String, Object))
            For Each pair As KeyValuePair(Of String, Object) In expected
                Dim field As FieldInfo = GetType(FormCreateMiddlePillars).GetField(pair.Key, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                Assert.That(field.GetValue(form), [Is].SameAs(pair.Value), pair.Key & " identity")
            Next
        End Sub

        Private Shared Function CapturePictureImages(form As FormCreateMiddlePillars) As Dictionary(Of String, Image)
            Return Descendants(form).OfType(Of PictureBox)().ToDictionary(Function(picture) picture.Name,
                                                                          Function(picture) picture.Image,
                                                                          StringComparer.Ordinal)
        End Function

        Private Shared Function CaptureVisualState(form As FormCreateMiddlePillars) As String()
            Return Descendants(form).Select(
                Function(control)
                    Dim autoScroll As Boolean = TypeOf control Is ScrollableControl AndAlso DirectCast(control, ScrollableControl).AutoScroll
                    Dim pictureMode As String = If(TypeOf control Is PictureBox, DirectCast(control, PictureBox).SizeMode.ToString(), String.Empty)
                    Return control.Parent.Name & "|" & control.Name & "|" & control.GetType().Name & "|" & control.Bounds.ToString() & "|" &
                           control.Dock.ToString() & "|" & CInt(control.Anchor).ToString(CultureInfo.InvariantCulture) & "|" &
                           autoScroll.ToString() & "|" & pictureMode
                End Function).ToArray()
        End Function

        Private Shared Function Descendants(parent As Control) As IEnumerable(Of Control)
            Dim result As New List(Of Control)()
            For Each child As Control In parent.Controls
                result.Add(child)
                result.AddRange(Descendants(child))
            Next
            Return result
        End Function

        Private Shared Function FindControl(parent As Control, name As String) As Control
            Return Descendants(parent).FirstOrDefault(Function(control) String.Equals(control.Name, name, StringComparison.Ordinal))
        End Function

        Private Shared Sub PerformLayoutTree(parent As Control)
            parent.PerformLayout()
            For Each child As Control In parent.Controls
                PerformLayoutTree(child)
            Next
            parent.PerformLayout()
            Application.DoEvents()
        End Sub

        Private Shared Sub ValidateTabs(tabs As TabControl)
            Assert.Multiple(
                Sub()
                    For index As Integer = 0 To tabs.TabCount - 1
                        tabs.SelectedIndex = index
                        Dim page As TabPage = tabs.TabPages(index)
                        PerformLayoutTree(tabs.FindForm())
                        For Each field As Control In page.Controls.Cast(Of Control)().Where(AddressOf IsSerializedControl)
                            If IsLegacyPreviewEditor(field) Then
                                Assert.That(field.Visible, [Is].False,
                                            field.Name & " must stay hidden while remaining in the legacy serialization hierarchy.")
                                Continue For
                            End If
                            If page.AutoScroll Then
                                page.ScrollControlIntoView(field)
                                PerformLayoutTree(page)
                                Assert.That(page.ClientRectangle.IntersectsWith(field.Bounds), [Is].True,
                                            field.Name & " must be reachable by scrolling " & page.Name & ".")
                            Else
                                AssertFullyInsideAncestor(field, page)
                            End If
                        Next
                        For Each picture As PictureBox In page.Controls.OfType(Of PictureBox)()
                            AssertFullyInsideAncestor(picture, page)
                        Next
                        For Each caption As Control In page.Controls.Cast(Of Control)().Where(
                            Function(control) TypeOf control Is Label OrElse TypeOf control Is Button)
                            AssertPreferredSizeFits(caption)
                        Next
                    Next
                End Sub)
        End Sub

        Private Shared Sub AssertFirstShownLayout(form As FormCreateMiddlePillars)
            Assert.That(form.TabControl1.SelectedTab, [Is].SameAs(form.TabPage1))
            Assert.That(form.TabControl2.SelectedTab, [Is].SameAs(form.TabPage6))
            AssertFullyInsideAncestor(form.PictureBox1, form.TabPage6)
            AssertLegacyPreviewEditorsHidden(form, form.TabPage6)
            AssertFullyInsideAncestor(form.DGV_Rigel, form.TabPage1)

            Dim footer As Control = FindControl(form, "ModernFooter")
            Assert.That(footer, [Is].Not.Null, "The modern footer must exist after appearance is applied.")
            For Each action As Button In New Button() {form.Button2, form.Button7, form.Button8}
                Assert.That(action.Parent, [Is].SameAs(footer), action.Name & " footer parent")
                AssertFullyInsideAncestor(action, footer)
                AssertFullyInsideAncestor(action, form)
            Next
        End Sub

        Private Shared Sub AssertCompactProjectGroupLayout(form As FormCreateMiddlePillars)
            Assert.That(form.GroupBox1.Height, [Is].LessThanOrEqualTo(ScaleForDpi(form, 160)),
                        "The project group must leave more vertical room for parameters and previews.")

            Dim firstLabels As Label() = {form.Label12, form.Label15, form.Label14, form.Label2}
            Dim firstInputs As Control() = {form.CBox_ListNamesArrProject, form.CBox_ListNamesBridge,
                                             form.CB_NameAlignment, form.CBox_ListNamesTemplateXML}
            Dim secondLabels As Label() = {form.Label13, form.Label16, form.Label26, form.Label47}
            Dim secondInputs As Control() = {form.CB_ProjectSurface, form.CB_EgSurface,
                                              form.CB_NumberPillar, form.ComboBox11}
            For Each fieldLabel As Label In firstLabels
                Assert.That(fieldLabel.Top, [Is].EqualTo(ScaleForDpi(form, 20)), fieldLabel.Name & " first label row")
            Next
            For Each input As Control In firstInputs
                Assert.That(input.Bounds.Y, [Is].EqualTo(ScaleForDpi(form, 44)), input.Name & " first input row")
                Assert.That(input.Height, [Is].GreaterThanOrEqualTo(input.GetPreferredSize(Size.Empty).Height),
                            input.Name & " preferred height")
                Assert.That(input.Height, [Is].LessThanOrEqualTo(ScaleForDpi(form, 26)), input.Name & " compact height")
            Next
            For Each fieldLabel As Label In secondLabels
                Assert.That(fieldLabel.Top, [Is].EqualTo(ScaleForDpi(form, 72)), fieldLabel.Name & " second label row")
            Next
            For Each input As Control In secondInputs
                Assert.That(input.Bounds.Y, [Is].EqualTo(ScaleForDpi(form, 96)), input.Name & " second input row")
                Assert.That(input.Height, [Is].GreaterThanOrEqualTo(input.GetPreferredSize(Size.Empty).Height),
                            input.Name & " preferred height")
                Assert.That(input.Height, [Is].LessThanOrEqualTo(ScaleForDpi(form, 26)), input.Name & " compact height")
            Next
            Assert.That(form.Button3.Top, [Is].EqualTo(ScaleForDpi(form, 96)))
            Assert.That(form.Button3.Height, [Is].EqualTo(ScaleForDpi(form, 26)))
            Assert.That(form.ChB_ProjectSurfaceInAlignment.Top, [Is].EqualTo(ScaleForDpi(form, 126)))
            Assert.That(form.ChB_ElevationLand.Top, [Is].EqualTo(ScaleForDpi(form, 126)))
            Assert.That(form.NUpD_ElevationLand.Top, [Is].EqualTo(ScaleForDpi(form, 124)))
            Assert.That(form.NUpD_ElevationLand.Height,
                        [Is].GreaterThanOrEqualTo(form.NUpD_ElevationLand.GetPreferredSize(Size.Empty).Height),
                        form.NUpD_ElevationLand.Name & " preferred height")
            Assert.That(form.NUpD_ElevationLand.Height, [Is].LessThanOrEqualTo(ScaleForDpi(form, 26)),
                        form.NUpD_ElevationLand.Name & " compact height")

            Dim visibleFields As Control() = form.GroupBox1.Controls.Cast(Of Control)().
                Where(Function(control) control.Visible).ToArray()
            For firstIndex As Integer = 0 To visibleFields.Length - 2
                For secondIndex As Integer = firstIndex + 1 To visibleFields.Length - 1
                    Assert.That(visibleFields(firstIndex).Bounds.IntersectsWith(visibleFields(secondIndex).Bounds),
                                [Is].False,
                                String.Format(CultureInfo.InvariantCulture,
                                              "{0} must not overlap {1}. {0}.Bounds={2}, {0}.PreferredSize={3}; " &
                                              "{1}.Bounds={4}, {1}.PreferredSize={5}; Form.Size={6}, ClientSize={7}, DPI={8}.",
                                              visibleFields(firstIndex).Name, visibleFields(secondIndex).Name,
                                              visibleFields(firstIndex).Bounds, visibleFields(firstIndex).GetPreferredSize(Size.Empty),
                                              visibleFields(secondIndex).Bounds, visibleFields(secondIndex).GetPreferredSize(Size.Empty),
                                              form.Size, form.ClientSize, form.DeviceDpi))
                Next
            Next
        End Sub

        Private Shared Sub AssertPreviewLayout(form As FormCreateMiddlePillars)
            form.TabControl2.SelectedTab = form.TabPage6
            PerformLayoutTree(form)
            AssertLegacyPreviewEditorsHidden(form, form.TabPage6)
            Assert.That(form.PictureBox1.Top, [Is].EqualTo(ScaleForDpi(form, 10)), "front preview top")
            AssertFullyInsideAncestor(form.PictureBox1, form.TabPage6)

            form.TabControl2.SelectedTab = form.TabPage7
            PerformLayoutTree(form)
            AssertLegacyPreviewEditorsHidden(form, form.TabPage7)
            Dim leftCaption As Control = FindControl(form.TabPage7, "ModernLeftPreviewCaption")
            Dim rightCaption As Control = FindControl(form.TabPage7, "ModernRightPreviewCaption")
            Assert.That(leftCaption, [Is].Not.Null)
            Assert.That(rightCaption, [Is].Not.Null)
            Assert.That(leftCaption.Top, [Is].EqualTo(ScaleForDpi(form, 10)))
            Assert.That(rightCaption.Top, [Is].EqualTo(ScaleForDpi(form, 10)))
            Assert.That(form.PictureBox2.Top, [Is].EqualTo(ScaleForDpi(form, 34)))
            Assert.That(form.PictureBox5.Top, [Is].EqualTo(ScaleForDpi(form, 34)))
            Assert.That(form.PictureBox2.Top, [Is].LessThanOrEqualTo(ScaleForDpi(form, 36)))
            Assert.That(form.PictureBox5.Top, [Is].LessThanOrEqualTo(ScaleForDpi(form, 36)))
            Assert.That(leftCaption.Left, [Is].GreaterThanOrEqualTo(form.PictureBox2.Left))
            Assert.That(leftCaption.Right, [Is].LessThanOrEqualTo(form.PictureBox2.Right))
            Assert.That(rightCaption.Left, [Is].GreaterThanOrEqualTo(form.PictureBox5.Left))
            Assert.That(rightCaption.Right, [Is].LessThanOrEqualTo(form.PictureBox5.Right))
            AssertFullyInsideAncestor(form.PictureBox2, form.TabPage7)
            AssertFullyInsideAncestor(form.PictureBox5, form.TabPage7)

            Dim layoutTab As TabPage = TryCast(FindControl(form, "ModernPileLayoutTab"), TabPage)
            Dim layoutPreview As PictureBox = TryCast(FindControl(form, "ModernPileLayoutPreview"), PictureBox)
            Assert.That(layoutTab, [Is].Not.Null)
            Assert.That(layoutPreview, [Is].Not.Null)
            form.TabControl2.SelectedTab = layoutTab
            PerformLayoutTree(form)
            AssertFullyInsideAncestor(layoutPreview, layoutTab)
        End Sub

        Private Shared Sub AssertLegacyPreviewEditorsHidden(form As FormCreateMiddlePillars, page As TabPage)
            Assert.That(form.TabControl2.SelectedTab, [Is].SameAs(page))
            For Each control As Control In LegacyPreviewEditors(form).Where(Function(item) item.Parent Is page)
                Assert.That(control.Visible, [Is].False, control.Name)
            Next
        End Sub

        Private Shared Function IsLegacyPreviewEditor(control As Control) As Boolean
            Return TypeOf control Is NumericUpDown AndAlso
                   (control.Name.StartsWith("NUpD_Scale", StringComparison.Ordinal) OrElse
                    control.Name.StartsWith("NUpD_dx", StringComparison.Ordinal) OrElse
                    control.Name.StartsWith("NUpD_dy", StringComparison.Ordinal))
        End Function

        Private Shared Function LegacyPreviewEditors(form As FormCreateMiddlePillars) As Control()
            Return New Control() {
                form.Label3, form.NUpD_ScaleFront, form.Label5, form.NUpD_dxFront, form.Label23, form.NUpD_dyFront,
                form.Label27, form.NUpD_ScaleLeft, form.Label25, form.NUpD_dxLeft, form.Label24, form.NUpD_dyLeft,
                form.Label30, form.NUpD_ScaleRight, form.Label29, form.NUpD_dxRight, form.Label28, form.NUpD_dyRight
            }
        End Function

        Private Shared Sub AssertPilePageLayout(form As FormCreateMiddlePillars)
            form.TabControl1.SelectedTab = form.TabPage5
            PerformLayoutTree(form)
            Assert.That(form.TabPage5.AutoScroll, [Is].True)
            Dim scaledMinimumHeight As Integer = CInt(Math.Round(380.0R * Math.Max(1.0R, form.DeviceDpi / 96.0R)))
            Assert.That(form.TabPage5.AutoScrollMinSize.Height, [Is].GreaterThanOrEqualTo(scaledMinimumHeight))

            Assert.That(form.Button1.Text, [Is].EqualTo("Применить схему"))
            AssertPreferredSizeFits(form.Button1)
            Dim nonOverlapping As Control() = {form.CB_PileTLC, form.Label9,
                                                form.ChB_InsertPileInRack, form.ChB_PileExpand,
                                                form.Label10}
            For firstIndex As Integer = 0 To nonOverlapping.Length - 2
                For secondIndex As Integer = firstIndex + 1 To nonOverlapping.Length - 1
                    Assert.That(nonOverlapping(firstIndex).Bounds.IntersectsWith(nonOverlapping(secondIndex).Bounds),
                                [Is].False,
                                nonOverlapping(firstIndex).Name & " must not overlap " & nonOverlapping(secondIndex).Name & ".")
                Next
            Next
        End Sub

        Private Shared Sub AssertFullyInsideAncestor(control As Control, ancestor As Control)
            Assert.That(control.Visible, [Is].True, control.Name & " visible")
            Assert.That(control.Width, [Is].GreaterThan(0), control.Name & " width")
            Assert.That(control.Height, [Is].GreaterThan(0), control.Name & " height")
            Dim bounds As Rectangle = ancestor.RectangleToClient(control.Parent.RectangleToScreen(control.Bounds))
            Assert.That(ancestor.ClientRectangle.Contains(bounds), [Is].True,
                        control.Name & " must fit fully inside " & ancestor.Name & ".")
        End Sub

        Private Shared Sub AssertPreferredSizeFits(control As Control)
            If Not control.Visible Then Return
            Dim label As Label = TryCast(control, Label)
            Dim proposed As Size = If(label IsNot Nothing AndAlso Not label.AutoSize,
                                      New Size(Math.Max(1, label.ClientSize.Width), 0), Size.Empty)
            Dim preferred As Size = control.GetPreferredSize(proposed)
            Assert.Multiple(
                Sub()
                    Assert.That(control.ClientSize.Width, [Is].GreaterThanOrEqualTo(preferred.Width),
                                control.Name & " preferred width")
                    Assert.That(control.ClientSize.Height, [Is].GreaterThanOrEqualTo(preferred.Height),
                                control.Name & " preferred height")
                End Sub)
        End Sub

        Private Shared Function ScaleForDpi(control As Control, logicalValue As Integer) As Integer
            Return CInt(Math.Round(logicalValue * Math.Max(1.0R, control.DeviceDpi / 96.0R)))
        End Function

        Private Shared Function MaximumVerticalScrollOffset(control As ScrollableControl) As Integer
            If Not control.VerticalScroll.Visible Then Return 0
            Return Math.Max(0, control.VerticalScroll.Maximum - control.VerticalScroll.LargeChange + 1)
        End Function
    End Class
End Namespace
