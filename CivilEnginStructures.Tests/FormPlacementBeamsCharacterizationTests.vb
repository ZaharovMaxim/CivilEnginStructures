Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class FormPlacementBeamsCharacterizationTests
        <Test>
        Public Sub DesignerInitializesExistingFunctionalDefaults()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                Assert.Multiple(
                    Sub()
                        AssertNumeric(form.NUpD_CountLeftRows, 0D, 100000000D, 3D, True)
                        AssertNumeric(form.NumericUpDown3, 0D, 999999999D, 960D, False)
                        AssertNumeric(form.NumericUpDown5, 0D, 9999999D, 7000D, False)
                        AssertNumeric(form.NUpD_CountRightRows, 0D, 100000000D, 3D, True)
                        AssertNumeric(form.NumericUpDown8, 0D, 999999999D, 960D, False)
                        AssertNumeric(form.NumericUpDown9, 0D, 9999999D, 7000D, False)
                        AssertNumeric(form.NUpD_dimLeftBridge, 0D, 99999999999D, 10000D, False)
                        AssertNumeric(form.NUpD_dimRightBridge, 0D, 9999999999999D, 10000D, False)
                        AssertNumeric(form.NUpD_CountProlet, 1D, 10000000D, 1D, True)
                        AssertNumeric(form.NUpD_TraverseOffset, -999999999999D, 999999999999D, 0D, False)
                        AssertNumeric(form.NUpD_VerticalOffset, 0D, 100D, 0D, False)
                    End Sub)
            End Using
            AssertPkMaskAndComboBoxEntryModes()
            AssertTableColumnsMenusAndBooleanChoices()
        End Sub

        Private Shared Sub AssertPkMaskAndComboBoxEntryModes()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                Assert.Multiple(
                    Sub()
                        Assert.That(form.MaskTB_PK.Mask, [Is].EqualTo("0+00\.000"))
                        Assert.That(form.MaskTB_PK.Text, [Is].EqualTo("0+00.000"))
                        Assert.That(form.MaskTB_PK.MaskFull, [Is].True)
                        Assert.That(form.MaskTB_PK.Enabled, [Is].False)

                        Assert.That(form.CBox_ListModelStructures.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))
                        Assert.That(form.CBox_ListProjectSurfaces.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))
                        Assert.That(form.CBox_ListAxisRoads.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))
                        Assert.That(form.CBox_ListTemplateXML.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))
                        Assert.That(form.CBox_ShemaPlacementBeams.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))

                        Assert.That(form.CBox_ListNameStructures.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDown))
                        Assert.That(form.CBox_ListPlacementBeams.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDown))
                        Assert.That(form.CBox_AlbumsBeams.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDown))
                    End Sub)
            End Using
        End Sub

        Private Shared Sub AssertTableColumnsMenusAndBooleanChoices()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                AssertColumns(form.DG_PillarsProperties,
                              New String() {"Column3", "Column1", "Column2", "Column4", "Column5"},
                              New Type() {GetType(DataGridViewCheckBoxColumn),
                                          GetType(DataGridViewTextBoxColumn),
                                          GetType(DataGridViewTextBoxColumn),
                                          GetType(DataGridViewTextBoxColumn),
                                          GetType(DataGridViewButtonColumn)})
                AssertColumns(form.DG_RowProperties,
                              New String() {"Column7", "Column8", "Column9"},
                              New Type() {GetType(DataGridViewTextBoxColumn),
                                          GetType(DataGridViewTextBoxColumn),
                                          GetType(DataGridViewButtonColumn)})
                Assert.That(form.DG_ProletListBeams.Columns, [Is].Empty)

                AssertMenu(form.ContextMenuStrip1,
                           New ToolStripItem() {form.ВставитьСтрокуВышеToolStripMenuItem,
                                                form.УдалитьСтрокуToolStripMenuItem,
                                                form.ПоменятьЗазорыМестамиToolStripMenuItem})
                AssertMenu(form.ContextMenuRowData,
                           New ToolStripItem() {form.РасчитатьОтступОтОсиToolStripMenuItem,
                                                form.ЗаполнитьТолщинуПокрытияToolStripMenuItem})
                AssertMenu(form.ContextMenuProletData,
                           New ToolStripItem() {form.ЗаполнитьВнизToolStripMenuItem,
                                                form.СменитьАльбомToolStripMenuItem})

                Assert.Multiple(
                    Sub()
                        Assert.That(form.ChB_CenterBeam.Checked, [Is].False)
                        Assert.That(form.CB_SurfaceFromAlign.Checked, [Is].False)
                        Assert.That(form.CheckBox3.Checked, [Is].False)
                        Assert.That(form.Column3.ThreeState, [Is].False)
                        Assert.That(form.DG_PillarsProperties.ContextMenuStrip, [Is].Null)
                        Assert.That(form.DG_RowProperties.ContextMenuStrip, [Is].SameAs(form.ContextMenuRowData))
                        Assert.That(form.DG_ProletListBeams.ContextMenuStrip, [Is].SameAs(form.ContextMenuProletData))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub AppearancePreservesFunctionalStateAndDesignerInstances()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the form.")

                Dim originalInstances As Dictionary(Of String, Object) = CaptureDesignerInstances(form)
                Dim before As String() = CaptureFunctionalState(form, originalInstances.Keys)

                InvokeAppearance(method, form)

                Assert.That(CaptureFunctionalState(form, originalInstances.Keys), [Is].EqualTo(before),
                            "Appearance must not change functional control state.")
                AssertDesignerInstancesArePreserved(form, originalInstances)
                AssertAllDesignerControlsAreReachable(form, originalInstances)
            End Using
        End Sub

        <Test>
        Public Sub AppearanceCanBeAppliedTwiceWithoutFurtherChanges()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the form.")

                InvokeAppearance(method, form)
                Dim afterFirstCall As String() = CaptureVisualState(form)

                InvokeAppearance(method, form)

                Assert.That(CaptureVisualState(form), [Is].EqualTo(afterFirstCall),
                            "A repeated appearance call must be idempotent.")
            End Using
        End Sub

        <TestCase(False, TestName:="AppearanceLayoutFitsPreferredWindowOnEveryTab")>
        <TestCase(True, TestName:="AppearanceLayoutFitsMinimumWindowOnEveryTab")>
        Public Sub AppearanceLayoutKeepsEditorsLabelsAndFooterUsable(useMinimumSize As Boolean)
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the form.")
                InvokeAppearance(method, form)

                If useMinimumSize Then
                    form.Size = form.MinimumSize
                End If
                form.CreateControl()
                PerformLayoutTree(form)

                Dim editorsByTab As Control()() = {
                    New Control() {form.NUpD_CountProlet, form.CBox_AlbumsBeams},
                    New Control() {form.NUpD_TraverseOffset, form.NUpD_VerticalOffset},
                    New Control() {form.MaskTB_PK, form.CheckBox3}
                }

                For tabIndex As Integer = 0 To form.TabControl1.TabCount - 1
                    form.TabControl1.SelectedIndex = tabIndex
                    PerformLayoutTree(form)
                    For Each editor As Control In editorsByTab(tabIndex)
                        Assert.That(editor.Width, [Is].GreaterThanOrEqualTo(48),
                                    editor.Name & " must retain a useful editor width.")
                        AssertInsideParent(editor)
                    Next
                Next

                Dim footer As Control = FindDescendant(form, "ModernFooter")
                Assert.That(footer, [Is].Not.Null, "The appearance footer must remain in the form layout.")
                AssertInsideParent(footer)
                AssertInsideParent(form.Button1)
                AssertInsideParent(form.Button2)
                AssertLabelTextFits(form.Label16)
                AssertLabelTextFits(form.Label2)
                AssertCardContentFits(form.GroupBox3, form.NUpD_dimLeftBridge)
                AssertCardContentFits(form.GroupBox4, form.NUpD_dimRightBridge)
            End Using
        End Sub

        <Test>
        Public Sub SavedPlacementOptionsRestoreEveryNonDefaultField()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                PrepareTypeChoices(form)

                InvokeRestoreSavedOptions(form,
                                          typeBridge:=2,
                                          transverseOffset:=-1.25R,
                                          verticalOffset:=0.075R,
                                          startPlacementPosition:=1234.567R,
                                          spanCount:=4,
                                          leftRowsCount:=2,
                                          rightRowsCount:=5,
                                          leftWidth:=8.5R,
                                          rightWidth:=9.25R)

                Dim parsedPK As Double
                Assert.Multiple(
                    Sub()
                        Assert.That(form.CBox_ListPlacementBeams.SelectedIndex, [Is].EqualTo(2))
                        Assert.That(form.NUpD_TraverseOffset.Value, [Is].EqualTo(-1250D))
                        Assert.That(form.NUpD_VerticalOffset.Value, [Is].EqualTo(75D))
                        Assert.That(form.CheckBox3.Checked, [Is].True)
                        Assert.That(form.MaskTB_PK.Enabled, [Is].True)
                        Assert.That(form.MaskTB_PK.Text, [Is].EqualTo("12+34.567"))
                        Assert.That(FuncFormatZn.TryParsePKText(form.MaskTB_PK.Text, parsedPK), [Is].True)
                        Assert.That(parsedPK, [Is].EqualTo(1234.567R).Within(0.000001R))
                        Assert.That(form.NUpD_CountProlet.Value, [Is].EqualTo(4D))
                        Assert.That(form.NUpD_CountLeftRows.Value, [Is].EqualTo(2D))
                        Assert.That(form.NUpD_CountRightRows.Value, [Is].EqualTo(5D))
                        Assert.That(form.NUpD_dimLeftBridge.Value, [Is].EqualTo(8500D))
                        Assert.That(form.NUpD_dimRightBridge.Value, [Is].EqualTo(9250D))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub RepeatedRestoreReplacesPriorValuesWithSavedBoundaryValues()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                PrepareTypeChoices(form)
                InvokeRestoreSavedOptions(form, 2, 1.5R, 0.1R, 945.25R, 6, 4, 3, 12R, 11R)

                InvokeRestoreSavedOptions(form,
                                          typeBridge:=0,
                                          transverseOffset:=0R,
                                          verticalOffset:=0R,
                                          startPlacementPosition:=0R,
                                          spanCount:=1,
                                          leftRowsCount:=0,
                                          rightRowsCount:=0,
                                          leftWidth:=0R,
                                          rightWidth:=0R)

                Assert.Multiple(
                    Sub()
                        Assert.That(form.CBox_ListPlacementBeams.SelectedIndex, [Is].EqualTo(0))
                        Assert.That(form.NUpD_TraverseOffset.Value, [Is].EqualTo(0D))
                        Assert.That(form.NUpD_VerticalOffset.Value, [Is].EqualTo(0D))
                        Assert.That(form.CheckBox3.Checked, [Is].False)
                        Assert.That(form.MaskTB_PK.Enabled, [Is].False)
                        Assert.That(form.MaskTB_PK.Text, [Is].EqualTo("0+00.000"))
                        Assert.That(form.NUpD_CountProlet.Value, [Is].EqualTo(1D))
                        Assert.That(form.NUpD_CountLeftRows.Value, [Is].EqualTo(0D))
                        Assert.That(form.NUpD_CountRightRows.Value, [Is].EqualTo(0D))
                        Assert.That(form.NUpD_dimLeftBridge.Value, [Is].EqualTo(0D))
                        Assert.That(form.NUpD_dimRightBridge.Value, [Is].EqualTo(0D))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub SavedFloatingTypeAndPkRemainSelectableAndParseable()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                PrepareTypeChoices(form)

                InvokeRestoreSavedOptions(form, 1, 0R, 0R, 98.765R, 1, 0, 0, 0R, 0R)

                Dim parsedPK As Double
                Assert.Multiple(
                    Sub()
                        Assert.That(form.CBox_ListPlacementBeams.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(form.CheckBox3.Checked, [Is].True)
                        Assert.That(form.MaskTB_PK.Text, [Is].EqualTo("0+98.765"))
                        Assert.That(FuncFormatZn.TryParsePKText(form.MaskTB_PK.Text, parsedPK), [Is].True)
                        Assert.That(parsedPK, [Is].EqualTo(98.765R).Within(0.000001R))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub CreateTablesSupportsBridgeWithoutBeamRows()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                form.NUpD_CountProlet.Value = 1D
                form.NUpD_CountLeftRows.Value = 0D
                form.NUpD_CountRightRows.Value = 0D
                form.ChB_CenterBeam.Checked = False

                Assert.That(InvokeCreateTables(form), [Is].True)

                Assert.Multiple(
                    Sub()
                        Assert.That(form.DG_PillarsProperties.RowCount, [Is].EqualTo(2))
                        Assert.That(form.DG_ProletListBeams.ColumnCount, [Is].EqualTo(1))
                        Assert.That(form.DG_ProletListBeams.RowCount, [Is].EqualTo(0))
                        Assert.That(form.DG_RowProperties.RowCount, [Is].EqualTo(0))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub CreateTablesUsesFirstOffsetForSingleRowOnEitherSide()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                form.NUpD_CountProlet.Value = 1D
                form.ChB_CenterBeam.Checked = False
                form.NUpD_CountLeftRows.Value = 1D
                form.NUpD_CountRightRows.Value = 0D

                Assert.That(InvokeCreateTables(form), [Is].True)
                Assert.Multiple(
                    Sub()
                        Assert.That(RowTags(form.DG_RowProperties), [Is].EqualTo(New Integer() {-1}))
                        Assert.That(RowOffsets(form.DG_RowProperties), [Is].EqualTo(New Integer() {960}))
                    End Sub)

                form.NUpD_CountLeftRows.Value = 0D
                form.NUpD_CountRightRows.Value = 1D

                Assert.That(InvokeCreateTables(form), [Is].True)
                Assert.Multiple(
                    Sub()
                        Assert.That(RowTags(form.DG_RowProperties), [Is].EqualTo(New Integer() {1}))
                        Assert.That(RowOffsets(form.DG_RowProperties), [Is].EqualTo(New Integer() {960}))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub CreateTablesPreservesLegacyThreeByThreeLayoutWithCenterRow()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                form.NUpD_CountProlet.Value = 1D
                form.NUpD_CountLeftRows.Value = 3D
                form.NUpD_CountRightRows.Value = 3D
                form.ChB_CenterBeam.Checked = True

                Assert.That(InvokeCreateTables(form), [Is].True)

                Dim expectedTags As Integer() = {-3, -2, -1, 0, 1, 2, 3}
                Dim expectedOffsets As Integer() = {7000, 3980, 960, 0, 960, 3980, 7000}
                Assert.Multiple(
                    Sub()
                        Assert.That(form.DG_ProletListBeams.RowCount, [Is].EqualTo(7))
                        Assert.That(RowTags(form.DG_ProletListBeams), [Is].EqualTo(expectedTags))
                        Assert.That(RowTags(form.DG_RowProperties), [Is].EqualTo(expectedTags))
                        Assert.That(RowOffsets(form.DG_RowProperties), [Is].EqualTo(expectedOffsets))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub CreateTablesTrimsSpanColumnsAndSupportsWhenSpanCountDecreases()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                form.ChB_CenterBeam.Checked = True
                form.NUpD_CountLeftRows.Value = 0D
                form.NUpD_CountRightRows.Value = 0D
                form.NUpD_CountProlet.Value = 3D

                Assert.That(InvokeCreateTables(form), [Is].True)
                Assert.Multiple(
                    Sub()
                        Assert.That(form.DG_ProletListBeams.ColumnCount, [Is].EqualTo(3))
                        Assert.That(form.DG_PillarsProperties.RowCount, [Is].EqualTo(4))
                    End Sub)

                form.NUpD_CountProlet.Value = 1D

                Assert.That(InvokeCreateTables(form), [Is].True)
                Assert.Multiple(
                    Sub()
                        Assert.That(form.DG_ProletListBeams.ColumnCount, [Is].EqualTo(1))
                        Assert.That(form.DG_PillarsProperties.RowCount, [Is].EqualTo(2))
                    End Sub)
            End Using
        End Sub

        Private Shared Function InvokeCreateTables(form As FormPlacementBeams) As Boolean
            Dim method As MethodInfo = GetType(FormPlacementBeams).GetMethod(
                "FuncCreateTables",
                BindingFlags.Instance Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=Type.EmptyTypes,
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "FuncCreateTables must remain available in the testable tables partial.")
            Try
                Return CBool(method.Invoke(form, Nothing))
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then Throw exception.InnerException
                Throw
            End Try
        End Function

        Private Shared Function RowTags(grid As DataGridView) As Integer()
            Dim result As New List(Of Integer)(grid.RowCount)
            For Each row As DataGridViewRow In grid.Rows
                result.Add(CInt(row.Tag))
            Next
            Return result.ToArray()
        End Function

        Private Shared Function RowOffsets(grid As DataGridView) As Integer()
            Dim result As New List(Of Integer)(grid.RowCount)
            For Each row As DataGridViewRow In grid.Rows
                result.Add(CInt(row.Cells(0).Value))
            Next
            Return result.ToArray()
        End Function

        Private Shared Sub PrepareTypeChoices(form As FormPlacementBeams)
            form.CBox_ListPlacementBeams.DataSource = New String() {
                "Фиксированная балка",
                "С расчетом длины балки",
                "С расчетом максимального зазора"
            }
        End Sub

        Private Shared Sub InvokeRestoreSavedOptions(form As FormPlacementBeams,
                                                     typeBridge As Integer,
                                                     transverseOffset As Double,
                                                     verticalOffset As Double,
                                                     startPlacementPosition As Double,
                                                     spanCount As Integer,
                                                     leftRowsCount As Integer,
                                                     rightRowsCount As Integer,
                                                     leftWidth As Double,
                                                     rightWidth As Double)
            Dim method As MethodInfo = GetType(FormPlacementBeams).GetMethod(
                "RestoreSavedPlacementOptions",
                BindingFlags.Instance Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {
                    GetType(Integer), GetType(Double), GetType(Double), GetType(Double),
                    GetType(Integer), GetType(Integer), GetType(Integer), GetType(Double), GetType(Double)
                },
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "RestoreSavedPlacementOptions must restore every persisted beam-placement option.")

            Try
                method.Invoke(form,
                              New Object() {
                                  typeBridge, transverseOffset, verticalOffset, startPlacementPosition,
                                  spanCount, leftRowsCount, rightRowsCount, leftWidth, rightWidth
                              })
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then Throw exception.InnerException
                Throw
            End Try
        End Sub

        Private Shared Sub AssertNumeric(control As NumericUpDown,
                                         expectedMinimum As Decimal,
                                         expectedMaximum As Decimal,
                                         expectedValue As Decimal,
                                         expectedReadOnly As Boolean)
            Assert.That(control.Minimum, [Is].EqualTo(expectedMinimum), control.Name & " minimum")
            Assert.That(control.Maximum, [Is].EqualTo(expectedMaximum), control.Name & " maximum")
            Assert.That(control.Value, [Is].EqualTo(expectedValue), control.Name & " value")
            Assert.That(control.ReadOnly, [Is].EqualTo(expectedReadOnly), control.Name & " read-only state")
        End Sub

        Private Shared Sub AssertColumns(grid As DataGridView,
                                         expectedNames As String(),
                                         expectedTypes As Type())
            Assert.That(grid.Columns.Count, [Is].EqualTo(expectedNames.Length), grid.Name & " column count")
            For index As Integer = 0 To expectedNames.Length - 1
                Dim currentIndex As Integer = index
                Assert.Multiple(
                    Sub()
                        Assert.That(grid.Columns(currentIndex).Name, [Is].EqualTo(expectedNames(currentIndex)),
                                    grid.Name & " column name at " & currentIndex)
                        Assert.That(grid.Columns(currentIndex).GetType(), [Is].EqualTo(expectedTypes(currentIndex)),
                                    grid.Name & " column type at " & currentIndex)
                    End Sub)
            Next
        End Sub

        Private Shared Sub PerformLayoutTree(parent As Control)
            parent.PerformLayout()
            For Each child As Control In parent.Controls
                PerformLayoutTree(child)
            Next
            parent.PerformLayout()
        End Sub

        Private Shared Sub AssertInsideParent(control As Control)
            Assert.That(control.Parent, [Is].Not.Null, control.Name & " must have a parent.")
            Assert.That(control.Width, [Is].GreaterThan(0), control.Name & " width")
            Assert.That(control.Height, [Is].GreaterThan(0), control.Name & " height")
            Assert.That(control.Parent.ClientRectangle.Contains(control.Bounds), [Is].True,
                        control.Name & " must fit inside " & control.Parent.Name & ".")
        End Sub

        Private Shared Sub AssertLabelTextFits(label As Label)
            AssertInsideParent(label)
            Dim preferred As Size = label.GetPreferredSize(New Size(label.ClientSize.Width, 0))
            Assert.That(label.ClientSize.Height, [Is].GreaterThanOrEqualTo(preferred.Height),
                        label.Name & " must have enough height for wrapped text.")
        End Sub

        Private Shared Sub AssertCardContentFits(card As GroupBox, bottomEditor As Control)
            Assert.That(card.Controls.Count, [Is].GreaterThan(0), card.Name & " must contain its fields table.")
            Dim fields As Control = card.Controls(0)
            Assert.That(fields, [Is].TypeOf(Of TableLayoutPanel)())
            Assert.That(card.DisplayRectangle.Contains(fields.Bounds), [Is].True,
                        card.Name & " fields must fit inside the card display area.")
            AssertInsideParent(bottomEditor)
        End Sub

        Private Shared Function FindDescendant(parent As Control, name As String) As Control
            If String.Equals(parent.Name, name, StringComparison.Ordinal) Then Return parent
            For Each child As Control In parent.Controls
                Dim match As Control = FindDescendant(child, name)
                If match IsNot Nothing Then Return match
            Next
            Return Nothing
        End Function

        Private Shared Sub AssertMenu(expectedMenu As ContextMenuStrip,
                                      expectedItems As ToolStripItem())
            Assert.That(expectedMenu.Items.Count, [Is].EqualTo(expectedItems.Length))
            For index As Integer = 0 To expectedItems.Length - 1
                Assert.That(expectedMenu.Items(index), [Is].SameAs(expectedItems(index)),
                            expectedMenu.Name & " item at " & index)
                Assert.That(DirectCast(expectedMenu.Items(index), ToolStripMenuItem).Checked, [Is].False,
                            expectedMenu.Items(index).Name & " checked state")
            Next
        End Sub

        Private Shared Function FindAppearanceMethod() As MethodInfo
            Return GetType(FormPlacementBeams).GetMethod(
                "ApplyModernAppearance",
                BindingFlags.Instance Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=Type.EmptyTypes,
                modifiers:=Nothing)
        End Function

        Private Shared Sub InvokeAppearance(method As MethodInfo, form As FormPlacementBeams)
            Try
                method.Invoke(form, Nothing)
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then
                    Throw exception.InnerException
                End If
                Throw
            End Try
        End Sub

        Private Shared Function CaptureDesignerInstances(form As FormPlacementBeams) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            For Each field As FieldInfo In DesignerFields()
                Dim value As Object = field.GetValue(form)
                If value IsNot Nothing AndAlso TypeOf value Is Component Then
                    result.Add(field.Name, value)
                End If
            Next
            Return result
        End Function

        Private Shared Sub AssertDesignerInstancesArePreserved(form As FormPlacementBeams,
                                                                originals As Dictionary(Of String, Object))
            For Each entry As KeyValuePair(Of String, Object) In originals
                Dim current As Object = GetType(FormPlacementBeams).GetField(
                    entry.Key,
                    BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic).GetValue(form)
                Assert.That(current, [Is].SameAs(entry.Value), entry.Key & " instance")
            Next
        End Sub

        Private Shared Sub AssertAllDesignerControlsAreReachable(form As FormPlacementBeams,
                                                                  originals As Dictionary(Of String, Object))
            Dim reachable As New HashSet(Of Control)(ReferenceEqualityComparer(Of Control).Instance)
            CollectControls(form, reachable)

            For Each entry As KeyValuePair(Of String, Object) In originals
                Dim control As Control = TryCast(entry.Value, Control)
                If control IsNot Nothing AndAlso Not TypeOf control Is ContextMenuStrip Then
                    Assert.That(reachable.Contains(control), [Is].True,
                                entry.Key & " must remain reachable from the form control tree.")
                End If
            Next
        End Sub

        Private Shared Sub CollectControls(parent As Control, result As HashSet(Of Control))
            result.Add(parent)
            For Each child As Control In parent.Controls
                CollectControls(child, result)
            Next
        End Sub

        Private Shared Function CaptureFunctionalState(form As FormPlacementBeams,
                                                        originalFieldNames As IEnumerable(Of String)) As String()
            Dim state As New List(Of String) From {
                "Mask=" & form.MaskTB_PK.Mask,
                "MaskText=" & form.MaskTB_PK.Text,
                "MaskEnabled=" & form.MaskTB_PK.Enabled,
                "MaskFull=" & form.MaskTB_PK.MaskFull,
                "PillarMenu=" & RuntimeHelpers.GetHashCode(form.DG_PillarsProperties.ContextMenuStrip),
                "RowMenu=" & RuntimeHelpers.GetHashCode(form.DG_RowProperties.ContextMenuStrip),
                "SpanMenu=" & RuntimeHelpers.GetHashCode(form.DG_ProletListBeams.ContextMenuStrip)
            }

            For Each fieldName As String In originalFieldNames
                Dim field As FieldInfo = GetType(FormPlacementBeams).GetField(
                    fieldName,
                    BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                Dim value As Object = field.GetValue(form)
                Dim control As Control = TryCast(value, Control)
                If control IsNot Nothing Then
                    state.Add(field.Name & ".Enabled=" & control.Enabled)
                End If

                Dim numeric As NumericUpDown = TryCast(value, NumericUpDown)
                If numeric IsNot Nothing Then
                    state.Add(field.Name & ".Numeric=" & numeric.Minimum & ";" & numeric.Maximum & ";" &
                              numeric.Value & ";" & numeric.ReadOnly & ";" & numeric.DecimalPlaces & ";" & numeric.Increment)
                End If

                Dim combo As ComboBox = TryCast(value, ComboBox)
                If combo IsNot Nothing Then
                    state.Add(field.Name & ".Combo=" & combo.DropDownStyle & ";" & combo.Sorted & ";" & combo.SelectedIndex)
                End If

                Dim checkBox As CheckBox = TryCast(value, CheckBox)
                If checkBox IsNot Nothing Then
                    state.Add(field.Name & ".Check=" & checkBox.Checked & ";" & checkBox.CheckState & ";" & checkBox.ThreeState)
                End If

                Dim menuItem As ToolStripMenuItem = TryCast(value, ToolStripMenuItem)
                If menuItem IsNot Nothing Then
                    state.Add(field.Name & ".Menu=" & menuItem.Checked & ";" & menuItem.CheckState & ";" &
                              menuItem.CheckOnClick & ";" & menuItem.Enabled)
                End If
            Next

            AppendGridState(state, form.DG_PillarsProperties)
            AppendGridState(state, form.DG_RowProperties)
            AppendGridState(state, form.DG_ProletListBeams)
            state.Sort(StringComparer.Ordinal)
            Return state.ToArray()
        End Function

        Private Shared Sub AppendGridState(state As List(Of String), grid As DataGridView)
            state.Add(grid.Name & ".Grid=" & grid.AllowUserToAddRows & ";" & grid.AllowUserToDeleteRows & ";" &
                      grid.ReadOnly & ";" & grid.MultiSelect & ";" & grid.SelectionMode)
            For index As Integer = 0 To grid.Columns.Count - 1
                Dim column As DataGridViewColumn = grid.Columns(index)
                state.Add(grid.Name & ".Column" & index & "=" & column.Name & ";" &
                          column.GetType().FullName & ";" & column.DataPropertyName & ";" &
                          column.ReadOnly & ";" & column.Visible & ";" & column.SortMode)
            Next
        End Sub

        Private Shared Function CaptureVisualState(form As FormPlacementBeams) As String()
            Dim state As New List(Of String)
            For Each field As FieldInfo In DesignerFields()
                Dim control As Control = TryCast(field.GetValue(form), Control)
                If control IsNot Nothing Then
                    state.Add(field.Name & "=" & control.GetType().FullName & ";" &
                              control.Bounds.ToString() & ";" & control.Dock & ";" & control.Anchor & ";" &
                              control.Margin.ToString() & ";" & control.Padding.ToString() & ";" &
                              control.BackColor.ToArgb() & ";" & control.ForeColor.ToArgb() & ";" &
                              FontKey(control.Font))
                End If
            Next
            state.Sort(StringComparer.Ordinal)
            Return state.ToArray()
        End Function

        Private Shared Function FontKey(font As Font) As String
            Return font.Name & ";" & font.SizeInPoints & ";" & CInt(font.Style) & ";" & font.Unit
        End Function

        Private Shared Function DesignerFields() As FieldInfo()
            Return GetType(FormPlacementBeams).GetFields(
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
        End Function

        Private NotInheritable Class ReferenceEqualityComparer(Of T As Class)
            Implements IEqualityComparer(Of T)

            Friend Shared ReadOnly Instance As New ReferenceEqualityComparer(Of T)()

            Private Sub New()
            End Sub

            Public Overloads Function Equals(left As T, right As T) As Boolean Implements IEqualityComparer(Of T).Equals
                Return Object.ReferenceEquals(left, right)
            End Function

            Public Overloads Function GetHashCode(value As T) As Integer Implements IEqualityComparer(Of T).GetHashCode
                Return RuntimeHelpers.GetHashCode(value)
            End Function
        End Class
    End Class
End Namespace
