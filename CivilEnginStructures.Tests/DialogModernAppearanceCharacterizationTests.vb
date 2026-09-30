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
    Public Class DialogModernAppearanceCharacterizationTests
        <Test>
        Public Sub AppearanceContractsArePrivateAndCalledOnceAfterDesignerInitialization()
            AssertAppearanceContract(
                GetType(FormCreateConeLastPillars),
                Path.Combine("UserForms", "Bridge", "FormCreateConeLastPillars.vb"))
            AssertAppearanceContract(
                GetType(FormCreateMonolitSitesBeams),
                Path.Combine("UserForms", "Bridge", "FormCreateMonolitSitesBeams.vb"))
            AssertAppearanceContract(
                GetType(FormUserTemptate),
                Path.Combine("UserForms", "FormUserTemptate.vb"))
        End Sub

        <Test>
        Public Sub ConeAppearancePreservesTheFifteenRowEditorAndHidesTheRetiredSiteControls()
            Using form As FormCreateConeLastPillars = FormCreateConeLastPillars.CreateDesignerOnly()
                Dim grid As DataGridView = form.DGV_PropertiesCone
                Dim parameterColumn As DataGridViewColumn = form.Номер
                Dim valueColumn As DataGridViewColumn = form.Column1
                Dim gridParent As Control = grid.Parent
                Dim segmentParent As Control = form.NUpD_CountSegmets.Parent
                Dim actions As Button() = {form.Button2, form.Button7, form.Button8}
                Dim actionResults As DialogResult() = actions.Select(Function(item) item.DialogResult).ToArray()

                Assert.Multiple(
                    Sub()
                        Assert.That(grid.ColumnCount, [Is].EqualTo(2))
                        Assert.That(grid.Columns(0), [Is].SameAs(parameterColumn))
                        Assert.That(grid.Columns(1), [Is].SameAs(valueColumn))
                        Assert.That(parameterColumn.Name, [Is].EqualTo("Номер"))
                        Assert.That(parameterColumn.HeaderText, [Is].EqualTo("Параметр"))
                        Assert.That(valueColumn.Name, [Is].EqualTo("Column1"))
                        Assert.That(valueColumn.HeaderText, [Is].EqualTo("Значение"))
                        Assert.That(form.NUpD_CountSegmets.Minimum, [Is].EqualTo(0D))
                        Assert.That(form.NUpD_CountSegmets.Maximum, [Is].EqualTo(100D))
                        Assert.That(form.NUpD_CountSegmets.Value, [Is].EqualTo(5D))
                    End Sub)

                grid.Rows.Clear()
                grid.Rows.Add(15)
                Dim rows As DataGridViewRow() = WorkingRows(grid)
                For index As Integer = 0 To rows.Length - 1
                    rows(index).Cells(0).Value = "Параметр " & index.ToString()
                    rows(index).Cells(1).Value = index + 0.25R
                    rows(index).Tag = "cone-row-" & index.ToString()
                Next
                grid.Tag = "cone-grid-tag"
                form.NUpD_CountSegmets.Value = 17D
                form.NUpD_CountSegmets.Tag = "segments-tag"
                form.CBox_ListNamesArrProject.Tag = "bridge-model-tag"

                ApplyAppearance(form)
                ApplyAppearance(form)
                ShowOffscreen(form, form.Size)

                Assert.Multiple(
                    Sub()
                        Assert.That(form.DGV_PropertiesCone, [Is].SameAs(grid))
                        Assert.That(form.Номер, [Is].SameAs(parameterColumn))
                        Assert.That(form.Column1, [Is].SameAs(valueColumn))
                        Assert.That(grid.Parent, [Is].SameAs(gridParent))
                        Assert.That(form.NUpD_CountSegmets.Parent, [Is].SameAs(segmentParent))
                        Assert.That(WorkingRows(grid), [Is].EqualTo(rows))
                        Assert.That(WorkingRows(grid).Select(Function(row) CStr(row.Tag)),
                                    [Is].EqualTo(rows.Select(Function(row) CStr(row.Tag))))
                        Assert.That(form.NUpD_CountSegmets.Value, [Is].EqualTo(17D))
                        Assert.That(form.NUpD_CountSegmets.Tag, [Is].EqualTo("segments-tag"))
                        Assert.That(form.CBox_ListNamesArrProject.Tag, [Is].EqualTo("bridge-model-tag"))
                        Assert.That(grid.Tag, [Is].EqualTo("cone-grid-tag"))
                        Assert.That(form.Label5.Visible, [Is].False)
                        Assert.That(form.CB_NameSites.Visible, [Is].False)
                        Assert.That(form.Button1.Visible, [Is].False)
                    End Sub)
                Dim currentActions As Button() = {form.Button2, form.Button7, form.Button8}
                For index As Integer = 0 To actions.Length - 1
                    Assert.That(currentActions(index), [Is].SameAs(actions(index)), actions(index).Name)
                    Assert.That(currentActions(index).DialogResult, [Is].EqualTo(actionResults(index)), actions(index).Name)
                Next
                For index As Integer = 0 To rows.Length - 1
                    Assert.That(CStr(rows(index).Cells(0).Value), [Is].EqualTo("Параметр " & index.ToString()))
                    Assert.That(CDbl(rows(index).Cells(1).Value), [Is].EqualTo(index + 0.25R))
                Next
            End Using
        End Sub

        <Test>
        Public Sub ConeLayoutKeepsTheGridAndActionsReadableAtPreferredAndMinimumSizes()
            Using form As FormCreateConeLastPillars = FormCreateConeLastPillars.CreateDesignerOnly()
                SeedConeRows(form)
                ApplyAppearance(form)
                Assert.That(form.MinimumSize.Width, [Is].GreaterThanOrEqualTo(900))
                Assert.That(form.MinimumSize.Height, [Is].GreaterThanOrEqualTo(620))
                Assert.That(form.Size.Width, [Is].GreaterThanOrEqualTo(form.MinimumSize.Width))
                Assert.That(form.Size.Height, [Is].GreaterThanOrEqualTo(form.MinimumSize.Height))

                For Each target As Size In New Size() {form.Size, form.MinimumSize}
                    ShowOffscreen(form, target)
                    AssertReadableInside(form, form.GroupBox1, 300, 120)
                    AssertReadableInside(form, form.GroupBox2, 300, 220)
                    AssertReadableInside(form, form.DGV_PropertiesCone, 380, 180)
                    For Each action As Button In New Button() {form.Button2, form.Button7, form.Button8}
                        AssertReadableInside(form, action, 80, 24)
                        Assert.That(ScreenBounds(form.DGV_PropertiesCone).IntersectsWith(ScreenBounds(action)),
                                    [Is].False, action.Name & " must not cover the cone grid.")
                    Next
                Next
            End Using
        End Sub

        <Test>
        Public Sub MonolithAppearancePreservesBothIndependentRadioGroupsAndDisabledChoice()
            Using form As FormCreateMonolitSitesBeams = FormCreateMonolitSitesBeams.CreateDesignerOnly()
                Dim radios As RadioButton() = {
                    form.RadioButton1, form.RadioButton2, form.RadioButton3,
                    form.RadioButton4, form.RadioButton5}
                Dim parents As Control() = radios.Select(Function(item) item.Parent).ToArray()
                Dim cancelButton As Button = form.Button1
                Dim createButton As Button = form.Button2
                Dim cancelResult As DialogResult = cancelButton.DialogResult
                Dim createResult As DialogResult = createButton.DialogResult
                Assert.Multiple(
                    Sub()
                        Assert.That(form.RadioButton4.Parent, [Is].SameAs(form.GroupBox1))
                        Assert.That(form.RadioButton5.Parent, [Is].SameAs(form.GroupBox1))
                        Assert.That(form.RadioButton1.Parent, [Is].SameAs(form.GroupBox2))
                        Assert.That(form.RadioButton2.Parent, [Is].SameAs(form.GroupBox2))
                        Assert.That(form.RadioButton3.Parent, [Is].SameAs(form.GroupBox2))
                        Assert.That(form.RadioButton4.Checked, [Is].True)
                        Assert.That(form.RadioButton1.Checked, [Is].True)
                        Assert.That(form.RadioButton3.Enabled, [Is].False)
                    End Sub)

                form.RadioButton5.Checked = True
                form.RadioButton2.Checked = True
                form.RadioButton3.Tag = "disabled-manual-choice"
                form.CBox_ListNamesArrProject.Items.AddRange(New Object() {"Мосты 1", "Мосты 2"})
                form.CBox_ListNamesArrProject.SelectedIndex = 1
                form.CBox_ListNamesArrProject.Tag = "model-choice-tag"

                ApplyAppearance(form)
                ApplyAppearance(form)

                For index As Integer = 0 To radios.Length - 1
                    Assert.That(radios(index).Parent, [Is].SameAs(parents(index)), radios(index).Name)
                Next
                Assert.Multiple(
                    Sub()
                        Assert.That(form.RadioButton5.Checked, [Is].True)
                        Assert.That(form.RadioButton4.Checked, [Is].False)
                        Assert.That(form.RadioButton2.Checked, [Is].True)
                        Assert.That(form.RadioButton1.Checked, [Is].False)
                        Assert.That(form.RadioButton3.Enabled, [Is].False)
                        Assert.That(form.RadioButton3.Tag, [Is].EqualTo("disabled-manual-choice"))
                        Assert.That(form.CBox_ListNamesArrProject.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(form.CBox_ListNamesArrProject.Tag, [Is].EqualTo("model-choice-tag"))
                        Assert.That(form.Button1, [Is].SameAs(cancelButton))
                        Assert.That(form.Button2, [Is].SameAs(createButton))
                        Assert.That(form.Button1.DialogResult, [Is].EqualTo(cancelResult))
                        Assert.That(form.Button2.DialogResult, [Is].EqualTo(createResult))
                    End Sub)
            End Using
        End Sub

        <Test>
        Public Sub MonolithLayoutKeepsBothRadioGroupsAndActionsReadableAtTwoSizes()
            Using form As FormCreateMonolitSitesBeams = FormCreateMonolitSitesBeams.CreateDesignerOnly()
                ApplyAppearance(form)
                Assert.That(form.MinimumSize.Width, [Is].GreaterThanOrEqualTo(760))
                Assert.That(form.MinimumSize.Height, [Is].GreaterThanOrEqualTo(560))
                Assert.That(form.Size.Width, [Is].GreaterThanOrEqualTo(form.MinimumSize.Width))
                Assert.That(form.Size.Height, [Is].GreaterThanOrEqualTo(form.MinimumSize.Height))

                For Each target As Size In New Size() {form.Size, form.MinimumSize}
                    ShowOffscreen(form, target)
                    AssertReadableInside(form, form.Panel1, 300, 100)
                    AssertReadableInside(form, form.Panel2, 300, 180)
                    AssertReadableInside(form, form.GroupBox1, 260, 60)
                    AssertReadableInside(form, form.GroupBox2, 260, 100)
                    For Each radio As RadioButton In New RadioButton() {
                        form.RadioButton1, form.RadioButton2, form.RadioButton3,
                        form.RadioButton4, form.RadioButton5}
                        Assert.That(radio.Width, [Is].GreaterThan(80), radio.Name)
                        Assert.That(radio.Parent.ClientRectangle.Contains(radio.Bounds), [Is].True, radio.Name)
                    Next
                    AssertReadableInside(form, form.Button1, 80, 24)
                    AssertReadableInside(form, form.Button2, 80, 24)
                Next
            End Using
        End Sub

        <Test>
        Public Sub LibraryAppearancePreservesItemsSelectionPathTagsAndButtonRoles()
            Using form As FormUserTemptate = FormUserTemptate.CreateDesignerOnly()
                Dim list As ListBox = form.ListBox1
                Dim editor As TextBox = form.TextBox1
                Dim pathLabel As Label = form.Label3
                Dim listParent As Control = list.Parent
                Dim editorParent As Control = editor.Parent
                Dim actions As Button() = {form.Button1, form.Button2, form.Button3}
                Dim actionResults As DialogResult() = actions.Select(Function(item) item.DialogResult).ToArray()
                list.Items.AddRange(New Object() {"Схема А", "Схема Б", "Схема В"})
                list.SelectedIndex = 1
                editor.Text = "Пользовательская схема"
                pathLabel.Text = "D:\\Шаблоны мостов"
                list.Tag = "library-items-tag"
                editor.Tag = "library-name-tag"
                pathLabel.Tag = "library-path-tag"
                form.Button1.Tag = "save-tag"
                form.Button2.Tag = "delete-tag"
                form.Button3.Tag = "close-tag"

                ApplyAppearance(form)
                ApplyAppearance(form)

                Assert.Multiple(
                    Sub()
                        Assert.That(form.ListBox1, [Is].SameAs(list))
                        Assert.That(form.TextBox1, [Is].SameAs(editor))
                        Assert.That(form.Label3, [Is].SameAs(pathLabel))
                        Assert.That(list.Parent, [Is].SameAs(listParent))
                        Assert.That(editor.Parent, [Is].SameAs(editorParent))
                        Assert.That(list.Items.Cast(Of String)(),
                                    [Is].EqualTo(New String() {"Схема А", "Схема Б", "Схема В"}))
                        Assert.That(list.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(editor.Text, [Is].EqualTo("Пользовательская схема"))
                        Assert.That(pathLabel.Text, [Is].EqualTo("D:\\Шаблоны мостов"))
                        Assert.That(list.Tag, [Is].EqualTo("library-items-tag"))
                        Assert.That(editor.Tag, [Is].EqualTo("library-name-tag"))
                        Assert.That(pathLabel.Tag, [Is].EqualTo("library-path-tag"))
                        Assert.That(form.Button1.Tag, [Is].EqualTo("save-tag"))
                        Assert.That(form.Button2.Tag, [Is].EqualTo("delete-tag"))
                        Assert.That(form.Button3.Tag, [Is].EqualTo("close-tag"))
                    End Sub)
                Dim currentActions As Button() = {form.Button1, form.Button2, form.Button3}
                For index As Integer = 0 To actions.Length - 1
                    Assert.That(currentActions(index), [Is].SameAs(actions(index)), actions(index).Name)
                    Assert.That(currentActions(index).DialogResult, [Is].EqualTo(actionResults(index)), actions(index).Name)
                Next
            End Using
        End Sub

        <Test>
        Public Sub LibraryLayoutKeepsEditorListPathAndAllActionsReadableAtTwoSizes()
            Using form As FormUserTemptate = FormUserTemptate.CreateDesignerOnly()
                form.ListBox1.Items.AddRange(New Object() {"Схема А", "Схема Б"})
                form.Label3.Text = "D:\\Шаблоны мостов"
                ApplyAppearance(form)
                Assert.That(form.MinimumSize.Width, [Is].GreaterThanOrEqualTo(600))
                Assert.That(form.MinimumSize.Height, [Is].GreaterThanOrEqualTo(460))
                Assert.That(form.Size.Width, [Is].GreaterThanOrEqualTo(form.MinimumSize.Width))
                Assert.That(form.Size.Height, [Is].GreaterThanOrEqualTo(form.MinimumSize.Height))

                For Each target As Size In New Size() {form.Size, form.MinimumSize}
                    ShowOffscreen(form, target)
                    AssertReadableInside(form, form.GroupBox1, 360, 260)
                    AssertReadableInside(form, form.TextBox1, 260, 20)
                    AssertReadableInside(form, form.ListBox1, 300, 180)
                    AssertReadableInside(form, form.Label3, 120, 13)
                    Dim nameLabelBounds As Rectangle = ScreenBounds(form.Label1)
                    Dim nameEditorBounds As Rectangle = ScreenBounds(form.TextBox1)
                    TestContext.Progress.WriteLine(
                        "Library {0}x{1}: Label1={2}; TextBox1={3}; intersects={4}",
                        form.Width, form.Height, nameLabelBounds, nameEditorBounds,
                        nameLabelBounds.IntersectsWith(nameEditorBounds))
                    Assert.That(nameLabelBounds.IntersectsWith(nameEditorBounds), [Is].False,
                                "The library name label must not overlap its editor.")
                    For Each action As Button In New Button() {form.Button1, form.Button2, form.Button3}
                        AssertReadableInside(form, action, 80, 24)
                    Next
                Next
            End Using
        End Sub

        <Test, Explicit("Manual offline visual QA; writes preferred and minimum PNG previews to the test work directory.")>
        <Category("ManualVisual")>
        Public Sub RenderModernDialogPreviewsToPng()
            Dim outputDirectory As String = Path.Combine(
                TestContext.CurrentContext.WorkDirectory, "dialog-modern-previews")
            Directory.CreateDirectory(outputDirectory)
            Dim cone As FormCreateConeLastPillars = FormCreateConeLastPillars.CreateDesignerOnly()
            Dim monolith As FormCreateMonolitSitesBeams = FormCreateMonolitSitesBeams.CreateDesignerOnly()
            Dim library As FormUserTemptate = FormUserTemptate.CreateDesignerOnly()
            Try
                PopulateConePreview(cone)
                PopulateMonolithPreview(monolith)
                PopulateLibraryPreview(library)
                ApplyAppearance(cone)
                ApplyAppearance(monolith)
                ApplyAppearance(library)

                RenderPreview(cone, outputDirectory, "cone-preferred.png", New Size(1120, 800))
                RenderPreview(cone, outputDirectory, "cone-minimum.png", cone.MinimumSize)
                RenderPreview(monolith, outputDirectory, "monolith-preferred.png", New Size(900, 650))
                RenderPreview(monolith, outputDirectory, "monolith-minimum.png", monolith.MinimumSize)
                RenderPreview(library, outputDirectory, "library-preferred.png", New Size(720, 560))
                RenderPreview(library, outputDirectory, "library-minimum.png", library.MinimumSize)
            Finally
                cone.Dispose()
                monolith.Dispose()
                library.Dispose()
            End Try
        End Sub

        Private Shared Sub AssertAppearanceContract(formType As Type, relativeSourcePath As String)
            Dim method As MethodInfo = formType.GetMethod(
                "ApplyModernAppearance", BindingFlags.Instance Or BindingFlags.NonPublic)
            Assert.That(method, [Is].Not.Null, formType.Name)
            Assert.That(method.IsPrivate, [Is].True, formType.Name)
            Dim stateFields As FieldInfo() = formType.GetFields(
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.DeclaredOnly).
                Where(Function(field) field.Name.IndexOf("modern", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                      field.Name.IndexOf("appearance", StringComparison.OrdinalIgnoreCase) >= 0).
                ToArray()
            For Each field As FieldInfo In stateFields
                Assert.That(field.IsPrivate, [Is].True, formType.Name & "." & field.Name)
            Next

            Dim source As String = File.ReadAllText(Path.Combine(RepositoryRoot(), relativeSourcePath))
            Dim calls As MatchCollection = Regex.Matches(source, "\bApplyModernAppearance\s*\(\s*\)")
            Assert.That(calls.Count, [Is].EqualTo(1), formType.Name)
            Dim initializeIndex As Integer = source.IndexOf("InitializeComponent()", StringComparison.Ordinal)
            Dim appearanceIndex As Integer = source.IndexOf("ApplyModernAppearance()", StringComparison.Ordinal)
            Assert.That(initializeIndex, [Is].GreaterThanOrEqualTo(0), formType.Name)
            Assert.That(appearanceIndex, [Is].GreaterThan(initializeIndex),
                        formType.Name & " must apply appearance after InitializeComponent.")
        End Sub

        Private Shared Sub ApplyAppearance(form As Form)
            Dim method As MethodInfo = form.GetType().GetMethod(
                "ApplyModernAppearance", BindingFlags.Instance Or BindingFlags.NonPublic)
            Assert.That(method, [Is].Not.Null, form.GetType().Name)
            Try
                method.Invoke(form, Nothing)
            Catch ex As TargetInvocationException When ex.InnerException IsNot Nothing
                Throw ex.InnerException
            End Try
        End Sub

        Private Shared Sub SeedConeRows(form As FormCreateConeLastPillars)
            Dim labels As String() = {
                "Длина конуса слева до края откосного крыла, м",
                "Длина конуса справа до края откосного крыла, м",
                "Заложение откоса конуса слева",
                "Заложение откоса конуса справа",
                "Ширина бермы слева, м",
                "Ширина бермы справа, м",
                "Толщина укрепления откоса, м",
                "Высота первого участка конуса, м",
                "Высота второго участка конуса, м",
                "Глубина заложения подошвы, м",
                "Смещение начала конуса слева, м",
                "Смещение начала конуса справа, м",
                "Дополнительный уклон слева, ‰",
                "Дополнительный уклон справа, ‰",
                "Отметка контрольной точки конуса, м"}
            Dim values As Object() = {
                5.0R, 5.0R, 1.0R, 1.0R, 0.75R,
                0.75R, 0.15R, 1.0R, 2.0R, 0.3R,
                0.0R, 0.0R, 0.0R, 0.0R, 127.45R}
            form.DGV_PropertiesCone.Rows.Clear()
            form.DGV_PropertiesCone.Rows.Add(15)
            For index As Integer = 0 To 14
                form.DGV_PropertiesCone.Rows(index).Cells(0).Value = labels(index)
                form.DGV_PropertiesCone.Rows(index).Cells(1).Value = values(index)
            Next
        End Sub

        Private Shared Sub PopulateConePreview(form As FormCreateConeLastPillars)
            SeedConeRows(form)
            SelectOnly(form.CBox_ListNamesArrProject,
                       "Мосты — путепровод через реку Большая Липовица")
            SelectOnly(form.CBox_ListNamesBridge,
                       "Путепровод на ПК 124+56,78 (левое направление)")
            SelectOnly(form.CB_ProjectSurface,
                       "Проектная поверхность автомобильной дороги — основной ход")
            SelectOnly(form.CB_EgSurface,
                       "Поверхность существующей земли по инженерным изысканиям")
            SelectOnly(form.CB_NameAlignment,
                       "Ось трассы М-12, направление Москва — Казань")
            SelectOnly(form.CBox_ListNamesTemplateXML,
                       "Конус насыпи с укреплением откоса монолитным бетоном")
            SelectOnly(form.CB_NumberPillar, "Опора 12 — береговая")
            form.NUpD_CountSegmets.Value = 8D
        End Sub

        Private Shared Sub PopulateMonolithPreview(form As FormCreateMonolitSitesBeams)
            SelectOnly(form.CBox_ListNamesArrProject,
                       "Мосты — путепровод через реку Большая Липовица")
            SelectOnly(form.CBox_ListNamesBridge,
                       "Путепровод на ПК 124+56,78 (левое направление)")
            SelectOnly(form.CBox_ListNamesTemplateXML,
                       "Монолитный участок сопряжения балок индивидуального пролёта")
            SelectOnly(form.CBox_ListProlet,
                       "Пролёт 3: опора 4 — опора 5, расчётная длина 33,0 м")
            form.RadioButton5.Checked = True
            form.RadioButton2.Checked = True
        End Sub

        Private Shared Sub PopulateLibraryPreview(form As FormUserTemptate)
            form.ListBox1.Items.AddRange(New Object() {
                "Типовая схема: мост 3×33 м с береговыми конусами",
                "Путепровод: индивидуальная раскладка балок над автомагистралью",
                "Монолитные участки и опоры для косого пересечения 72°"})
            form.ListBox1.SelectedIndex = 1
            form.TextBox1.Text = "Путепровод над М-12 — рабочая схема раскладки"
            form.Label3.Text = "D:\Проекты\Мосты\Библиотека шаблонов\Раскладки балок\2026"
        End Sub

        Private Shared Sub SelectOnly(combo As ComboBox, value As String)
            combo.Items.Clear()
            combo.Items.Add(value)
            combo.SelectedIndex = 0
        End Sub

        Private Shared Sub RenderPreview(form As Form,
                                         outputDirectory As String,
                                         fileName As String,
                                         requestedSize As Size)
            Dim targetSize As New Size(Math.Max(requestedSize.Width, form.MinimumSize.Width),
                                       Math.Max(requestedSize.Height, form.MinimumSize.Height))
            ShowOffscreen(form, targetSize)
            Dim previewPath As String = Path.Combine(outputDirectory, fileName)
            Using bitmap As New Bitmap(form.Width, form.Height)
                form.DrawToBitmap(bitmap, New Rectangle(0, 0, form.Width, form.Height))
                bitmap.Save(previewPath, Imaging.ImageFormat.Png)
            End Using
            Assert.That(File.Exists(previewPath), [Is].True)
            TestContext.Progress.WriteLine(previewPath)
        End Sub

        Private Shared Function WorkingRows(grid As DataGridView) As DataGridViewRow()
            Return grid.Rows.Cast(Of DataGridViewRow)().
                Where(Function(row) Not row.IsNewRow).ToArray()
        End Function

        Private Shared Sub ShowOffscreen(form As Form, targetSize As Size)
            form.ShowInTaskbar = False
            form.StartPosition = FormStartPosition.Manual
            form.Location = New Point(-30000, -30000)
            form.Size = targetSize
            If Not form.Visible Then form.Show()
            Application.DoEvents()
            PerformLayoutTree(form)
        End Sub

        Private Shared Sub PerformLayoutTree(parent As Control)
            For Each child As Control In parent.Controls
                PerformLayoutTree(child)
            Next
            parent.PerformLayout()
        End Sub

        Private Shared Sub AssertReadableInside(form As Form,
                                                control As Control,
                                                minimumWidth As Integer,
                                                minimumHeight As Integer)
            Assert.Multiple(
                Sub()
                    Assert.That(control.Visible, [Is].True, control.Name & " visible")
                    Assert.That(control.Width, [Is].GreaterThanOrEqualTo(minimumWidth), control.Name & " width")
                    Assert.That(control.Height, [Is].GreaterThanOrEqualTo(minimumHeight), control.Name & " height")
                    Assert.That(form.ClientRectangle.IntersectsWith(BoundsInForm(form, control)),
                                [Is].True, control.Name & " must remain reachable in the window.")
                End Sub)
        End Sub

        Private Shared Function BoundsInForm(form As Form, control As Control) As Rectangle
            Return form.RectangleToClient(control.Parent.RectangleToScreen(control.Bounds))
        End Function

        Private Shared Function ScreenBounds(control As Control) As Rectangle
            Return control.RectangleToScreen(control.ClientRectangle)
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
