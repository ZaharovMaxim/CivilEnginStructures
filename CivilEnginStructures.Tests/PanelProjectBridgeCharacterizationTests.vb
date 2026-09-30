Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class PanelProjectBridgeCharacterizationTests
        Private Shared ReadOnly MenuImageKeys As String() = {
            "001_Структура проекта.png", "002_Обновить структуру проекта.png", "003_Сооружение.png",
            "004_Перестроить сооружение.png", "005_Пролетные строения.png", "006_Оси сооружения.png",
            "007_Прочие элементы.png", "010_Мостовое полотно.png", "020_Опоры.png",
            "021_Промежуточные опоры.png", "022_Ось промежуточной опоры.png", "023_Сваи.png",
            "040_Балки.png", "040_Участки омоноличивания балок.png", "041_Представление по пролетам.png",
            "042_Представление по рядам.png", "060_Отчеты.png"
        }

        <Test>
        Public Sub DesignerInitializesExistingPaletteContract()
            Using panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Assert.Multiple(
                    Sub()
                        Assert.That(panel.CBox_ListNamesArrProject.DropDownStyle, [Is].EqualTo(ComboBoxStyle.DropDownList))
                        Assert.That(panel.CBox_ListNamesArrProject.SelectedIndex, [Is].EqualTo(-1))
                        Assert.That(panel.CBox_ListNamesArrProject.DataSource, [Is].Null)
                        Assert.That(panel.PropertyGrid1.SelectedObject, [Is].Null)
                        Assert.That(panel.PropertyGrid1.PropertySort, [Is].EqualTo(PropertySort.NoSort))
                        Assert.That(panel.TreeView1.SelectedNode, [Is].Null)
                        Assert.That(panel.TreeView1.Nodes, [Is].Empty)
                        Assert.That(panel.TreeView1.ImageList, [Is].SameAs(panel.ImageListTree))
                        Assert.That(panel.TreeView1.ContextMenuStrip, [Is].SameAs(panel.MenuSelectedNodeTree))
                    End Sub)

                AssertMenu(panel.ReportMenu,
                           panel.ТочкиОпиранияБалокToolStripMenuItem,
                           panel.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem,
                           panel.ОсиОпорToolStripMenuItem,
                           panel.ДеформационныеЗазорыToolStripMenuItem,
                           panel.ЭкспортЭлементовВDwgToolStripMenuItem)
                AssertMenu(panel.PillarsMenu,
                           panel.СоздатьОсьОпорыToolStripMenuItem,
                           panel.УдалитьВсеОпорыToolStripMenuItem,
                           panel.УдалитьКрайнююОпоруToolStripMenuItem)
                AssertMenu(panel.MenuBeams,
                           panel.СоздатьБалкуToolStripMenuItem,
                           panel.ВосстановитьЭлементыБалкиToolStripMenuItem,
                           panel.УдалитьToolStripMenuItem,
                           panel.ПереместитьБалкуВдольОсиToolStripMenuItem,
                           panel.ПереместитьБалкуВдольОсиОпиранияToolStripMenuItem,
                           panel.Razd1,
                           panel.ПоднятьОпуститьРядБалокToolStripMenuItem,
                           panel.СместитьРядБалокВдольОсейОпиранияToolStripMenuItem)
                AssertMenu(panel.UpdateStructureMenu, panel.ToolStripMenuItem1, panel.ToolStripMenuItem2)
                AssertMenu(panel.MenuSelectedNodeTree,
                           panel.ПоказатьToolStripMenuItem,
                           panel.ВыбратьЭлементToolStripMenuItem,
                           panel.ПоказатьНаПоперечникеToolStripMenuItem,
                           panel.ПоказатьВсеНаПоперечникеToolStripMenuItem,
                           panel.ОбновитьToolStripMenuItem)
                AssertMenu(panel.BridgeMenu, panel.UpdateBridgeToolStripMenuItem, panel.DeleteBridgeToolStripMenuItem)

                Assert.Multiple(
                    Sub()
                        AssertButtonMapping(panel.Button1, panel.ImageListMenu, "002_Обновить структуру проекта.png", panel.UpdateStructureMenu)
                        AssertButtonMapping(panel.Button2, panel.ImageListMenu, "004_Перестроить сооружение.png", Nothing)
                        AssertButtonMapping(panel.Button4, panel.ImageListMenu, "040_Балки.png", panel.MenuBeams)
                        AssertButtonMapping(panel.Button6, panel.ImageListMenu, "020_Опоры.png", panel.PillarsMenu)
                        AssertButtonMapping(panel.Button3, panel.ImageListMenu, "060_Отчеты.png", panel.ReportMenu)
                    End Sub)
                AssertImageKeys(panel.ImageListMenu, MenuImageKeys)
                AssertImageKeys(panel.ImageListTree, CopyWithoutLast(MenuImageKeys))
                AssertOriginalNamesAndTags(panel)
            End Using
        End Sub

        <Test>
        Public Sub AppearancePreservesFunctionalStateAndDesignerInstances()
            Using panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the palette.")

                Dim nodeTag As New Object()
                Dim selectedNode As New TreeNode("Seed node") With {.Name = "SeedNode", .Tag = nodeTag}
                panel.TreeView1.Nodes.Add(selectedNode)
                panel.TreeView1.SelectedNode = selectedNode
                Dim selectedObject As New SeedPropertyObject With {.Value = "Seed property"}
                panel.PropertyGrid1.SelectedObject = selectedObject
                Dim projects As New List(Of String) From {"First", "Selected"}
                panel.CBox_ListNamesArrProject.DataSource = projects
                panel.CBox_ListNamesArrProject.SelectedIndex = 1

                Dim originals As Dictionary(Of String, Object) = CaptureDesignerInstances(panel)
                Dim before As String() = CaptureFunctionalState(panel, originals.Keys, includeTreeStructure:=False)
                InvokeAppearance(method, panel)

                Assert.That(CaptureFunctionalState(panel, originals.Keys, includeTreeStructure:=False), [Is].EqualTo(before))
                Assert.Multiple(
                    Sub()
                        Assert.That(panel.TreeView1.SelectedNode, [Is].SameAs(selectedNode))
                        Assert.That(selectedNode.Name, [Is].EqualTo("SeedNode"))
                        Assert.That(selectedNode.Tag, [Is].SameAs(nodeTag))
                        Assert.That(panel.PropertyGrid1.SelectedObject, [Is].SameAs(selectedObject))
                        Assert.That(panel.CBox_ListNamesArrProject.DataSource, [Is].SameAs(projects))
                        Assert.That(panel.CBox_ListNamesArrProject.SelectedIndex, [Is].EqualTo(1))
                        Assert.That(panel.CBox_ListNamesArrProject.SelectedItem, [Is].EqualTo("Selected"))
                    End Sub)
                AssertDesignerInstancesArePreserved(panel, originals)
                AssertOriginalControlsRemainReachable(panel, originals)
            End Using
        End Sub

        <Test>
        Public Sub AppearanceCanBeAppliedTwiceWithoutFurtherChanges()
            Using panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the palette.")

                InvokeAppearance(method, panel)
                Dim instances As Dictionary(Of String, Object) = CaptureDesignerInstances(panel)
                Dim afterFirst As String() = CaptureFunctionalState(panel, instances.Keys, includeTreeStructure:=True)
                InvokeAppearance(method, panel)

                Assert.That(CaptureFunctionalState(panel, instances.Keys, includeTreeStructure:=True), [Is].EqualTo(afterFirst))
                AssertDesignerInstancesArePreserved(panel, instances)
            End Using
        End Sub

        <TestCase(456, 740)>
        <TestCase(340, 620)>
        <TestCase(300, 620)>
        <TestCase(456, 420)>
        Public Sub AppearanceLayoutKeepsPaletteContentAccessible(width As Integer, height As Integer)
            Using panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Dim method As MethodInfo = FindAppearanceMethod()
                Assert.That(method, [Is].Not.Null, "ApplyModernAppearance must remain available to the palette.")
                InvokeAppearance(method, panel)
                panel.Size = New Drawing.Size(width, height)
                panel.CreateControl()
                PerformLayoutTree(panel)

                Dim split As SplitContainer = DirectCast(FindDescendant(panel, "ModernPaletteSplit"), SplitContainer)
                Dim scrollHost As Panel = DirectCast(FindDescendant(panel, "ModernPaletteScroll"), Panel)
                Dim canvas As Control = FindDescendant(panel, "ModernPaletteCanvas")
                Dim title As Label = DirectCast(FindDescendant(panel, "ModernPaletteTitle"), Label)
                Assert.Multiple(
                    Sub()
                        For Each button As Button In New Button() {panel.Button1, panel.Button2, panel.Button4, panel.Button6, panel.Button3}
                            AssertInsideParent(button)
                        Next
                        AssertInsideParent(panel.CBox_ListNamesArrProject)
                        AssertInsideParent(panel.TreeView1)
                        AssertInsideParent(panel.PropertyGrid1)
                        AssertInsideParent(split)
                        Assert.That(split.Panel1.ClientSize.Height, [Is].GreaterThan(0))
                        Assert.That(split.Panel2.ClientSize.Height, [Is].GreaterThan(0))
                        Assert.That(scrollHost.AutoScroll, [Is].True)
                        Assert.That(scrollHost.DisplayRectangle.Bottom, [Is].GreaterThanOrEqualTo(canvas.Bottom))
                        Assert.That(title.ClientSize.Width,
                                    [Is].GreaterThanOrEqualTo(title.GetPreferredSize(Drawing.Size.Empty).Width),
                                    "The one-line palette title must fit without clipping.")
                        Dim preferredLabel As Drawing.Size = panel.Label1.GetPreferredSize(New Drawing.Size(panel.Label1.ClientSize.Width, 0))
                        Assert.That(panel.Label1.ClientSize.Height, [Is].GreaterThanOrEqualTo(preferredLabel.Height))
                    End Sub)

                Dim available As Integer = split.Height - split.SplitterWidth
                Dim lower As Integer = split.Panel1MinSize
                Dim upper As Integer = available - split.Panel2MinSize
                Assert.That(upper, [Is].GreaterThan(lower), "The splitter must have a movable range.")
                Dim target As Integer = If(split.SplitterDistance = lower, upper, lower)
                split.SplitterDistance = target
                Assert.That(split.SplitterDistance, [Is].InRange(lower, upper))
            End Using
        End Sub

        <Test>
        Public Sub ShownPaletteSequentialResizeUsesOnlyRequiredScrollBars()
            Using host As New Form(), panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                host.ShowInTaskbar = False
                host.StartPosition = FormStartPosition.Manual
                host.Location = New Drawing.Point(-32000, -32000)
                host.AutoScaleMode = AutoScaleMode.None
                host.ClientSize = New Drawing.Size(456, 740)

                Dim rootTag As New Object()
                Dim childTag As New Object()
                Dim root As New TreeNode("Very long populated bridge project root node") With {
                    .Name = "PopulatedRoot",
                    .Tag = rootTag
                }
                Dim child As New TreeNode("Very long nested structure element used during sequential resize") With {
                    .Name = "PopulatedChild",
                    .Tag = childTag
                }
                root.Nodes.Add(child)
                panel.TreeView1.Nodes.Add(root)
                panel.TreeView1.SelectedNode = child
                panel.PropertyGrid1.SelectedObject = New SeedPropertyObject With {.Value = "Populated property"}

                panel.Dock = DockStyle.Fill
                host.Controls.Add(panel)
                InvokeAppearance(FindAppearanceMethod(), panel)
                host.Show()
                Application.DoEvents()

                Dim viewports As Drawing.Size() = {
                    New Drawing.Size(456, 740),
                    New Drawing.Size(340, 620),
                    New Drawing.Size(456, 420),
                    New Drawing.Size(300, 620),
                    New Drawing.Size(456, 740)
                }
                Dim expectedHorizontal As Boolean() = {False, False, False, True, False}
                Dim expectedVertical As Boolean() = {False, False, True, False, False}
                Dim scrollHost As Panel = DirectCast(FindDescendant(panel, "ModernPaletteScroll"), Panel)
                Dim canvas As Control = FindDescendant(panel, "ModernPaletteCanvas")

                For index As Integer = 0 To viewports.Length - 1
                    host.ClientSize = viewports(index)
                    host.PerformLayout()
                    Application.DoEvents()
                    Dim currentIndex As Integer = index

                    Assert.Multiple(
                        Sub()
                            Assert.That(scrollHost.HorizontalScroll.Visible, [Is].EqualTo(expectedHorizontal(currentIndex)),
                                        "horizontal scrollbar at " & viewports(currentIndex).ToString())
                            Assert.That(scrollHost.VerticalScroll.Visible, [Is].EqualTo(expectedVertical(currentIndex)),
                                        "vertical scrollbar at " & viewports(currentIndex).ToString())
                            If Not expectedHorizontal(currentIndex) AndAlso Not expectedVertical(currentIndex) Then
                                Assert.That(canvas.Location, [Is].EqualTo(Drawing.Point.Empty),
                                            "canvas origin after overflow clears")
                            End If
                        End Sub)
                Next
            End Using
        End Sub

        <Test>
        Public Sub ShownPaletteRestoresSemanticAppearanceAfterHostThemeMutation()
            Using host As Form = CreateOffscreenHost(New Drawing.Size(600, 740)),
                  panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Dim nodeTag As New Object()
                Dim selectedNode As New TreeNode("Selected bridge element") With {.Name = "SelectedNode", .Tag = nodeTag}
                Dim selectedObject As New SeedPropertyObject With {.Value = "Selected property"}
                Dim projects As New List(Of String) From {"First", "Selected"}
                panel.TreeView1.Nodes.Add(selectedNode)
                panel.TreeView1.SelectedNode = selectedNode
                panel.PropertyGrid1.SelectedObject = selectedObject
                panel.CBox_ListNamesArrProject.DataSource = projects
                panel.CBox_ListNamesArrProject.SelectedIndex = 1

                panel.Dock = DockStyle.Fill
                host.Controls.Add(panel)
                InvokeAppearance(FindAppearanceMethod(), panel)
                host.Show()
                Application.DoEvents()

                ApplySimulatedHostTheme(panel)
                panel.Visible = False
                host.ClientSize = New Drawing.Size(601, 741)
                panel.Visible = True
                host.PerformLayout()
                Application.DoEvents()

                Dim canvas As Control = FindDescendant(panel, "ModernPaletteCanvas")
                Dim header As Control = FindDescendant(panel, "ModernPaletteHeader")
                Dim selectorCard As Control = FindDescendant(panel, "ModernModelSelector")
                Dim title As Label = DirectCast(FindDescendant(panel, "ModernPaletteTitle"), Label)
                Dim accent As Panel = FindHeaderAccent(header)
                Dim white As Drawing.Color = Drawing.Color.White
                Assert.Multiple(
                    Sub()
                        Assert.That(canvas.BackColor, [Is].EqualTo(Drawing.Color.FromArgb(243, 246, 250)))
                        Assert.That(header.BackColor, [Is].EqualTo(white))
                        Assert.That(selectorCard.BackColor, [Is].EqualTo(white))
                        Assert.That(panel.TreeView1.BackColor, [Is].EqualTo(white))
                        Assert.That(panel.CBox_ListNamesArrProject.BackColor, [Is].EqualTo(white))
                        Assert.That(panel.TreeView1.Parent.BackColor, [Is].EqualTo(white))
                        Assert.That(panel.PropertyGrid1.Parent.BackColor, [Is].EqualTo(white))
                        Assert.That(accent.BackColor, [Is].EqualTo(Drawing.Color.FromArgb(37, 99, 235)))
                        Assert.That(title.ForeColor, [Is].EqualTo(Drawing.Color.FromArgb(32, 50, 77)))
                        Assert.That(panel.TreeView1.SelectedNode, [Is].SameAs(selectedNode))
                        Assert.That(selectedNode.Tag, [Is].SameAs(nodeTag))
                        Assert.That(panel.PropertyGrid1.SelectedObject, [Is].SameAs(selectedObject))
                        Assert.That(panel.CBox_ListNamesArrProject.DataSource, [Is].SameAs(projects))
                        Assert.That(panel.CBox_ListNamesArrProject.SelectedIndex, [Is].EqualTo(1))
                    End Sub)
                For Each button As Button In PaletteButtons(panel)
                    Assert.Multiple(
                        Sub()
                            Assert.That(button.BackColor, [Is].EqualTo(white), button.Name & " background")
                            Assert.That(button.FlatStyle, [Is].EqualTo(FlatStyle.Flat), button.Name & " flat style")
                            Assert.That(button.UseVisualStyleBackColor, [Is].False, button.Name & " visual style")
                            Assert.That(button.TextImageRelation, [Is].EqualTo(TextImageRelation.ImageAboveText), button.Name & " relation")
                            Assert.That(button.TextAlign, [Is].EqualTo(Drawing.ContentAlignment.BottomCenter), button.Name & " text alignment")
                            Assert.That(button.ImageAlign, [Is].EqualTo(Drawing.ContentAlignment.TopCenter), button.Name & " image alignment")
                        End Sub)
                Next
            End Using
        End Sub

        <Test>
        Public Sub ShownPaletteToolbarCaptionsFitAcrossHostWidthsAndFontChange()
            Using host As Form = CreateOffscreenHost(New Drawing.Size(600, 740)),
                  panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly(),
                  largerFont As New Drawing.Font(Drawing.SystemFonts.MessageBoxFont.FontFamily, 13.0!, Drawing.FontStyle.Regular)
                panel.Dock = DockStyle.Fill
                host.Controls.Add(panel)
                InvokeAppearance(FindAppearanceMethod(), panel)
                host.Show()
                Application.DoEvents()

                Dim buttons As Button() = PaletteButtons(panel)
                Dim accessibleNames As New Dictionary(Of Button, String)()
                For Each button As Button In buttons
                    accessibleNames.Add(button, button.AccessibleName)
                Next
                host.Font = largerFont
                For Each button As Button In buttons
                    button.Font = largerFont
                Next
                Application.DoEvents()
                For Each button As Button In buttons
                    Assert.That(button.Font.SizeInPoints, [Is].EqualTo(9.5!).Within(0.1!),
                                button.Name & " font must recover after host theme mutation.")
                Next

                For Each width As Integer In New Integer() {380, 400, 420, 460, 600}
                    host.ClientSize = New Drawing.Size(width, 740)
                    host.PerformLayout()
                    Application.DoEvents()

                    Dim captionCount As Integer = 0
                    For Each button As Button In buttons
                        If button.Text.Length > 0 Then captionCount += 1
                    Next
                    Assert.That(captionCount = 0 OrElse captionCount = buttons.Length, [Is].True,
                                "Toolbar captions must switch as one group at width " & width)

                    If captionCount = 0 Then
                        For Each button As Button In buttons
                            Assert.That(button.AccessibleName, [Is].EqualTo(accessibleNames(button)))
                        Next
                    Else
                        For Each button As Button In buttons
                            Dim preferred As Drawing.Size = button.GetPreferredSize(Drawing.Size.Empty)
                            Assert.That(button.ClientSize.Width, [Is].GreaterThanOrEqualTo(preferred.Width),
                                        button.Name & " caption width at host width " & width)
                            Assert.That(button.ClientSize.Height, [Is].GreaterThanOrEqualTo(preferred.Height),
                                        button.Name & " caption height at host width " & width)
                        Next
                    End If
                    If width = 600 Then Assert.That(captionCount, [Is].EqualTo(buttons.Length))
                Next
            End Using
        End Sub

        <Test>
        Public Sub ShownPaletteRecoversPendingStyleAfterHandleIsDestroyedAndRecreated()
            Using host As Form = CreateOffscreenHost(New Drawing.Size(456, 740)),
                  panel As PanelProjectBridge = PanelProjectBridge.CreateDesignerOnly()
                Dim nodeTag As New Object()
                Dim selectedNode As New TreeNode("Selected bridge element") With {.Name = "SelectedNode", .Tag = nodeTag}
                Dim selectedObject As New SeedPropertyObject With {.Value = "Selected property"}
                panel.TreeView1.Nodes.Add(selectedNode)
                panel.TreeView1.SelectedNode = selectedNode
                panel.PropertyGrid1.SelectedObject = selectedObject
                panel.Dock = DockStyle.Fill
                host.Controls.Add(panel)
                InvokeAppearance(FindAppearanceMethod(), panel)
                host.Show()
                Application.DoEvents()

                Dim destroyHandle As MethodInfo = GetType(Control).GetMethod(
                    "DestroyHandle", BindingFlags.Instance Or BindingFlags.NonPublic)
                Dim createHandle As MethodInfo = GetType(Control).GetMethod(
                    "CreateHandle", BindingFlags.Instance Or BindingFlags.NonPublic)
                Dim hostGray As Drawing.Color = Drawing.Color.FromArgb(243, 243, 243)
                Dim modernCanvas As Drawing.Color = Drawing.Color.FromArgb(243, 246, 250)

                panel.BackColor = hostGray
                destroyHandle.Invoke(panel, Nothing)
                Application.DoEvents()
                createHandle.Invoke(panel, Nothing)
                Application.DoEvents()

                Dim canvas As Control = FindDescendant(panel, "ModernPaletteCanvas")
                Assert.Multiple(
                    Sub()
                        Assert.That(panel.IsHandleCreated, [Is].True)
                        Assert.That(panel.BackColor, [Is].EqualTo(modernCanvas))
                        Assert.That(canvas.BackColor, [Is].EqualTo(modernCanvas))
                        Assert.That(panel.TreeView1.SelectedNode, [Is].SameAs(selectedNode))
                        Assert.That(selectedNode.Tag, [Is].SameAs(nodeTag))
                        Assert.That(panel.PropertyGrid1.SelectedObject, [Is].SameAs(selectedObject))
                    End Sub)

                panel.BackColor = hostGray
                Application.DoEvents()
                Assert.That(panel.BackColor, [Is].EqualTo(modernCanvas),
                            "Style recovery must remain usable after handle recreation.")
            End Using
        End Sub

        Private Shared Sub AssertMenu(menu As ContextMenuStrip, ParamArray expectedItems As ToolStripItem())
            Assert.That(menu.Items.Count, [Is].EqualTo(expectedItems.Length), menu.Name & " item count")
            For index As Integer = 0 To expectedItems.Length - 1
                Assert.That(menu.Items(index), [Is].SameAs(expectedItems(index)), menu.Name & " item " & index)
            Next
        End Sub

        Private Shared Sub AssertButtonMapping(button As Button,
                                               images As ImageList,
                                               imageKey As String,
                                               menu As ContextMenuStrip)
            Assert.That(button.ImageList, [Is].SameAs(images), button.Name & " image list")
            Assert.That(button.ImageKey, [Is].EqualTo(imageKey), button.Name & " image key")
            Assert.That(button.ContextMenuStrip, [Is].SameAs(menu), button.Name & " context menu")
        End Sub

        Private Shared Sub AssertImageKeys(images As ImageList, expectedKeys As String())
            Assert.That(images.Images.Count, [Is].EqualTo(expectedKeys.Length))
            For index As Integer = 0 To expectedKeys.Length - 1
                Assert.That(images.Images.Keys(index), [Is].EqualTo(expectedKeys(index)), "image key " & index)
            Next
        End Sub

        Private Shared Function CopyWithoutLast(values As String()) As String()
            Dim result(values.Length - 2) As String
            Array.Copy(values, result, result.Length)
            Return result
        End Function

        Private Shared Sub PerformLayoutTree(parent As Control)
            parent.PerformLayout()
            For Each child As Control In parent.Controls
                PerformLayoutTree(child)
            Next
            parent.PerformLayout()
        End Sub

        Private Shared Function CreateOffscreenHost(clientSize As Drawing.Size) As Form
            Return New Form With {
                .ShowInTaskbar = False,
                .StartPosition = FormStartPosition.Manual,
                .Location = New Drawing.Point(-32000, -32000),
                .AutoScaleMode = AutoScaleMode.None,
                .ClientSize = clientSize
            }
        End Function

        Private Shared Sub ApplySimulatedHostTheme(parent As Control)
            parent.BackColor = Drawing.Color.FromArgb(243, 243, 243)
            parent.ForeColor = Drawing.Color.FromArgb(33, 33, 33)
            parent.Font = Drawing.SystemFonts.MessageBoxFont
            Dim button As Button = TryCast(parent, Button)
            If button IsNot Nothing Then
                button.FlatStyle = FlatStyle.Standard
                button.UseVisualStyleBackColor = True
                button.TextImageRelation = TextImageRelation.Overlay
                button.TextAlign = Drawing.ContentAlignment.MiddleCenter
                button.ImageAlign = Drawing.ContentAlignment.MiddleCenter
            End If
            For Each child As Control In parent.Controls
                ApplySimulatedHostTheme(child)
            Next
        End Sub

        Private Shared Function FindHeaderAccent(header As Control) As Panel
            For Each child As Control In header.Controls
                Dim accent As Panel = TryCast(child, Panel)
                If accent IsNot Nothing AndAlso accent.Width = 5 Then Return accent
            Next
            Assert.Fail("Could not find the palette header accent panel.")
            Return Nothing
        End Function

        Private Shared Function PaletteButtons(panel As PanelProjectBridge) As Button()
            Return New Button() {panel.Button1, panel.Button2, panel.Button4, panel.Button6, panel.Button3}
        End Function

        Private Shared Sub AssertInsideParent(control As Control)
            Assert.That(control.Parent, [Is].Not.Null, control.Name & " parent")
            Assert.That(control.Width, [Is].GreaterThan(0), control.Name & " width")
            Assert.That(control.Height, [Is].GreaterThan(0), control.Name & " height")
            Assert.That(control.Parent.ClientRectangle.Contains(control.Bounds), [Is].True,
                        control.Name & " must fit inside " & control.Parent.Name & ".")
        End Sub

        Private Shared Function FindDescendant(parent As Control, name As String) As Control
            If String.Equals(parent.Name, name, StringComparison.Ordinal) Then Return parent
            For Each child As Control In parent.Controls
                Dim match As Control = FindDescendant(child, name)
                If match IsNot Nothing Then Return match
            Next
            Return Nothing
        End Function

        Private Shared Sub AssertOriginalNamesAndTags(panel As PanelProjectBridge)
            For Each field As FieldInfo In DesignerFields()
                Dim value As Object = field.GetValue(panel)
                Dim control As Control = TryCast(value, Control)
                If control IsNot Nothing Then
                    Dim logicalFieldName As String = field.Name.TrimStart("_"c)
                    Dim expectedName As String = logicalFieldName
                    If logicalFieldName = NameOf(panel.MenuBeams) Then expectedName = "ContextMenuStrip3"
                    If logicalFieldName = NameOf(panel.UpdateStructureMenu) Then expectedName = "MenuSelectedNodeTree"
                    Assert.That(control.Name, [Is].EqualTo(expectedName), field.Name & " Name")
                    Assert.That(control.Tag, [Is].Null, field.Name & " Tag")
                End If

                Dim item As ToolStripItem = TryCast(value, ToolStripItem)
                If item IsNot Nothing Then
                    Assert.That(item.Name, [Is].EqualTo(field.Name.TrimStart("_"c)), field.Name & " Name")
                    Assert.That(item.Tag, [Is].Null, field.Name & " Tag")
                End If
            Next
        End Sub

        Private Shared Function FindAppearanceMethod() As MethodInfo
            Return GetType(PanelProjectBridge).GetMethod("ApplyModernAppearance",
                                                        BindingFlags.Instance Or BindingFlags.NonPublic,
                                                        Nothing,
                                                        Type.EmptyTypes,
                                                        Nothing)
        End Function

        Private Shared Sub InvokeAppearance(method As MethodInfo, panel As PanelProjectBridge)
            Try
                method.Invoke(panel, Nothing)
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then Throw exception.InnerException
                Throw
            End Try
        End Sub

        Private Shared Function CaptureDesignerInstances(panel As PanelProjectBridge) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)(StringComparer.Ordinal)
            For Each field As FieldInfo In DesignerFields()
                Dim value As Object = field.GetValue(panel)
                If value IsNot Nothing AndAlso TypeOf value Is Component Then result.Add(field.Name, value)
            Next
            Return result
        End Function

        Private Shared Sub AssertDesignerInstancesArePreserved(panel As PanelProjectBridge,
                                                                originals As Dictionary(Of String, Object))
            For Each entry As KeyValuePair(Of String, Object) In originals
                Dim field As FieldInfo = GetType(PanelProjectBridge).GetField(
                    entry.Key, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                Assert.That(field.GetValue(panel), [Is].SameAs(entry.Value), entry.Key & " instance")
            Next
        End Sub

        Private Shared Sub AssertOriginalControlsRemainReachable(panel As PanelProjectBridge,
                                                                 originals As Dictionary(Of String, Object))
            Dim reachable As New HashSet(Of Control)(ReferenceEqualityComparer(Of Control).Instance)
            CollectControls(panel, reachable)
            For Each entry As KeyValuePair(Of String, Object) In originals
                Dim control As Control = TryCast(entry.Value, Control)
                If control IsNot Nothing AndAlso Not TypeOf control Is ContextMenuStrip Then
                    Assert.That(reachable.Contains(control), [Is].True, entry.Key & " reachability")
                End If
            Next
        End Sub

        Private Shared Sub CollectControls(parent As Control, result As HashSet(Of Control))
            result.Add(parent)
            For Each child As Control In parent.Controls
                CollectControls(child, result)
            Next
        End Sub

        Private Shared Function CaptureFunctionalState(panel As PanelProjectBridge,
                                                        fieldNames As IEnumerable(Of String),
                                                        includeTreeStructure As Boolean) As String()
            Dim state As New List(Of String) From {
                "Combo=" & panel.CBox_ListNamesArrProject.DropDownStyle & ";" & panel.CBox_ListNamesArrProject.SelectedIndex & ";" & Identity(panel.CBox_ListNamesArrProject.DataSource),
                "Property=" & Identity(panel.PropertyGrid1.SelectedObject) & ";" & panel.PropertyGrid1.PropertySort,
                "Tree=" & Identity(panel.TreeView1.SelectedNode) & ";" & panel.TreeView1.Nodes.Count & ";" & panel.TreeView1.ImageIndex & ";" & panel.TreeView1.SelectedImageIndex,
                "TreeImages=" & Identity(panel.TreeView1.ImageList),
                "TreeMenu=" & Identity(panel.TreeView1.ContextMenuStrip)
            }

            For Each fieldName As String In fieldNames
                Dim field As FieldInfo = GetType(PanelProjectBridge).GetField(
                    fieldName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                Dim value As Object = field.GetValue(panel)
                Dim control As Control = TryCast(value, Control)
                If control IsNot Nothing Then state.Add(fieldName & "|Control|" & control.Name & "|" & Identity(control.Tag))
                Dim item As ToolStripItem = TryCast(value, ToolStripItem)
                If item IsNot Nothing Then
                    state.Add(fieldName & "|Item|" & item.Name & "|" & Identity(item.Tag) & "|" & item.Enabled)
                End If
            Next

            AppendMenuState(state, panel.ReportMenu)
            AppendMenuState(state, panel.PillarsMenu)
            AppendMenuState(state, panel.MenuBeams)
            AppendMenuState(state, panel.UpdateStructureMenu)
            AppendMenuState(state, panel.MenuSelectedNodeTree)
            AppendMenuState(state, panel.BridgeMenu)
            AppendImageState(state, "MenuImages", panel.ImageListMenu)
            AppendImageState(state, "TreeImages", panel.ImageListTree)
            If includeTreeStructure Then AppendControlTree(state, panel, "")
            state.Sort(StringComparer.Ordinal)
            Return state.ToArray()
        End Function

        Private Shared Sub AppendMenuState(state As List(Of String), menu As ContextMenuStrip)
            For index As Integer = 0 To menu.Items.Count - 1
                state.Add(menu.Name & "[" & index & "]=" & Identity(menu.Items(index)))
            Next
        End Sub

        Private Shared Sub AppendImageState(state As List(Of String), prefix As String, images As ImageList)
            For index As Integer = 0 To images.Images.Count - 1
                state.Add(prefix & "[" & index & "]=" & images.Images.Keys(index))
            Next
        End Sub

        Private Shared Sub AppendControlTree(state As List(Of String), parent As Control, path As String)
            Dim currentPath As String = path & "/" & parent.GetType().Name & ":" & parent.Name
            state.Add("ControlTree=" & currentPath & ";Children=" & parent.Controls.Count)
            For Each child As Control In parent.Controls
                AppendControlTree(state, child, currentPath)
            Next
        End Sub

        Private Shared Function Identity(value As Object) As String
            If value Is Nothing Then Return "Nothing"
            Return value.GetType().FullName & "#" & RuntimeHelpers.GetHashCode(value)
        End Function

        Private Shared Function DesignerFields() As FieldInfo()
            Return GetType(PanelProjectBridge).GetFields(
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

        Private NotInheritable Class SeedPropertyObject
            Public Property Value As String
        End Class
    End Class
End Namespace
