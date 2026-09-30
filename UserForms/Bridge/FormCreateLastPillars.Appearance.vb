Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.IO

Partial Public Class FormCreateLastPillars
    Private modernAppearanceApplied As Boolean
    Private modernLayoutInProgress As Boolean
    Private modernRegularFont As Font
    Private modernBoldFont As Font
    Private modernTitleFont As Font
    Private modernHeader As Panel
    Private modernFooter As Panel
    Private modernParametersCaption As Label
    Private modernPreviewCaption As Label
    Private modernLeftPreviewCaption As Label
    Private modernRightPreviewCaption As Label
    Private modernPreviewZoom As PreviewImageZoom
    Private modernPileLayoutTab As TabPage
    Private modernPileLayoutPreview As PictureBox
    Private modernPileLayoutGeometry As PileLayoutGeometry
    Private modernPileLayoutBitmap As Bitmap
    Private modernPileLayoutBitmapSize As Size
    Private modernParameterSchemeButton As Button
    Private modernParameterSchemePanel As Panel
    Private modernParameterSchemeClose As Button
    Private modernParameterSchemeImage As PictureBox
    Private modernParameterSchemeTitle As Label
    Private modernParameterSchemeOwnedImage As Image
    Private ReadOnly modernParameterSchemePaths As New Dictionary(Of TabPage, String)()

    Private Shared ReadOnly ModernCanvasColor As Color = Color.FromArgb(243, 246, 250)
    Private Shared ReadOnly ModernCardColor As Color = Color.White
    Private Shared ReadOnly ModernTextColor As Color = Color.FromArgb(32, 50, 77)
    Private Shared ReadOnly ModernMutedTextColor As Color = Color.FromArgb(93, 109, 133)
    Private Shared ReadOnly ModernAccentColor As Color = Color.FromArgb(37, 99, 235)
    Private Shared ReadOnly ModernAccentPaleColor As Color = Color.FromArgb(232, 240, 254)
    Private Shared ReadOnly ModernBorderColor As Color = Color.FromArgb(218, 226, 237)

    Private Sub ApplyModernAppearance()
        If modernAppearanceApplied Then Return
        modernAppearanceApplied = True

        SuspendLayout()
        Try
            modernRegularFont = New Font("Segoe UI", 9.5!, FontStyle.Regular, GraphicsUnit.Point)
            modernBoldFont = New Font("Segoe UI Semibold", 9.5!, FontStyle.Bold, GraphicsUnit.Point)
            modernTitleFont = New Font("Segoe UI Semibold", 17.0!, FontStyle.Bold, GraphicsUnit.Point)
            ConfigureModernWindow()
            CreateModernChrome()
            ConfigureProjectControls()
            ConfigureParameterTabs()
            ConfigurePreviewTabs()
            ConfigureActionButtons()
            ConfigureModernWindowSize()
            AddHandler SizeChanged, AddressOf ModernSizeChanged
            AddHandler Shown, AddressOf ModernShown
            AddHandler TabControl1.SizeChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl2.SizeChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl3.SizeChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl4.SizeChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl1.SelectedIndexChanged, AddressOf ModernParameterTabChanged
            AddHandler TabControl1.SelectedIndexChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl2.SelectedIndexChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl3.SelectedIndexChanged, AddressOf ModernTabLayoutChanged
            AddHandler TabControl4.SelectedIndexChanged, AddressOf ModernTabLayoutChanged
            For Each page As TabPage In New TabPage() {TabPage1, TabPage16, TabPage8, TabPage7,
                                                        TabPage9, TabPage10, TabPage6, TabPage11,
                                                        TabPage12, TabPage2, TabPage4, TabPage3,
                                                        TabPage5, TabPage13, TabPage14,
                                                        modernPileLayoutTab}
                AddHandler page.ClientSizeChanged, AddressOf ModernTabLayoutChanged
            Next
            AddHandler Disposed, AddressOf DisposeModernAppearanceResources
            LayoutModernAppearance()
        Finally
            ResumeLayout(True)
        End Try
    End Sub

    Private Sub ConfigureModernWindow()
        AutoScaleMode = AutoScaleMode.Dpi
        AutoScaleDimensions = New SizeF(DeviceDpi, DeviceDpi)
        Font = modernRegularFont
        BackColor = ModernCanvasColor
        ForeColor = ModernTextColor
        DoubleBuffered = True
        MaximumSize = Size.Empty
        Padding = Padding.Empty
        AutoScroll = True
        AutoScrollMinSize = ModernScaleSize(New Size(760, 600))
        StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConfigureModernWindowSize()
        Dim workArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim safeWidth As Integer = Math.Max(1, workArea.Width - ModernScale(32))
        Dim safeHeight As Integer = Math.Max(1, workArea.Height - ModernScale(32))
        Dim preferred As Size = ModernScaleSize(New Size(1360, 960))
        Dim requestedMinimum As Size = ModernScaleSize(New Size(1120, 760))
        MinimumSize = New Size(Math.Min(requestedMinimum.Width, safeWidth),
                               Math.Min(requestedMinimum.Height, safeHeight))
        Size = New Size(Math.Min(preferred.Width, safeWidth),
                        Math.Min(preferred.Height, safeHeight))
    End Sub

    Private Sub CreateModernChrome()
        modernHeader = New Panel() With {.Name = "ModernHeader", .BackColor = ModernCardColor}
        Dim accent As New Panel() With {.Name = "ModernHeaderAccent", .BackColor = ModernAccentColor}
        Dim title As New Label() With {
            .Name = "ModernTitle", .AutoSize = False, .Font = modernTitleFont,
            .ForeColor = ModernTextColor, .Text = "Крайняя опора",
            .TextAlign = ContentAlignment.BottomLeft
        }
        Dim subtitle As New Label() With {
            .Name = "ModernSubtitle", .AutoSize = False, .Font = modernRegularFont,
            .ForeColor = ModernMutedTextColor,
            .Text = "Настройте элементы опоры и проверьте виды перед размещением",
            .TextAlign = ContentAlignment.TopLeft
        }
        modernHeader.Controls.Add(accent)
        modernHeader.Controls.Add(title)
        modernHeader.Controls.Add(subtitle)
        modernFooter = New Panel() With {.Name = "ModernFooter", .BackColor = ModernCardColor}
        modernParametersCaption = CreateSectionCaption("ModernParametersCaption", "Параметры элементов")
        modernPreviewCaption = CreateSectionCaption("ModernPreviewCaption", "Предпросмотр · масштаб колёсиком")
        modernParameterSchemeButton = New Button() With {
            .Name = "ModernParameterSchemeButton",
            .Text = "Схема элемента"
        }
        StyleSecondaryButton(modernParameterSchemeButton)
        modernParameterSchemePanel = New Panel() With {
            .Name = "ModernParameterSchemePanel",
            .BackColor = ModernCardColor,
            .BorderStyle = BorderStyle.FixedSingle,
            .Visible = False
        }
        modernParameterSchemeTitle = CreateSectionCaption("ModernParameterSchemeTitle", "Схема элемента")
        modernParameterSchemeClose = New Button() With {
            .Name = "ModernParameterSchemeClose",
            .Text = "×"
        }
        StyleSecondaryButton(modernParameterSchemeClose)
        modernParameterSchemeImage = New PictureBox() With {
            .Name = "ModernParameterSchemeImage",
            .BackColor = ModernCardColor,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
        modernParameterSchemePanel.Controls.Add(modernParameterSchemeTitle)
        modernParameterSchemePanel.Controls.Add(modernParameterSchemeClose)
        modernParameterSchemePanel.Controls.Add(modernParameterSchemeImage)
        Controls.Add(modernHeader)
        Controls.Add(modernFooter)
        Controls.Add(modernParametersCaption)
        Controls.Add(modernPreviewCaption)
        Controls.Add(modernParameterSchemeButton)
        Controls.Add(modernParameterSchemePanel)
        modernHeader.SendToBack()
        modernFooter.SendToBack()
        AddHandler modernParameterSchemeButton.Click, AddressOf ModernParameterSchemeButtonClick
        AddHandler modernParameterSchemeClose.Click, AddressOf ModernParameterSchemeCloseClick
    End Sub

    Private Function CreateSectionCaption(name As String, caption As String) As Label
        Return New Label() With {
            .Name = name, .AutoSize = False, .Font = modernBoldFont,
            .ForeColor = ModernTextColor, .Text = caption,
            .TextAlign = ContentAlignment.MiddleLeft
        }
    End Function

    Private Sub ConfigureProjectControls()
        GroupBox1.Text = "Настройки проекта"
        GroupBox1.Font = modernBoldFont
        GroupBox1.ForeColor = ModernTextColor
        GroupBox1.BackColor = ModernCardColor
        Label12.Text = "Модель проекта"
        Label15.Text = "Сооружение"
        Label14.Text = "Ось трассы"
        Label2.Text = "Шаблон оформления"
        Label13.Text = "Проектная поверхность"
        Label16.Text = "Поверхность земли"
        Label26.Text = "Номер опоры"
        Label47.Text = "Схема опоры"
        Label17.Visible = False
        For Each fieldLabel As Label In New Label() {Label12, Label15, Label14, Label2,
                                                       Label13, Label16, Label26, Label47}
            StyleFieldLabel(fieldLabel)
        Next
        For Each input As Control In New Control() {CBox_ListNamesArrProject, CBox_ListNamesBridge,
                                                     CB_NameAlignment, CBox_ListNamesTemplateXML,
                                                     CB_ProjectSurface, CB_EgSurface, CB_NumberPillar,
                                                     ComboBox11, NUpD_ElevationLand}
            StyleInput(input)
        Next
        StyleCheckBox(ChB_ProjectSurfaceInAlignment)
        StyleCheckBox(ChB_ElevationLand)
        StyleSecondaryButton(Button3)
    End Sub

    Private Sub ConfigureParameterTabs()
        StyleTabControl(TabControl1)
        StyleTabControl(TabControl2)
        StyleTabControl(TabControl3)
        TabControl1.Multiline = True
        For Each page As TabPage In New TabPage() {TabPage1, TabPage16, TabPage8, TabPage7,
                                                    TabPage9, TabPage10, TabPage6, TabPage11,
                                                    TabPage12, TabPage2, TabPage4, TabPage3,
                                                    TabPage5}
            page.BackColor = ModernCardColor
            page.ForeColor = ModernTextColor
            page.Padding = New Padding(ModernScale(10))
        Next
        TabPage1.Text = "Насадка"
        TabPage16.Text = "Подферменник"
        TabPage8.Text = "Шкафная стенка"
        TabPage7.Text = "Откосные крылья"
        TabPage6.Text = "Открылки"
        TabPage2.Text = "Стойки"
        TabPage4.Text = "Ростверк"
        TabPage3.Text = "Подготовка"
        TabPage5.Text = "Сваи"
        TabPage9.Text = "Левая сторона"
        TabPage10.Text = "Правая сторона"
        TabPage11.Text = "Левая сторона"
        TabPage12.Text = "Правая сторона"
        TabPage5.AutoScroll = True
        For Each input As Control In New Control() {CB_NozzleTLC, CB_CabinetWallTLC,
                                                     CB_LeftHandTLC, CB_RightHandTLC,
                                                     CB_LeftPostcardTLC, CB_RightPostcardTLC,
                                                     CB_RackTLC, CB_GrillageTLC,
                                                     CB_PreparationTLC, CB_PileTLC, NUpD_CountRack,
                                                     NUpD_CountRowsPile, NUpD_CountColumnsPile,
                                                     NUpD_OffsetRowsPile, NUpD_OffsetColumnsPile,
                                                     TxtB_PileRowDiagram, TxtB_PileCollDiagram}
            StyleInput(input)
        Next
        For Each checkBox As CheckBox In New CheckBox() {ChB_FixedHeightCabinetWall,
                                                          ChB_CalculateVerticalLineLeftHand,
                                                          ChB_CalculateVerticalLineRightHand,
                                                          CheckBox1, ChB_LeftPostcsrdFixedLenght,
                                                          ChB_RightPostcsrdFixedLenght, CheckBox2,
                                                          ChB_CreateRack, ChB_fixedHeightRack,
                                                          ChB_EgeParallel, ChB_CreateGrillage,
                                                          ChB_GrillageEge, ChB_CreatePreparation,
                                                          ChB_InsertPileInRack, ChB_PileExpand}
            StyleCheckBox(checkBox)
        Next
        For Each fieldLabel As Label In New Label() {Label1, Label3, Label23, Label24, Label25,
                                                       Label5, Label4, Label6, Label7, Label8,
                                                       Label9, Label10, Label11, Label18, Label19,
                                                       Label20, Label21, Label22, Label39, Label40}
            StyleFieldLabel(fieldLabel)
        Next
        For Each grid As DataGridView In ParameterGrids()
            StyleGrid(grid)
        Next
        For Each button As Button In New Button() {Button1, Button4, Button5, Button6}
            StyleSecondaryButton(button)
        Next
        ChB_EgeParallel.Text = "Грани стоек параллельны граням насадки"
        Label20.Text = "Отношение длины сваи к её смещению"
        Label21.Visible = False
    End Sub

    Private Sub ConfigurePreviewTabs()
        StyleTabControl(TabControl4)
        For Each page As TabPage In New TabPage() {TabPage13, TabPage14}
            page.BackColor = ModernCardColor
            page.ForeColor = ModernTextColor
            page.Padding = New Padding(ModernScale(10))
        Next
        For Each legacyEditor As Control In New Control() {
            Label30, Label31, NUpD_ScaleFront, Label32, NUpD_dxFront, NUpD_dyFront,
            Label36, Label37, NUpD_ScaleRight, Label38, NUpD_dxRight, NUpD_dyRight,
            Label33, Label34, NUpD_ScaleLeft, Label35, NUpD_dxLeft, NUpD_dyLeft
        }
            legacyEditor.Visible = False
        Next
        For Each picture As PictureBox In New PictureBox() {PictureBox1, PictureBox2, PictureBox5}
            picture.BackColor = ModernCardColor
            picture.BorderStyle = BorderStyle.FixedSingle
            picture.SizeMode = PictureBoxSizeMode.Zoom
        Next
        modernLeftPreviewCaption = CreateSectionCaption("ModernLeftPreviewCaption", "Слева")
        modernRightPreviewCaption = CreateSectionCaption("ModernRightPreviewCaption", "Справа")
        modernLeftPreviewCaption.AutoSize = True
        modernRightPreviewCaption.AutoSize = True
        TabPage14.Controls.Add(modernLeftPreviewCaption)
        TabPage14.Controls.Add(modernRightPreviewCaption)
        modernPileLayoutTab = New TabPage() With {
            .Name = "ModernPileLayoutTab",
            .Text = "Раскладка свай",
            .BackColor = ModernCardColor,
            .ForeColor = ModernTextColor,
            .Padding = New Padding(ModernScale(10))
        }
        modernPileLayoutPreview = New PictureBox() With {
            .Name = "ModernPileLayoutPreview",
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .SizeMode = PictureBoxSizeMode.Zoom
        }
        modernPileLayoutTab.Controls.Add(modernPileLayoutPreview)
        TabControl4.TabPages.Add(modernPileLayoutTab)
        modernPileLayoutGeometry = PileLayoutGeometry.Create(Nothing, Nothing, Nothing)
        modernPreviewZoom = New PreviewImageZoom(Me, PictureBox1, PictureBox2, PictureBox5,
                                                  modernPileLayoutPreview)
    End Sub

    Private Sub ConfigureActionButtons()
        StyleSecondaryButton(Button2)
        StyleSecondaryButton(Button7)
        StylePrimaryButton(Button8)
        Button2.Text = "Отмена"
        Button7.Text = "Расчёт"
        Button8.Text = "Разместить опору"
    End Sub

    Private Sub StyleFieldLabel(fieldLabel As Label)
        fieldLabel.AutoSize = True
        fieldLabel.Font = modernRegularFont
        fieldLabel.ForeColor = ModernMutedTextColor
        fieldLabel.BackColor = Color.Transparent
    End Sub

    Private Sub StyleInput(input As Control)
        input.Font = modernRegularFont
        input.ForeColor = ModernTextColor
        input.BackColor = ModernCardColor
        If TypeOf input Is ComboBox Then DirectCast(input, ComboBox).FlatStyle = FlatStyle.Flat
        If TypeOf input Is TextBox Then DirectCast(input, TextBox).BorderStyle = BorderStyle.FixedSingle
    End Sub

    Private Sub StyleCheckBox(checkBox As CheckBox)
        checkBox.AutoSize = True
        checkBox.Font = modernRegularFont
        checkBox.ForeColor = ModernTextColor
        checkBox.BackColor = ModernCardColor
        checkBox.UseVisualStyleBackColor = False
    End Sub

    Private Sub StyleTabControl(tabs As TabControl)
        tabs.Font = modernRegularFont
        tabs.Padding = New Point(ModernScale(14), ModernScale(5))
    End Sub

    Private Sub StyleGrid(grid As DataGridView)
        grid.Font = modernRegularFont
        grid.BackgroundColor = ModernCardColor
        grid.BorderStyle = BorderStyle.FixedSingle
        grid.GridColor = ModernBorderColor
        grid.EnableHeadersVisualStyles = False
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        grid.ColumnHeadersDefaultCellStyle.BackColor = ModernAccentPaleColor
        grid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTextColor
        grid.ColumnHeadersDefaultCellStyle.Font = modernBoldFont
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTextColor
        grid.DefaultCellStyle.BackColor = ModernCardColor
        grid.DefaultCellStyle.ForeColor = ModernTextColor
        grid.DefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        grid.DefaultCellStyle.SelectionForeColor = ModernTextColor
    End Sub

    Private Sub StylePrimaryButton(button As Button)
        button.Font = modernBoldFont
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.UseVisualStyleBackColor = False
        AddHandler button.EnabledChanged, AddressOf PrimaryButtonEnabledChanged
        UpdatePrimaryButtonState(button)
    End Sub

    Private Sub PrimaryButtonEnabledChanged(sender As Object, e As EventArgs)
        UpdatePrimaryButtonState(DirectCast(sender, Button))
    End Sub

    Private Sub UpdatePrimaryButtonState(button As Button)
        button.BackColor = If(button.Enabled, ModernAccentColor, ModernAccentPaleColor)
        button.ForeColor = If(button.Enabled, Color.White, ModernMutedTextColor)
    End Sub

    Private Sub StyleSecondaryButton(button As Button)
        button.Font = modernBoldFont
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderColor = ModernBorderColor
        button.FlatAppearance.BorderSize = 1
        button.BackColor = ModernCardColor
        button.ForeColor = ModernTextColor
        button.UseVisualStyleBackColor = False
    End Sub

    Private Sub ModernSizeChanged(sender As Object, e As EventArgs)
        LayoutModernAppearance()
    End Sub

    Private Sub ModernShown(sender As Object, e As EventArgs)
        BeginInvoke(New MethodInvoker(AddressOf LayoutModernAppearance))
    End Sub

    Private Sub ModernTabLayoutChanged(sender As Object, e As EventArgs)
        If modernLayoutInProgress Then Return
        LayoutModernAppearance()
    End Sub

    Private Sub ModernParameterTabChanged(sender As Object, e As EventArgs)
        HideParameterScheme()
        UpdateParameterSchemeButton()
    End Sub

    Private Sub LayoutModernAppearance()
        If Not modernAppearanceApplied OrElse modernLayoutInProgress OrElse modernHeader Is Nothing Then Return
        Dim rootScrollOffset As New Point(Math.Max(0, -AutoScrollPosition.X),
                                          Math.Max(0, -AutoScrollPosition.Y))
        modernLayoutInProgress = True
        SuspendLayout()
        Try
            If rootScrollOffset <> Point.Empty Then AutoScrollPosition = Point.Empty
            Dim margin As Integer = ModernScale(16)
            Dim gap As Integer = ModernScale(12)
            Dim clientWidth As Integer = Math.Max(ModernScale(760), ClientSize.Width)
            Dim clientHeight As Integer = Math.Max(ModernScale(600), ClientSize.Height)
            modernHeader.SetBounds(margin, ModernScale(12), clientWidth - margin * 2, ModernScale(72))
            LayoutModernHeader()
            GroupBox1.SetBounds(margin, modernHeader.Bottom + gap, clientWidth - margin * 2, ModernScale(158))
            LayoutProjectControls()
            Dim footerTop As Integer = clientHeight - ModernScale(68)
            modernFooter.SetBounds(0, footerTop, clientWidth, ModernScale(68))
            LayoutFooterActions()
            Dim captionTop As Integer = GroupBox1.Bottom + ModernScale(8)
            Dim bodyTop As Integer = captionTop + ModernScale(28)
            Dim bodyHeight As Integer = Math.Max(ModernScale(190), footerTop - ModernScale(12) - bodyTop)
            Dim availableWidth As Integer = clientWidth - margin * 2
            Dim leftWidth As Integer = CInt(Math.Floor((availableWidth - gap) * 0.52R))
            Dim rightWidth As Integer = availableWidth - gap - leftWidth
            Dim schemeButtonWidth As Integer = Math.Min(ModernScale(144),
                                                        Math.Max(ModernScale(112), leftWidth \ 3))
            modernParametersCaption.SetBounds(margin, captionTop,
                                              Math.Max(ModernScale(120), leftWidth - schemeButtonWidth - ModernScale(8)),
                                              ModernScale(24))
            modernParameterSchemeButton.SetBounds(margin + leftWidth - schemeButtonWidth, captionTop,
                                                  schemeButtonWidth, ModernScale(26))
            modernPreviewCaption.SetBounds(margin + leftWidth + gap, captionTop, rightWidth, ModernScale(24))
            TabControl1.SetBounds(margin, bodyTop, leftWidth, bodyHeight)
            TabControl4.SetBounds(margin + leftWidth + gap, bodyTop, rightWidth, bodyHeight)
            TabControl1.PerformLayout()
            TabControl4.PerformLayout()
            LayoutParameterPages()
            LayoutPreviewPages()
            LayoutParameterSchemePanel()
            For Each control As Control In New Control() {GroupBox1, modernParametersCaption,
                                                           modernParameterSchemeButton, modernPreviewCaption,
                                                           TabControl1, TabControl4}
                control.BringToFront()
            Next
            Button2.BringToFront()
            Button7.BringToFront()
            Button8.BringToFront()
            If modernParameterSchemePanel.Visible Then modernParameterSchemePanel.BringToFront()
        Finally
            ResumeLayout(False)
            If rootScrollOffset <> Point.Empty Then AutoScrollPosition = rootScrollOffset
            modernLayoutInProgress = False
        End Try
    End Sub

    Private Sub LayoutModernHeader()
        modernHeader.Controls("ModernHeaderAccent").SetBounds(0, 0, ModernScale(6), modernHeader.ClientSize.Height)
        modernHeader.Controls("ModernTitle").SetBounds(ModernScale(22), ModernScale(8),
                                                        modernHeader.ClientSize.Width - ModernScale(38), ModernScale(36))
        modernHeader.Controls("ModernSubtitle").SetBounds(ModernScale(22), ModernScale(43),
                                                           modernHeader.ClientSize.Width - ModernScale(38), ModernScale(22))
    End Sub

    Private Sub LayoutProjectControls()
        Dim left As Integer = ModernScale(12)
        Dim gap As Integer = ModernScale(12)
        Dim columnWidth As Integer = Math.Max(ModernScale(170),
            (GroupBox1.ClientSize.Width - left * 2 - gap * 3) \ 4)
        Dim firstLabels As Label() = {Label12, Label15, Label14, Label2}
        Dim firstInputs As Control() = {CBox_ListNamesArrProject, CBox_ListNamesBridge,
                                        CB_NameAlignment, CBox_ListNamesTemplateXML}
        Dim secondLabels As Label() = {Label13, Label16, Label26, Label47}
        Dim secondInputs As Control() = {CB_ProjectSurface, CB_EgSurface, CB_NumberPillar, ComboBox11}
        For index As Integer = 0 To 3
            Dim x As Integer = left + index * (columnWidth + gap)
            firstLabels(index).Location = New Point(x, ModernScale(20))
            firstInputs(index).SetBounds(x, ModernScale(44), columnWidth, ModernScale(26))
            secondLabels(index).Location = New Point(x, ModernScale(72))
            If index < 3 Then
                secondInputs(index).SetBounds(x, ModernScale(96), columnWidth, ModernScale(26))
            Else
                Dim libraryWidth As Integer = ModernScale(104)
                secondInputs(index).SetBounds(x, ModernScale(96),
                    Math.Max(ModernScale(70), columnWidth - libraryWidth - ModernScale(6)), ModernScale(26))
                Button3.SetBounds(x + columnWidth - libraryWidth, ModernScale(96), libraryWidth, ModernScale(26))
            End If
        Next
        ChB_ProjectSurfaceInAlignment.Location = New Point(left, ModernScale(126))
        ChB_ElevationLand.Location = New Point(Math.Max(left + columnWidth + gap,
                                                       GroupBox1.ClientSize.Width \ 2), ModernScale(126))
        NUpD_ElevationLand.SetBounds(ChB_ElevationLand.Right + ModernScale(8), ModernScale(124),
                                    ModernScale(112), ModernScale(26))
    End Sub

    Private Sub LayoutFooterActions()
        Dim margin As Integer = ModernScale(16)
        Dim y As Integer = modernFooter.Top + ModernScale(14)
        Button2.SetBounds(margin, y, ModernScale(120), ModernScale(40))
        Button8.SetBounds(modernFooter.ClientSize.Width - margin - ModernScale(176), y,
                          ModernScale(176), ModernScale(40))
        Button7.SetBounds(Button8.Left - ModernScale(136), y, ModernScale(126), ModernScale(40))
    End Sub

    Private Sub LayoutParameterPages()
        LayoutModelGridPage(TabPage1, CB_NozzleTLC, Label1, DGV_Nozzle)
        LayoutSubFermenterPage()
        LayoutCabinetWallPage()
        LayoutNestedParameterPage(TabPage7, TabControl2)
        LayoutHandPage(TabPage9, CB_LeftHandTLC, Label23, DGV_LeftHand,
                       ChB_CalculateVerticalLineLeftHand, Nothing)
        LayoutHandPage(TabPage10, CB_RightHandTLC, Label24, DGV_RightHand,
                       ChB_CalculateVerticalLineRightHand, CheckBox1)
        LayoutNestedParameterPage(TabPage6, TabControl3)
        LayoutPostcardPage(TabPage11, CB_LeftPostcardTLC, Label25, DGV_LeftPostcard,
                           ChB_LeftPostcsrdFixedLenght, Nothing)
        LayoutPostcardPage(TabPage12, CB_RightPostcardTLC, Label5, DGV_RightPostcard,
                           ChB_RightPostcsrdFixedLenght, CheckBox2)
        LayoutRackPage()
        LayoutGrillagePage()
        LayoutPreparationPage()
        LayoutPilePage()
    End Sub

    Private Sub LayoutModelGridPage(page As TabPage, combo As ComboBox,
                                    modelLabel As Label, grid As DataGridView)
        Dim margin As Integer = ModernScale(10)
        Dim comboWidth As Integer = Math.Min(ModernScale(250), Math.Max(ModernScale(160), page.ClientSize.Width \ 2))
        combo.SetBounds(margin, margin, comboWidth, ModernScale(28))
        modelLabel.AutoSize = True
        modelLabel.Location = New Point(combo.Right + ModernScale(10), margin + ModernScale(4))
        grid.SetBounds(margin, margin + ModernScale(38),
                       Math.Max(ModernScale(100), page.ClientSize.Width - margin * 2),
                       Math.Max(ModernScale(80), page.ClientSize.Height - margin * 2 - ModernScale(38)))
    End Sub

    Private Sub LayoutSubFermenterPage()
        Dim margin As Integer = ModernScale(10)
        Button6.SetBounds(margin, margin, ModernScale(174), ModernScale(32))
        DGV_SubFermenter.SetBounds(margin, margin + ModernScale(40),
                                   Math.Max(ModernScale(100), TabPage16.ClientSize.Width - margin * 2),
                                   Math.Max(ModernScale(80), TabPage16.ClientSize.Height - margin * 2 - ModernScale(40)))
    End Sub

    Private Sub LayoutCabinetWallPage()
        Dim margin As Integer = ModernScale(10)
        Dim comboWidth As Integer = Math.Min(ModernScale(250),
            Math.Max(ModernScale(160), TabPage8.ClientSize.Width \ 2))
        CB_CabinetWallTLC.SetBounds(margin, margin, comboWidth, ModernScale(28))
        Label3.Location = New Point(CB_CabinetWallTLC.Right + ModernScale(10), margin + ModernScale(4))
        ChB_FixedHeightCabinetWall.Location = New Point(margin, ModernScale(45))
        DGV_CabinetWall.SetBounds(margin, ModernScale(72),
                                  Math.Max(ModernScale(100), TabPage8.ClientSize.Width - margin * 2),
                                  Math.Max(ModernScale(80), TabPage8.ClientSize.Height - ModernScale(82)))
    End Sub

    Private Sub LayoutNestedParameterPage(page As TabPage, nestedTabs As TabControl)
        Dim margin As Integer = ModernScale(4)
        nestedTabs.SetBounds(margin, margin,
                             Math.Max(ModernScale(100), page.ClientSize.Width - margin * 2),
                             Math.Max(ModernScale(100), page.ClientSize.Height - margin * 2))
    End Sub

    Private Sub LayoutHandPage(page As TabPage, combo As ComboBox, modelLabel As Label,
                               grid As DataGridView, calculateVertical As CheckBox,
                               copyFromLeft As CheckBox)
        LayoutNestedSidePage(page, combo, modelLabel, grid, calculateVertical, copyFromLeft)
    End Sub

    Private Sub LayoutPostcardPage(page As TabPage, combo As ComboBox, modelLabel As Label,
                                   grid As DataGridView, fixedLength As CheckBox,
                                   copyFromLeft As CheckBox)
        LayoutNestedSidePage(page, combo, modelLabel, grid, fixedLength, copyFromLeft)
    End Sub

    Private Sub LayoutNestedSidePage(page As TabPage, combo As ComboBox, modelLabel As Label,
                                     grid As DataGridView, optionCheckBox As CheckBox,
                                     copyFromLeft As CheckBox)
        Dim margin As Integer = ModernScale(8)
        Dim comboWidth As Integer = Math.Min(ModernScale(230),
            Math.Max(ModernScale(150), page.ClientSize.Width \ 2 - ModernScale(18)))
        combo.SetBounds(margin, margin, comboWidth, ModernScale(26))
        modelLabel.Location = New Point(combo.Right + ModernScale(8), margin + ModernScale(4))
        optionCheckBox.Location = New Point(margin, ModernScale(43))
        If copyFromLeft IsNot Nothing Then
            copyFromLeft.AutoSize = False
            copyFromLeft.SetBounds(Math.Max(margin, page.ClientSize.Width - ModernScale(248)),
                                   ModernScale(35), ModernScale(238), ModernScale(38))
        End If
        Dim gridTop As Integer = ModernScale(78)
        grid.SetBounds(margin, gridTop,
                       Math.Max(ModernScale(100), page.ClientSize.Width - margin * 2),
                       Math.Max(ModernScale(80), page.ClientSize.Height - gridTop - margin))
    End Sub

    Private Sub LayoutRackPage()
        Dim margin As Integer = ModernScale(10)
        Dim comboWidth As Integer = Math.Min(ModernScale(230), Math.Max(ModernScale(160), TabPage2.ClientSize.Width \ 2))
        CB_RackTLC.SetBounds(margin, margin, comboWidth, ModernScale(26))
        Label6.AutoSize = True
        Label6.Location = New Point(CB_RackTLC.Right + ModernScale(10), margin + ModernScale(4))
        ChB_CreateRack.Location = New Point(margin, ModernScale(43))
        ChB_fixedHeightRack.Location = New Point(Math.Max(ModernScale(310), TabPage2.ClientSize.Width - ModernScale(190)), ModernScale(43))
        ChB_EgeParallel.Location = New Point(margin, ModernScale(70))
        NUpD_CountRack.SetBounds(margin, ModernScale(96), ModernScale(74), ModernScale(26))
        Label4.AutoSize = True
        Label4.Location = New Point(NUpD_CountRack.Right + ModernScale(8), ModernScale(100))
        DGV_Rack.SetBounds(margin, ModernScale(129),
                           Math.Max(ModernScale(100), TabPage2.ClientSize.Width - margin * 2),
                           Math.Max(ModernScale(80), TabPage2.ClientSize.Height - ModernScale(139)))
    End Sub

    Private Sub LayoutGrillagePage()
        Dim margin As Integer = ModernScale(10)
        Dim comboWidth As Integer = Math.Min(ModernScale(230),
            Math.Max(ModernScale(160), TabPage4.ClientSize.Width \ 2))
        CB_GrillageTLC.SetBounds(margin, margin, comboWidth, ModernScale(26))
        Label7.Location = New Point(CB_GrillageTLC.Right + ModernScale(10), margin + ModernScale(4))
        ChB_CreateGrillage.Location = New Point(margin, ModernScale(43))
        ChB_GrillageEge.Location = New Point(Math.Max(ModernScale(190), TabPage4.ClientSize.Width \ 3),
                                             ModernScale(43))
        DGV_Grillage.SetBounds(margin, ModernScale(72),
                               Math.Max(ModernScale(100), TabPage4.ClientSize.Width - margin * 2),
                               Math.Max(ModernScale(80), TabPage4.ClientSize.Height - ModernScale(82)))
    End Sub

    Private Sub LayoutPreparationPage()
        Dim margin As Integer = ModernScale(10)
        Dim comboWidth As Integer = Math.Min(ModernScale(230),
            Math.Max(ModernScale(160), TabPage3.ClientSize.Width \ 2))
        CB_PreparationTLC.SetBounds(margin, margin, comboWidth, ModernScale(26))
        Label8.Location = New Point(CB_PreparationTLC.Right + ModernScale(10), margin + ModernScale(4))
        ChB_CreatePreparation.Location = New Point(margin, ModernScale(43))
        DGV_Preparation.SetBounds(margin, ModernScale(72),
                                 Math.Max(ModernScale(100), TabPage3.ClientSize.Width - margin * 2),
                                 Math.Max(ModernScale(80), TabPage3.ClientSize.Height - ModernScale(82)))
    End Sub

    Private Sub LayoutPilePage()
        Dim scrollOffset As New Point(Math.Max(0, -TabPage5.AutoScrollPosition.X),
                                      Math.Max(0, -TabPage5.AutoScrollPosition.Y))
        If scrollOffset <> Point.Empty Then TabPage5.AutoScrollPosition = Point.Empty
        Dim margin As Integer = ModernScale(10)
        Dim contentWidth As Integer = Math.Max(ModernScale(500), TabPage5.ClientSize.Width - margin * 2)
        Dim editorButtonWidth As Integer = ModernScale(150)
        Dim editorWidth As Integer = Math.Max(ModernScale(220), contentWidth - editorButtonWidth - ModernScale(8))
        CB_PileTLC.SetBounds(margin, margin, ModernScale(210), ModernScale(26))
        Label9.AutoSize = True
        Label9.Location = New Point(CB_PileTLC.Right + ModernScale(8), margin + ModernScale(4))
        ChB_InsertPileInRack.Location = New Point(margin, ModernScale(42))
        ChB_PileExpand.Location = New Point(margin + ModernScale(220), ModernScale(42))
        Label10.Text = "Схема свай в ряду (например, 500+5*800+500)"
        Label11.Text = "Схема свай в столбце (например, 500+2*800+500)"
        SetWrappedLabelBounds(Label10, margin, ModernScale(70), contentWidth, ModernScale(22))
        TxtB_PileRowDiagram.SetBounds(margin, ModernScale(93), editorWidth, ModernScale(26))
        Button1.SetBounds(margin + editorWidth + ModernScale(8), ModernScale(91),
                          editorButtonWidth, ModernScale(32))
        SetWrappedLabelBounds(Label11, margin, ModernScale(126), contentWidth, ModernScale(22))
        TxtB_PileCollDiagram.SetBounds(margin, ModernScale(149), editorWidth, ModernScale(26))
        Label22.Location = New Point(margin, ModernScale(183))
        NUpD_CountRowsPile.SetBounds(margin, ModernScale(206), ModernScale(72), ModernScale(26))
        Label18.Location = New Point(NUpD_CountRowsPile.Right + ModernScale(7), ModernScale(210))
        NUpD_OffsetRowsPile.SetBounds(margin + ModernScale(205), ModernScale(206), ModernScale(82), ModernScale(26))
        Label39.Location = New Point(NUpD_OffsetRowsPile.Left - ModernScale(22), ModernScale(210))
        Button4.SetBounds(margin + ModernScale(288), ModernScale(204), ModernScale(96), ModernScale(32))
        NUpD_CountColumnsPile.SetBounds(margin, ModernScale(240), ModernScale(72), ModernScale(26))
        Label19.Location = New Point(NUpD_CountColumnsPile.Right + ModernScale(7), ModernScale(244))
        NUpD_OffsetColumnsPile.SetBounds(margin + ModernScale(205), ModernScale(240), ModernScale(82), ModernScale(26))
        Label40.Location = New Point(NUpD_OffsetColumnsPile.Left - ModernScale(22), ModernScale(244))
        Button5.SetBounds(margin + ModernScale(288), ModernScale(238), ModernScale(96), ModernScale(32))
        Dim explanationLeft As Integer = Button4.Right + ModernScale(8)
        SetWrappedLabelBounds(Label20, explanationLeft, ModernScale(201),
                              Math.Max(ModernScale(108), margin + contentWidth - explanationLeft),
                              ModernScale(69))
        Dim gridTop As Integer = ModernScale(276)
        DGV_Piles.SetBounds(margin, gridTop, contentWidth,
                            Math.Max(ModernScale(90), TabPage5.ClientSize.Height - gridTop - margin))
        TabPage5.AutoScrollMinSize = New Size(ModernScale(520), ModernScale(380))
        If scrollOffset <> Point.Empty Then TabPage5.AutoScrollPosition = scrollOffset
    End Sub

    Private Sub SetWrappedLabelBounds(target As Label, x As Integer, y As Integer,
                                      width As Integer, height As Integer)
        target.AutoSize = False
        target.SetBounds(x, y, width, height)
        target.TextAlign = ContentAlignment.MiddleLeft
    End Sub

    Private Sub LayoutPreviewPages()
        Dim margin As Integer = ModernScale(10)
        Dim frontWidth As Integer = Math.Max(ModernScale(100), TabPage13.ClientSize.Width - margin * 2)
        Dim availableHeight As Integer = Math.Max(ModernScale(100), TabPage13.ClientSize.Height - margin * 2)
        PictureBox1.SetBounds(margin, margin, frontWidth, availableHeight)

        Dim halfWidth As Integer = Math.Max(ModernScale(160), (TabPage14.ClientSize.Width - margin * 3) \ 2)
        Dim rightLeft As Integer = margin * 2 + halfWidth
        modernLeftPreviewCaption.Location = New Point(margin, ModernScale(10))
        modernRightPreviewCaption.Location = New Point(rightLeft, ModernScale(10))
        Dim pictureTop As Integer = ModernScale(34)
        Dim pictureHeight As Integer = Math.Max(ModernScale(100), TabPage14.ClientSize.Height - pictureTop - margin)
        PictureBox2.SetBounds(margin, pictureTop, halfWidth, pictureHeight)
        PictureBox5.SetBounds(rightLeft, pictureTop,
                              Math.Max(ModernScale(100), TabPage14.ClientSize.Width - margin * 3 - halfWidth), pictureHeight)

        modernPileLayoutPreview.SetBounds(margin, margin,
            Math.Max(ModernScale(100), modernPileLayoutTab.ClientSize.Width - margin * 2),
            Math.Max(ModernScale(100), modernPileLayoutTab.ClientSize.Height - margin * 2))
        UpdatePileLayoutPreviewBitmap(False)
    End Sub

    Private Function ParameterGrids() As DataGridView()
        Return New DataGridView() {DGV_Nozzle, DGV_SubFermenter, DGV_CabinetWall,
                                   DGV_LeftHand, DGV_RightHand, DGV_LeftPostcard,
                                   DGV_RightPostcard, DGV_Rack, DGV_Grillage,
                                   DGV_Preparation, DGV_Piles}
    End Function

    Private Function ModernScale(logicalValue As Integer) As Integer
        Return CInt(Math.Round(logicalValue * Math.Max(1.0R, DeviceDpi / 96.0R)))
    End Function

    Private Function ModernScaleSize(logicalSize As Size) As Size
        Return New Size(ModernScale(logicalSize.Width), ModernScale(logicalSize.Height))
    End Function

    Friend Sub RegisterParameterScheme(page As TabPage, imagePath As String)
        If page Is Nothing Then Return
        modernParameterSchemePaths(page) = imagePath
        UpdateParameterSchemeButton()
    End Sub

    Friend Sub SetPileLayoutPreview(corners As Double(,), piles As Double(,), edges As Double(,))
        modernPileLayoutGeometry = PileLayoutGeometry.Create(corners, piles, edges)
        UpdatePileLayoutPreviewBitmap(True)
    End Sub

    Friend Sub SetPileLayoutPreview(corners As Double(,), piles As Double(,), edges As Double(,),
                                    axisData As Double(,), nominalSize As Double())
        modernPileLayoutGeometry = PileLayoutGeometry.Create(corners, piles, edges,
                                                              axisData, nominalSize)
        UpdatePileLayoutPreviewBitmap(True)
    End Sub

    Private Sub UpdatePileLayoutPreviewBitmap(force As Boolean)
        If modernPileLayoutPreview Is Nothing OrElse modernPileLayoutGeometry Is Nothing Then Return
        If Not modernPileLayoutGeometry.IsReady Then
            Dim invalidBitmap As Bitmap = modernPileLayoutBitmap
            modernPileLayoutBitmap = Nothing
            modernPileLayoutBitmapSize = Size.Empty
            modernPileLayoutPreview.Image = Nothing
            If invalidBitmap IsNot Nothing Then invalidBitmap.Dispose()
            Return
        End If

        Dim requestedSize As Size = modernPileLayoutPreview.ClientSize
        If requestedSize.Width <= 0 OrElse requestedSize.Height <= 0 Then requestedSize = New Size(640, 480)
        If Not force AndAlso requestedSize = modernPileLayoutBitmapSize AndAlso
           modernPileLayoutBitmap IsNot Nothing Then Return

        Dim replacement As Bitmap = PileLayoutPreview.Render(modernPileLayoutGeometry, requestedSize)
        Dim previous As Bitmap = modernPileLayoutBitmap
        modernPileLayoutBitmap = replacement
        modernPileLayoutBitmapSize = requestedSize
        modernPileLayoutPreview.Image = replacement
        If previous IsNot Nothing Then previous.Dispose()
    End Sub

    Private Sub ModernParameterSchemeButtonClick(sender As Object, e As EventArgs)
        If modernParameterSchemePanel.Visible Then
            HideParameterScheme()
            Return
        End If
        Dim selectedPage As TabPage = TabControl1.SelectedTab
        If selectedPage Is Nothing OrElse Not modernParameterSchemePaths.ContainsKey(selectedPage) Then Return
        Dim imagePath As String = modernParameterSchemePaths(selectedPage)
        If String.IsNullOrWhiteSpace(imagePath) OrElse Not File.Exists(imagePath) Then Return

        Dim replacement As Image = LoadParameterSchemeImage(imagePath)
        If replacement Is Nothing Then Return
        ClearParameterSchemeImage()
        modernParameterSchemeOwnedImage = replacement
        modernParameterSchemeImage.Image = replacement
        modernParameterSchemeTitle.Text = selectedPage.Text
        modernParameterSchemePanel.Visible = True
        LayoutParameterSchemePanel()
        modernParameterSchemePanel.BringToFront()
    End Sub

    Private Shared Function LoadParameterSchemeImage(imagePath As String) As Image
        Try
            Using source As Image = Image.FromFile(imagePath)
                Return New Bitmap(source)
            End Using
        Catch ex As IOException
            Return Nothing
        Catch ex As ArgumentException
            Return Nothing
        Catch ex As OutOfMemoryException
            Return Nothing
        End Try
    End Function

    Private Sub ModernParameterSchemeCloseClick(sender As Object, e As EventArgs)
        HideParameterScheme()
    End Sub

    Private Sub UpdateParameterSchemeButton()
        If modernParameterSchemeButton Is Nothing OrElse TabControl1 Is Nothing Then Return
        modernParameterSchemeButton.Enabled = TabControl1.SelectedTab IsNot Nothing AndAlso
                                                modernParameterSchemePaths.ContainsKey(TabControl1.SelectedTab)
    End Sub

    Private Sub HideParameterScheme()
        If modernParameterSchemePanel IsNot Nothing Then modernParameterSchemePanel.Visible = False
        ClearParameterSchemeImage()
    End Sub

    Private Sub ClearParameterSchemeImage()
        If modernParameterSchemeImage IsNot Nothing Then modernParameterSchemeImage.Image = Nothing
        If modernParameterSchemeOwnedImage IsNot Nothing Then modernParameterSchemeOwnedImage.Dispose()
        modernParameterSchemeOwnedImage = Nothing
    End Sub

    Private Sub LayoutParameterSchemePanel()
        If modernParameterSchemePanel Is Nothing OrElse TabControl1 Is Nothing Then Return
        Dim displayBounds As Rectangle = TabControl1.DisplayRectangle
        modernParameterSchemePanel.SetBounds(TabControl1.Left + displayBounds.Left,
                                             TabControl1.Top + displayBounds.Top,
                                             displayBounds.Width, displayBounds.Height)
        Dim margin As Integer = ModernScale(10)
        modernParameterSchemeTitle.SetBounds(margin, margin,
                                             Math.Max(1, modernParameterSchemePanel.ClientSize.Width -
                                                          margin * 3 - ModernScale(34)), ModernScale(28))
        modernParameterSchemeClose.SetBounds(modernParameterSchemePanel.ClientSize.Width - margin - ModernScale(34),
                                             margin, ModernScale(34), ModernScale(28))
        modernParameterSchemeImage.SetBounds(margin, modernParameterSchemeTitle.Bottom + ModernScale(6),
                                             Math.Max(1, modernParameterSchemePanel.ClientSize.Width - margin * 2),
                                             Math.Max(1, modernParameterSchemePanel.ClientSize.Height -
                                                          modernParameterSchemeTitle.Bottom - margin - ModernScale(6)))
    End Sub

    Private Sub DisposeModernAppearanceResources(sender As Object, e As EventArgs)
        If modernPreviewZoom IsNot Nothing Then modernPreviewZoom.Dispose()
        modernPreviewZoom = Nothing
        ClearParameterSchemeImage()
        If modernPileLayoutPreview IsNot Nothing Then modernPileLayoutPreview.Image = Nothing
        If modernPileLayoutBitmap IsNot Nothing Then modernPileLayoutBitmap.Dispose()
        modernPileLayoutBitmap = Nothing
        modernPileLayoutGeometry = Nothing
        If modernRegularFont IsNot Nothing Then modernRegularFont.Dispose()
        If modernBoldFont IsNot Nothing Then modernBoldFont.Dispose()
        If modernTitleFont IsNot Nothing Then modernTitleFont.Dispose()
        modernRegularFont = Nothing
        modernBoldFont = Nothing
        modernTitleFont = Nothing
    End Sub
End Class
