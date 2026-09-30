Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class FormPlacementBeams
    Friend NUpD_LongitudinalOffset As New NumericUpDown With {
        .Name = "NUpD_LongitudinalOffset",
        .Minimum = -999999999999D,
        .Maximum = 999999999999D
    }
    Private modernAppearanceApplied As Boolean
    Private modernRegularFont As Font
    Private modernTitleFont As Font
    Private modernBoldFont As Font

    Private Shared ReadOnly ModernCanvasColor As Color = Color.FromArgb(243, 246, 250)
    Private Shared ReadOnly ModernCardColor As Color = Color.White
    Private Shared ReadOnly ModernTextColor As Color = Color.FromArgb(32, 50, 77)
    Private Shared ReadOnly ModernMutedTextColor As Color = Color.FromArgb(93, 109, 133)
    Private Shared ReadOnly ModernAccentColor As Color = Color.FromArgb(37, 99, 235)
    Private Shared ReadOnly ModernAccentPaleColor As Color = Color.FromArgb(232, 240, 254)
    Private Shared ReadOnly ModernBorderColor As Color = Color.FromArgb(218, 226, 237)

    Private Sub ApplyModernAppearance()
        If modernAppearanceApplied Then
            Return
        End If
        modernAppearanceApplied = True

        SuspendLayout()
        Try
            CreateAppearanceFonts()
            ConfigureFormWindow()
            ConfigureProjectCard()
            ConfigureBridgeParameterCards()
            ConfigurePlacementCard()

            Dim rootLayout As TableLayoutPanel = CreateRootLayout()
            Dim header As Control = CreateHeader()
            Dim body As Control = CreateBody()
            Dim footer As Control = CreateFooter()
            rootLayout.Controls.Add(header, 0, 0)
            rootLayout.SetColumnSpan(header, 2)
            rootLayout.Controls.Add(body, 0, 1)
            rootLayout.SetColumnSpan(body, 2)
            rootLayout.Controls.Add(footer, 0, 2)
            rootLayout.SetColumnSpan(footer, 2)

            Controls.Clear()
            Controls.Add(rootLayout)
            ApplyFontAndColors(Me)
            ConfigureActionButtons()
            ConfigureWindowSize()
            AddHandler Disposed, AddressOf DisposeAppearanceResources
        Finally
            ResumeLayout(True)
        End Try
    End Sub

    Private Sub CreateAppearanceFonts()
        modernRegularFont = New Font("Segoe UI", 9.5!, FontStyle.Regular, GraphicsUnit.Point)
        modernTitleFont = New Font("Segoe UI Semibold", 17.0!, FontStyle.Bold, GraphicsUnit.Point)
        modernBoldFont = New Font("Segoe UI Semibold", 9.5!, FontStyle.Bold, GraphicsUnit.Point)
    End Sub

    Private Sub ConfigureFormWindow()
        Font = modernRegularFont
        AutoScaleDimensions = CurrentAutoScaleDimensions
        BackColor = ModernCanvasColor
        ForeColor = ModernTextColor
        MaximumSize = Size.Empty
        Padding = Padding.Empty
        DoubleBuffered = True
        StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConfigureWindowSize()
        Dim workArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim safeWidth As Integer = Math.Max(760, workArea.Width - ScaleValue(32))
        Dim safeHeight As Integer = Math.Max(560, workArea.Height - ScaleValue(32))
        Dim preferredSize As Size = ScaleSize(New Size(1120, 800))
        Dim requestedMinimum As Size = ScaleSize(New Size(998, 680))

        MinimumSize = New Size(Math.Min(requestedMinimum.Width, safeWidth),
                               Math.Min(requestedMinimum.Height, safeHeight))
        Size = New Size(Math.Min(preferredSize.Width, safeWidth),
                        Math.Min(preferredSize.Height, safeHeight))
    End Sub

    Private Function CreateRootLayout() As TableLayoutPanel
        Dim layout As New TableLayoutPanel()
        layout.Name = "ModernRootLayout"
        layout.Dock = DockStyle.Fill
        layout.BackColor = ModernCanvasColor
        layout.ColumnCount = 2
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        layout.RowCount = 3
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(88)))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        layout.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(68)))
        layout.Margin = Padding.Empty
        layout.Padding = Padding.Empty
        Return layout
    End Function

    Private Function CreateHeader() As Control
        Dim header As New Panel()
        header.Name = "ModernHeader"
        header.Dock = DockStyle.Fill
        header.BackColor = ModernCardColor
        header.Margin = Padding.Empty
        header.Padding = New Padding(ScaleValue(28), ScaleValue(13), ScaleValue(22), ScaleValue(10))

        Dim accent As New Panel()
        accent.Name = "ModernHeaderAccent"
        accent.Dock = DockStyle.Left
        accent.Width = ScaleValue(6)
        accent.BackColor = ModernAccentColor
        accent.Margin = Padding.Empty

        Dim textLayout As New TableLayoutPanel()
        textLayout.Name = "ModernHeaderText"
        textLayout.Dock = DockStyle.Fill
        textLayout.ColumnCount = 1
        textLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        textLayout.RowCount = 2
        textLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 58.0!))
        textLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 42.0!))
        textLayout.Margin = Padding.Empty
        textLayout.Padding = New Padding(ScaleValue(10), 0, 0, 0)

        Dim title As New Label()
        title.Name = "ModernTitle"
        title.AutoSize = False
        title.Dock = DockStyle.Fill
        title.Font = modernTitleFont
        title.ForeColor = ModernTextColor
        title.Text = "Раскладка мостовых балок"
        title.TextAlign = ContentAlignment.BottomLeft
        title.Margin = Padding.Empty

        Dim subtitle As New Label()
        subtitle.Name = "ModernSubtitle"
        subtitle.AutoSize = False
        subtitle.Dock = DockStyle.Fill
        subtitle.ForeColor = ModernMutedTextColor
        subtitle.Text = "Настройте сооружение, параметры рядов и данные раскладки"
        subtitle.TextAlign = ContentAlignment.TopLeft
        subtitle.Margin = Padding.Empty

        textLayout.Controls.Add(title, 0, 0)
        textLayout.Controls.Add(subtitle, 0, 1)
        header.Controls.Add(textLayout)
        header.Controls.Add(accent)
        Return header
    End Function

    Private Function CreateBody() As Control
        Dim body As New TableLayoutPanel()
        body.Name = "ModernBody"
        body.Dock = DockStyle.Fill
        body.BackColor = ModernCanvasColor
        body.ColumnCount = 2
        body.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, GetSidebarWidth()))
        body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        body.RowCount = 1
        body.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        body.Margin = Padding.Empty
        body.Padding = New Padding(ScaleValue(18), ScaleValue(16), ScaleValue(18), ScaleValue(10))

        Dim parameterArea As Control = CreateParameterArea()
        parameterArea.Margin = New Padding(0, 0, ScaleValue(14), 0)
        body.Controls.Add(parameterArea, 0, 0)

        GroupBox2.Dock = DockStyle.Fill
        GroupBox2.Margin = Padding.Empty
        body.Controls.Add(GroupBox2, 1, 0)
        Return body
    End Function

    Private Function CreateParameterArea() As Control
        Dim scrollArea As New Panel()
        scrollArea.Name = "ModernParameterScrollArea"
        scrollArea.Dock = DockStyle.Fill
        scrollArea.BackColor = ModernCanvasColor
        scrollArea.AutoScroll = True
        scrollArea.Margin = Padding.Empty

        Dim stack As New TableLayoutPanel()
        stack.Name = "ModernParameterStack"
        stack.Dock = DockStyle.Top
        stack.AutoSize = True
        stack.AutoSizeMode = AutoSizeMode.GrowAndShrink
        stack.ColumnCount = 1
        stack.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        stack.RowCount = 2
        stack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        stack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        stack.Margin = Padding.Empty
        stack.Padding = Padding.Empty

        GroupBox1.Dock = DockStyle.Fill
        GroupBox1.Margin = New Padding(0, 0, 0, ScaleValue(12))
        stack.Controls.Add(GroupBox1, 0, 0)

        Dim sides As New TableLayoutPanel()
        sides.Name = "ModernBridgeSides"
        sides.Dock = DockStyle.Fill
        sides.AutoSize = True
        sides.AutoSizeMode = AutoSizeMode.GrowAndShrink
        sides.ColumnCount = 2
        sides.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        sides.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        sides.RowCount = 1
        sides.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        sides.Margin = Padding.Empty
        sides.Padding = Padding.Empty

        GroupBox3.Dock = DockStyle.Fill
        GroupBox3.Margin = New Padding(0, 0, ScaleValue(6), 0)
        GroupBox4.Dock = DockStyle.Fill
        GroupBox4.Margin = New Padding(ScaleValue(6), 0, 0, 0)
        sides.Controls.Add(GroupBox3, 0, 0)
        sides.Controls.Add(GroupBox4, 1, 0)
        stack.Controls.Add(sides, 0, 1)

        scrollArea.Controls.Add(stack)
        Return scrollArea
    End Function

    Private Function CreateFooter() As Control
        Dim footer As New Panel()
        footer.Name = "ModernFooter"
        footer.Dock = DockStyle.Fill
        footer.BackColor = ModernCardColor
        footer.Margin = Padding.Empty
        footer.Padding = New Padding(ScaleValue(18), ScaleValue(12), ScaleValue(18), ScaleValue(12))

        Dim actions As New FlowLayoutPanel()
        actions.Name = "ModernFooterActions"
        actions.Dock = DockStyle.Right
        actions.AutoSize = True
        actions.AutoSizeMode = AutoSizeMode.GrowAndShrink
        actions.FlowDirection = FlowDirection.RightToLeft
        actions.WrapContents = False
        actions.Margin = Padding.Empty
        actions.Padding = Padding.Empty

        Button2.Size = ScaleSize(New Size(174, 40))
        Button2.Margin = New Padding(ScaleValue(10), 0, 0, 0)
        Button1.Size = ScaleSize(New Size(116, 40))
        Button1.Margin = Padding.Empty
        actions.Controls.Add(Button2)
        actions.Controls.Add(Button1)
        footer.Controls.Add(actions)
        Return footer
    End Function

    Private Sub ConfigureProjectCard()
        ConfigureCard(GroupBox1, "Настройки проекта")
        GroupBox1.AutoSize = True
        GroupBox1.AutoSizeMode = AutoSizeMode.GrowAndShrink
        GroupBox1.Controls.Clear()

        Dim fields As TableLayoutPanel = CreateFieldsTable(3)
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, ScaleValue(142)))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, ScaleValue(88)))

        AddFieldRow(fields, 0, Label12, CBox_ListModelStructures)
        ButtonCreateArr.Dock = DockStyle.Fill
        ButtonCreateArr.Margin = New Padding(ScaleValue(6), ScaleValue(4), 0, ScaleValue(4))
        fields.Controls.Add(ButtonCreateArr, 2, 0)
        AddFieldRow(fields, 1, Label16, CBox_ListNameStructures, 2)
        AddFieldRow(fields, 2, Label13, CBox_ListProjectSurfaces, 2)
        AddFieldRow(fields, 3, Label14, CBox_ListAxisRoads, 2)
        AddFieldRow(fields, 4, Label2, CBox_ListTemplateXML, 2)

        ConfigureCheckBox(CB_SurfaceFromAlign)
        fields.Controls.Add(CB_SurfaceFromAlign, 0, 5)
        fields.SetColumnSpan(CB_SurfaceFromAlign, 3)
        ConfigureCheckBox(ChB_CenterBeam)
        fields.Controls.Add(ChB_CenterBeam, 0, 6)
        fields.SetColumnSpan(ChB_CenterBeam, 3)
        AddFieldRow(fields, 7, Label21, CBox_ListPlacementBeams, 2)

        Label17.Visible = False
        Label17.Size = Size.Empty
        Label17.Location = Point.Empty
        GroupBox1.Controls.Add(Label17)
        GroupBox1.Controls.Add(fields)
    End Sub

    Private Sub ConfigureBridgeParameterCards()
        ConfigureBridgeParameterCard(GroupBox3,
                                     "Слева",
                                     Label3, NumericUpDown3, "До первой балки, мм",
                                     Label5, NumericUpDown5, "До крайней балки, мм",
                                     Label6, NUpD_CountLeftRows, "Рядов",
                                     Label4, NUpD_dimLeftBridge, "Ширина, мм")
        ConfigureBridgeParameterCard(GroupBox4,
                                     "Справа",
                                     Label8, NumericUpDown8, "До первой балки, мм",
                                     Label10, NumericUpDown9, "До крайней балки, мм",
                                     Label9, NUpD_CountRightRows, "Рядов",
                                     Label11, NUpD_dimRightBridge, "Ширина, мм")
    End Sub

    Private Sub ConfigureBridgeParameterCard(card As GroupBox,
                                             cardTitle As String,
                                             firstLabel As Label,
                                             firstInput As NumericUpDown,
                                             firstCaption As String,
                                             lastLabel As Label,
                                             lastInput As NumericUpDown,
                                             lastCaption As String,
                                             rowsLabel As Label,
                                             rowsInput As NumericUpDown,
                                             rowsCaption As String,
                                             widthLabel As Label,
                                             widthInput As NumericUpDown,
                                             widthCaption As String)
        ConfigureCard(card, cardTitle)
        card.AutoSize = True
        card.AutoSizeMode = AutoSizeMode.GrowAndShrink

        PreserveFullCaption(firstLabel, firstCaption)
        PreserveFullCaption(lastLabel, lastCaption)
        PreserveFullCaption(rowsLabel, rowsCaption)
        PreserveFullCaption(widthLabel, widthCaption)

        card.Controls.Clear()
        Dim fields As TableLayoutPanel = CreateFieldsTable(2)
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, ScaleValue(78)))
        AddCompactFieldRow(fields, 0, firstLabel, firstInput)
        AddCompactFieldRow(fields, 1, lastLabel, lastInput)
        AddCompactFieldRow(fields, 2, rowsLabel, rowsInput)
        AddCompactFieldRow(fields, 3, widthLabel, widthInput)
        card.Controls.Add(fields)
        AddHandler card.Layout,
            Sub(sender As Object, e As LayoutEventArgs)
                EnsureBridgeCardHeight(card, fields)
            End Sub
    End Sub

    Private Sub EnsureBridgeCardHeight(card As GroupBox, fields As TableLayoutPanel)
        Dim availableWidth As Integer = Math.Max(1, card.DisplayRectangle.Width)
        Dim fieldsHeight As Integer = fields.GetPreferredSize(New Size(availableWidth, 0)).Height
        Dim requiredHeight As Integer = card.DisplayRectangle.Top + fieldsHeight + card.Padding.Bottom
        If card.MinimumSize.Height <> requiredHeight Then
            card.MinimumSize = New Size(0, requiredHeight)
        End If
    End Sub

    Private Sub ConfigurePlacementCard()
        ConfigureCard(GroupBox2, "Параметры раскладки")
        GroupBox2.Controls.Clear()

        Dim layout As New TableLayoutPanel()
        layout.Name = "ModernPlacementLayout"
        layout.Dock = DockStyle.Fill
        layout.BackColor = ModernCardColor
        layout.ColumnCount = 1
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        layout.RowCount = 2
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        layout.Padding = New Padding(ScaleValue(10), ScaleValue(8), ScaleValue(10), ScaleValue(10))
        layout.Margin = Padding.Empty

        Dim toolbar As TableLayoutPanel = CreateFieldsTable(2)
        toolbar.Name = "ModernSchemeToolbar"
        toolbar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        toolbar.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, ScaleValue(110)))
        EnsureAutoSizeRow(toolbar, 0)
        EnsureAutoSizeRow(toolbar, 1)
        ConfigureFieldLabel(Label15)
        toolbar.Controls.Add(Label15, 0, 0)
        ConfigureInputControl(CBox_ShemaPlacementBeams)
        toolbar.Controls.Add(CBox_ShemaPlacementBeams, 0, 1)
        Button3.Dock = DockStyle.Fill
        Button3.Margin = New Padding(ScaleValue(8), ScaleValue(4), 0, ScaleValue(4))
        toolbar.Controls.Add(Button3, 1, 1)

        ConfigureTabs()
        TabControl1.Dock = DockStyle.Fill
        TabControl1.Margin = New Padding(0, ScaleValue(8), 0, 0)
        layout.Controls.Add(toolbar, 0, 0)
        layout.Controls.Add(TabControl1, 0, 1)
        GroupBox2.Controls.Add(layout)
    End Sub

    Private Sub ConfigureTabs()
        TabControl1.Padding = New Point(ScaleValue(16), ScaleValue(6))
        TabPage2.Text = "Пролеты"
        TabPage3.Text = "Ряды"
        TabPage1.Text = "Опоры"
        ConfigureSpanTab()
        ConfigureRowsTab()
        ConfigureSupportsTab()
    End Sub

    Private Sub ConfigureSpanTab()
        Dim fields As TableLayoutPanel = CreateFieldsTable(2)
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38.0!))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62.0!))
        EnsureAutoSizeRow(fields, 0)
        fields.Controls.Add(CreateLabeledEditor(Label1, NUpD_CountProlet), 0, 0)
        fields.Controls.Add(CreateLabeledEditor(Label7, CBox_AlbumsBeams), 1, 0)
        ConfigureTabPage(TabPage2, fields, DG_ProletListBeams)
    End Sub

    Private Sub ConfigureRowsTab()
        Dim fields As TableLayoutPanel = CreateFieldsTable(3)
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.333!))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.333!))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.334!))
        EnsureAutoSizeRow(fields, 0)
        Dim longitudinalOffsetLabel As New Label With {
            .Name = "LabelLongitudinalOffset",
            .Text = "Продольное смещение, мм"
        }
        fields.Controls.Add(CreateLabeledEditor(Label18, NUpD_TraverseOffset), 0, 0)
        fields.Controls.Add(CreateLabeledEditor(longitudinalOffsetLabel, NUpD_LongitudinalOffset), 1, 0)
        fields.Controls.Add(CreateLabeledEditor(Label20, NUpD_VerticalOffset), 2, 0)
        ConfigureTabPage(TabPage3, fields, DG_RowProperties)
    End Sub

    Private Sub ConfigureSupportsTab()
        Dim fields As TableLayoutPanel = CreateFieldsTable(1)
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        EnsureAutoSizeRow(fields, 0)
        EnsureAutoSizeRow(fields, 1)
        ConfigureFieldLabel(Label19)
        fields.Controls.Add(Label19, 0, 0)

        Dim editorLine As TableLayoutPanel = CreateFieldsTable(2)
        editorLine.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 62.0!))
        editorLine.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 38.0!))
        EnsureAutoSizeRow(editorLine, 0)
        ConfigureInputControl(MaskTB_PK)
        editorLine.Controls.Add(MaskTB_PK, 0, 0)
        ConfigureCheckBox(CheckBox3)
        editorLine.Controls.Add(CheckBox3, 1, 0)
        fields.Controls.Add(editorLine, 0, 1)
        ConfigureTabPage(TabPage1, fields, DG_PillarsProperties)
    End Sub

    Private Function CreateLabeledEditor(fieldLabel As Label, input As Control) As TableLayoutPanel
        Dim field As TableLayoutPanel = CreateFieldsTable(1)
        field.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        EnsureAutoSizeRow(field, 0)
        EnsureAutoSizeRow(field, 1)
        field.Margin = New Padding(0, 0, ScaleValue(10), 0)
        ConfigureFieldLabel(fieldLabel)
        fieldLabel.Margin = New Padding(0, ScaleValue(2), 0, 0)
        field.Controls.Add(fieldLabel, 0, 0)
        ConfigureInputControl(input)
        field.Controls.Add(input, 0, 1)
        Return field
    End Function

    Private Sub ConfigureTabPage(page As TabPage,
                                 fields As TableLayoutPanel,
                                 targetGrid As DataGridView)
        page.SuspendLayout()
        Try
            page.BackColor = ModernCardColor
            page.Padding = New Padding(ScaleValue(10))
            page.Controls.Clear()

            Dim layout As New TableLayoutPanel()
            layout.Name = page.Name & "ModernLayout"
            layout.Dock = DockStyle.Fill
            layout.BackColor = ModernCardColor
            layout.ColumnCount = 1
            layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            layout.RowCount = 2
            layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
            layout.Margin = Padding.Empty
            layout.Padding = Padding.Empty

            fields.Margin = New Padding(0, 0, 0, ScaleValue(8))
            ConfigureModernGrid(targetGrid)
            targetGrid.Dock = DockStyle.Fill
            targetGrid.Margin = Padding.Empty
            layout.Controls.Add(fields, 0, 0)
            layout.Controls.Add(targetGrid, 0, 1)
            page.Controls.Add(layout)
        Finally
            page.ResumeLayout(True)
        End Try
    End Sub

    Private Function CreateFieldsTable(columnCount As Integer) As TableLayoutPanel
        Dim fields As New TableLayoutPanel()
        fields.Dock = DockStyle.Top
        fields.AutoSize = True
        fields.AutoSizeMode = AutoSizeMode.GrowAndShrink
        fields.BackColor = ModernCardColor
        fields.ColumnCount = columnCount
        fields.RowCount = 0
        fields.Margin = Padding.Empty
        fields.Padding = Padding.Empty
        Return fields
    End Function

    Private Sub AddFieldRow(fields As TableLayoutPanel,
                            rowIndex As Integer,
                            fieldLabel As Label,
                            input As Control,
                            Optional inputColumnSpan As Integer = 1)
        EnsureAutoSizeRow(fields, rowIndex)
        ConfigureFieldLabel(fieldLabel)
        fields.Controls.Add(fieldLabel, 0, rowIndex)
        ConfigureInputControl(input)
        fields.Controls.Add(input, 1, rowIndex)
        If inputColumnSpan > 1 Then
            fields.SetColumnSpan(input, inputColumnSpan)
        End If
    End Sub

    Private Sub AddCompactFieldRow(fields As TableLayoutPanel,
                                   rowIndex As Integer,
                                   fieldLabel As Label,
                                   input As Control)
        EnsureAutoSizeRow(fields, rowIndex)
        ConfigureFieldLabel(fieldLabel)
        fields.Controls.Add(fieldLabel, 0, rowIndex)
        ConfigureInputControl(input)
        fields.Controls.Add(input, 1, rowIndex)
    End Sub

    Private Sub EnsureAutoSizeRow(fields As TableLayoutPanel, rowIndex As Integer)
        While fields.RowCount <= rowIndex
            fields.RowCount += 1
            fields.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        End While
    End Sub

    Private Sub ConfigureCard(card As GroupBox, caption As String)
        card.Text = caption
        card.BackColor = ModernCardColor
        card.ForeColor = ModernTextColor
        card.Padding = New Padding(ScaleValue(12), ScaleValue(10), ScaleValue(12), ScaleValue(12))
        card.FlatStyle = FlatStyle.Flat
    End Sub

    Private Sub ConfigureFieldLabel(fieldLabel As Label)
        fieldLabel.AutoSize = True
        fieldLabel.Dock = DockStyle.Fill
        fieldLabel.ForeColor = ModernTextColor
        fieldLabel.TextAlign = ContentAlignment.MiddleLeft
        fieldLabel.Margin = New Padding(0, ScaleValue(4), ScaleValue(8), ScaleValue(4))
        fieldLabel.MinimumSize = Size.Empty
    End Sub

    Private Sub ConfigureInputControl(input As Control)
        input.Dock = DockStyle.Fill
        input.Margin = New Padding(0, ScaleValue(4), 0, ScaleValue(4))
        input.MinimumSize = New Size(0, ScaleValue(30))
        input.BackColor = ModernCardColor
        input.ForeColor = ModernTextColor

        Dim combo As ComboBox = TryCast(input, ComboBox)
        If combo IsNot Nothing Then
            combo.FlatStyle = FlatStyle.Flat
        End If

        Dim mask As MaskedTextBox = TryCast(input, MaskedTextBox)
        If mask IsNot Nothing Then
            mask.BorderStyle = BorderStyle.FixedSingle
        End If
    End Sub

    Private Sub ConfigureCheckBox(checkBox As CheckBox)
        checkBox.AutoSize = True
        checkBox.Dock = DockStyle.Fill
        checkBox.ForeColor = ModernTextColor
        checkBox.Margin = New Padding(0, ScaleValue(5), 0, ScaleValue(5))
        checkBox.MinimumSize = New Size(0, ScaleValue(28))
    End Sub

    Private Sub PreserveFullCaption(fieldLabel As Label, shortCaption As String)
        Dim fullCaption As String = fieldLabel.Text.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ").Trim()
        ToolTip1.SetToolTip(fieldLabel, fullCaption)
        fieldLabel.Text = shortCaption
    End Sub

    Private Sub ConfigureModernGrid(targetGrid As DataGridView)
        targetGrid.BackgroundColor = ModernCardColor
        targetGrid.BorderStyle = BorderStyle.None
        targetGrid.GridColor = ModernBorderColor
        targetGrid.EnableHeadersVisualStyles = False
        targetGrid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        targetGrid.ColumnHeadersDefaultCellStyle.BackColor = ModernAccentPaleColor
        targetGrid.ColumnHeadersDefaultCellStyle.ForeColor = ModernTextColor
        targetGrid.ColumnHeadersDefaultCellStyle.Font = modernBoldFont
        targetGrid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        targetGrid.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTextColor
        targetGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        targetGrid.ColumnHeadersHeight = ScaleValue(36)
        targetGrid.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252)
        targetGrid.RowHeadersDefaultCellStyle.ForeColor = ModernTextColor
        targetGrid.RowHeadersDefaultCellStyle.Font = modernRegularFont
        targetGrid.RowHeadersDefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        targetGrid.RowHeadersDefaultCellStyle.SelectionForeColor = ModernTextColor
        targetGrid.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False
        targetGrid.RowHeadersWidth = Math.Max(targetGrid.RowHeadersWidth, ScaleValue(96))
        targetGrid.DefaultCellStyle.BackColor = ModernCardColor
        targetGrid.DefaultCellStyle.ForeColor = ModernTextColor
        targetGrid.DefaultCellStyle.Font = modernRegularFont
        targetGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 232, 255)
        targetGrid.DefaultCellStyle.SelectionForeColor = ModernTextColor
        targetGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(247, 250, 255)
        targetGrid.RowTemplate.Height = ScaleValue(30)
        targetGrid.ScrollBars = ScrollBars.Both

        If Object.ReferenceEquals(targetGrid, DG_RowProperties) Then
            Column9.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Column9.MinimumWidth = ScaleValue(130)
            Column9.FlatStyle = FlatStyle.Flat
        ElseIf Object.ReferenceEquals(targetGrid, DG_PillarsProperties) Then
            Column5.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Column5.MinimumWidth = ScaleValue(90)
            Column5.FlatStyle = FlatStyle.Flat
        End If
    End Sub

    Private Sub ConfigureActionButtons()
        ConfigureSecondaryButton(Button1)
        ConfigurePrimaryButton(Button2)
        Button2.Text = "Разложить балки"
        ConfigureSecondaryButton(Button3)
        ConfigureSecondaryButton(ButtonCreateArr)
    End Sub

    Private Sub ConfigurePrimaryButton(button As Button)
        button.BackColor = ModernAccentColor
        button.ForeColor = Color.White
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.Font = modernBoldFont
        button.UseVisualStyleBackColor = False
    End Sub

    Private Sub ConfigureSecondaryButton(button As Button)
        button.BackColor = ModernCardColor
        button.ForeColor = ModernAccentColor
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderColor = ModernBorderColor
        button.FlatAppearance.BorderSize = 1
        button.UseVisualStyleBackColor = False
    End Sub

    Private Sub ApplyFontAndColors(parent As Control)
        For Each child As Control In parent.Controls
            If Not Object.ReferenceEquals(child.Font, modernTitleFont) AndAlso
               Not Object.ReferenceEquals(child.Font, modernBoldFont) Then
                child.Font = modernRegularFont
            End If
            ApplyFontAndColors(child)
        Next
    End Sub

    Private Function ScaleValue(logicalValue As Integer) As Integer
        Return CInt(Math.Round(logicalValue * DeviceDpi / 96.0R))
    End Function

    Private Function GetSidebarWidth() As Integer
        Dim workArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim expectedWindowWidth As Integer = Math.Min(ScaleValue(1120),
                                                      Math.Max(760, workArea.Width - ScaleValue(32)))
        Dim contentWidth As Integer = Math.Max(ScaleValue(760), expectedWindowWidth - ScaleValue(36))
        Return Math.Min(ScaleValue(430), CInt(Math.Round(contentWidth * 0.4R)))
    End Function

    Private Function ScaleSize(logicalSize As Size) As Size
        Return New Size(ScaleValue(logicalSize.Width), ScaleValue(logicalSize.Height))
    End Function

    Private Sub DisposeAppearanceResources(sender As Object, e As EventArgs)
        RemoveHandler Disposed, AddressOf DisposeAppearanceResources
        If modernTitleFont IsNot Nothing Then
            modernTitleFont.Dispose()
            modernTitleFont = Nothing
        End If
        If modernBoldFont IsNot Nothing Then
            modernBoldFont.Dispose()
            modernBoldFont = Nothing
        End If
        If modernRegularFont IsNot Nothing Then
            modernRegularFont.Dispose()
            modernRegularFont = Nothing
        End If
    End Sub
End Class
