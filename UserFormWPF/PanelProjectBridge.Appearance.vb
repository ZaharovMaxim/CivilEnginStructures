Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class PanelProjectBridge
    Private modernAppearanceApplied As Boolean
    Private modernRegularFont As Font
    Private modernBoldFont As Font
    Private modernTitleFont As Font
    Private modernPaletteLayoutPending As Boolean
    Private modernVisualStyleApplying As Boolean
    Private modernVisualStylePending As Boolean
    Private modernVisualStyleGeneration As Long
    Private modernToolbarCaptionsUpdating As Boolean

    Private Shared ReadOnly ModernCanvasColor As Color = Color.FromArgb(243, 246, 250)
    Private Shared ReadOnly ModernCardColor As Color = Color.White
    Private Shared ReadOnly ModernTextColor As Color = Color.FromArgb(32, 50, 77)
    Private Shared ReadOnly ModernMutedColor As Color = Color.FromArgb(93, 109, 133)
    Private Shared ReadOnly ModernAccentColor As Color = Color.FromArgb(37, 99, 235)
    Private Shared ReadOnly ModernAccentPaleColor As Color = Color.FromArgb(232, 240, 254)
    Private Shared ReadOnly ModernBorderColor As Color = Color.FromArgb(218, 226, 237)
    Private Shared ReadOnly ModernMinimumCanvasSize As New Size(320, 480)

    Private Sub ApplyModernAppearance()
        If modernAppearanceApplied Then Return
        modernAppearanceApplied = True

        SuspendLayout()
        Try
            CreateModernFonts()
            AutoScaleMode = AutoScaleMode.Dpi
            AutoScaleDimensions = New SizeF(DeviceDpi, DeviceDpi)
            Font = modernRegularFont
            MinimumSize = Size.Empty
            Padding = Padding.Empty

            GroupBox1.Controls.Clear()
            GroupBox1.Visible = False
            GroupBox1.Size = Size.Empty
            GroupBox1.Location = Point.Empty

            Dim scrollHost As New Panel()
            scrollHost.Name = "ModernPaletteScroll"
            scrollHost.Dock = DockStyle.Fill
            scrollHost.BackColor = ModernCanvasColor
            scrollHost.AutoScroll = True
            scrollHost.Margin = Padding.Empty

            Dim canvas As TableLayoutPanel = CreateCanvas()
            Dim split As SplitContainer = CreatePaletteSplit()
            canvas.Controls.Add(CreateHeader(), 0, 0)
            canvas.Controls.Add(CreateToolbar(), 0, 1)
            canvas.Controls.Add(CreateSelectorCard(), 0, 2)
            canvas.Controls.Add(split, 0, 3)

            Controls.Clear()
            Controls.Add(scrollHost)
            scrollHost.Controls.Add(GroupBox1)
            scrollHost.Controls.Add(canvas)

            Dim lastClientSize As Size = scrollHost.ClientSize
            Dim resizeCanvas As EventHandler =
                Sub(sender As Object, e As EventArgs)
                    Dim clientSizeChanged As Boolean = scrollHost.ClientSize <> lastClientSize
                    lastClientSize = scrollHost.ClientSize
                    ResizePaletteCanvas(scrollHost, canvas)
                    If clientSizeChanged Then QueuePaletteLayout(scrollHost, canvas)
                End Sub
            AddHandler scrollHost.ClientSizeChanged, resizeCanvas
            AddHandler scrollHost.HandleDestroyed,
                Sub(sender As Object, e As EventArgs)
                    modernPaletteLayoutPending = False
                End Sub
            ResizePaletteCanvas(scrollHost, canvas)
            ConfigureSplitResize(split)
            ReapplyModernVisualStyle()
            ConfigureModernStyleEvents(scrollHost, canvas, split)
            AddHandler Disposed, AddressOf DisposeModernResources
        Finally
            ResumeLayout(True)
        End Try
    End Sub

    Private Sub CreateModernFonts()
        modernRegularFont = New Font("Segoe UI", 9.5!, FontStyle.Regular, GraphicsUnit.Point)
        modernBoldFont = New Font("Segoe UI Semibold", 9.5!, FontStyle.Bold, GraphicsUnit.Point)
        modernTitleFont = New Font("Segoe UI Semibold", 14.0!, FontStyle.Bold, GraphicsUnit.Point)
    End Sub

    Private Function CreateCanvas() As TableLayoutPanel
        Dim canvas As New TableLayoutPanel()
        canvas.Name = "ModernPaletteCanvas"
        canvas.Location = Point.Empty
        canvas.BackColor = ModernCanvasColor
        canvas.ColumnCount = 1
        canvas.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        canvas.RowCount = 4
        canvas.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(62)))
        canvas.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(80)))
        canvas.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(66)))
        canvas.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        canvas.Padding = New Padding(ScaleValue(8))
        canvas.Margin = Padding.Empty
        canvas.MinimumSize = ScaleSize(ModernMinimumCanvasSize)
        Return canvas
    End Function

    Private Function CreateHeader() As Control
        Dim header As New Panel()
        header.Name = "ModernPaletteHeader"
        header.Dock = DockStyle.Fill
        header.BackColor = ModernCardColor
        header.Margin = New Padding(0, 0, 0, ScaleValue(8))
        header.Padding = New Padding(ScaleValue(18), ScaleValue(8), ScaleValue(10), ScaleValue(8))

        Dim accent As New Panel()
        accent.Dock = DockStyle.Left
        accent.Width = ScaleValue(5)
        accent.BackColor = ModernAccentColor

        Dim title As New Label()
        title.Name = "ModernPaletteTitle"
        title.Dock = DockStyle.Fill
        title.AutoSize = False
        title.Font = modernTitleFont
        title.ForeColor = ModernTextColor
        title.Text = "Искусственные сооружения"
        title.TextAlign = ContentAlignment.MiddleLeft
        title.Padding = New Padding(ScaleValue(9), 0, 0, 0)

        header.Controls.Add(title)
        header.Controls.Add(accent)
        Return header
    End Function

    Private Function CreateToolbar() As Control
        Dim toolbar As New TableLayoutPanel()
        toolbar.Name = "ModernPaletteToolbar"
        toolbar.Dock = DockStyle.Fill
        toolbar.BackColor = ModernCardColor
        toolbar.ColumnCount = 5
        toolbar.RowCount = 1
        toolbar.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        toolbar.Margin = New Padding(0, 0, 0, ScaleValue(8))
        toolbar.Padding = New Padding(0, ScaleValue(3), 0, ScaleValue(3))
        For index As Integer = 0 To 4
            toolbar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 20.0!))
        Next
        toolbar.Controls.Add(Button1, 0, 0)
        toolbar.Controls.Add(Button2, 1, 0)
        toolbar.Controls.Add(Button4, 2, 0)
        toolbar.Controls.Add(Button6, 3, 0)
        toolbar.Controls.Add(Button3, 4, 0)
        AddHandler toolbar.SizeChanged,
            Sub(sender As Object, e As EventArgs)
                UpdateToolbarCaptions(toolbar)
            End Sub
        For Each button As Button In PaletteButtons()
            AddHandler button.FontChanged,
                Sub(sender As Object, e As EventArgs)
                    UpdateToolbarCaptions(toolbar)
                End Sub
        Next
        UpdateToolbarCaptions(toolbar)
        Return toolbar
    End Function

    Private Function CreateSelectorCard() As Control
        Dim card As New TableLayoutPanel()
        card.Name = "ModernModelSelector"
        card.Dock = DockStyle.Fill
        card.BackColor = ModernCardColor
        card.ColumnCount = 1
        card.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        card.RowCount = 2
        card.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        card.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        card.Margin = New Padding(0, 0, 0, ScaleValue(8))
        card.Padding = New Padding(ScaleValue(10), ScaleValue(5), ScaleValue(10), ScaleValue(7))

        Label1.AutoSize = True
        Label1.Dock = DockStyle.Fill
        Label1.Font = modernBoldFont
        Label1.ForeColor = ModernTextColor
        Label1.Margin = Padding.Empty
        CBox_ListNamesArrProject.Dock = DockStyle.Fill
        CBox_ListNamesArrProject.Margin = New Padding(0, ScaleValue(3), 0, 0)
        CBox_ListNamesArrProject.BackColor = ModernCardColor
        CBox_ListNamesArrProject.ForeColor = ModernTextColor
        CBox_ListNamesArrProject.FlatStyle = FlatStyle.Flat

        card.Controls.Add(Label1, 0, 0)
        card.Controls.Add(CBox_ListNamesArrProject, 0, 1)
        Return card
    End Function

    Private Function CreatePaletteSplit() As SplitContainer
        Dim split As New SplitContainer()
        split.Name = "ModernPaletteSplit"
        split.Dock = DockStyle.Fill
        split.Orientation = Orientation.Horizontal
        split.BackColor = ModernCanvasColor
        split.BorderStyle = BorderStyle.None
        split.SplitterWidth = ScaleValue(6)
        split.Margin = Padding.Empty
        split.Panel1.Padding = New Padding(0, 0, 0, ScaleValue(3))
        split.Panel2.Padding = New Padding(0, ScaleValue(3), 0, 0)
        split.Panel1.Controls.Add(CreateContentCard("Структура сооружения", TreeView1))
        split.Panel2.Controls.Add(CreateContentCard("Свойства элемента", PropertyGrid1))
        Return split
    End Function

    Private Function CreateContentCard(caption As String, content As Control) As Control
        Dim card As New TableLayoutPanel()
        card.Dock = DockStyle.Fill
        card.BackColor = ModernCardColor
        card.ColumnCount = 1
        card.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        card.RowCount = 2
        card.RowStyles.Add(New RowStyle(SizeType.Absolute, ScaleValue(30)))
        card.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        card.Margin = Padding.Empty
        card.Padding = New Padding(ScaleValue(8), ScaleValue(4), ScaleValue(8), ScaleValue(8))

        Dim header As New Label()
        header.AutoSize = False
        header.Dock = DockStyle.Fill
        header.Font = modernBoldFont
        header.ForeColor = ModernTextColor
        header.Text = caption
        header.TextAlign = ContentAlignment.MiddleLeft
        header.Margin = Padding.Empty
        content.Dock = DockStyle.Fill
        content.Margin = Padding.Empty
        card.Controls.Add(header, 0, 0)
        card.Controls.Add(content, 0, 1)
        Return card
    End Function

    Private Sub ConfigureDisplayText()
        Label1.Text = "Модель проекта"
        СоздатьОсьОпорыToolStripMenuItem.Text = "Перерасчёт элементов опоры"
        ЭкспортЭлементовВDwgToolStripMenuItem.Text = "Экспорт элементов в DWG"
        ПоднятьОпуститьРядБалокToolStripMenuItem.Text = "Поднять/опустить ряд балок"
    End Sub

    Private Sub ConfigureButtons()
        ConfigureButton(Button1, "Обновить", "Обновить структуру проекта")
        ConfigureButton(Button2, "Пересчёт", "Пересчитать выбранное сооружение")
        ConfigureButton(Button4, "Балки", "Открыть команды для балок")
        ConfigureButton(Button6, "Опоры", "Открыть команды для опор")
        ConfigureButton(Button3, "Отчёты", "Сформировать отчёты или экспортировать данные")
    End Sub

    Private Sub ConfigureButton(button As Button, caption As String, hint As String)
        button.Dock = DockStyle.Fill
        button.Margin = New Padding(ScaleValue(2))
        button.Padding = Padding.Empty
        button.BackColor = ModernCardColor
        button.ForeColor = ModernTextColor
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.FlatAppearance.MouseOverBackColor = ModernAccentPaleColor
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(219, 232, 255)
        button.Font = modernRegularFont
        button.Text = caption
        button.AccessibleName = caption
        button.AccessibleDescription = hint
        button.TextAlign = ContentAlignment.BottomCenter
        button.ImageAlign = ContentAlignment.TopCenter
        button.TextImageRelation = TextImageRelation.ImageAboveText
        button.UseVisualStyleBackColor = False
        ToolTip1.SetToolTip(button, hint)
    End Sub

    Private Sub UpdateToolbarCaptions(toolbar As TableLayoutPanel)
        If modernToolbarCaptionsUpdating Then Return
        modernToolbarCaptionsUpdating = True
        Dim buttons As Button() = PaletteButtons()
        Dim captions As String() = PaletteCaptions()
        Dim showCaptions As Boolean = True
        For index As Integer = 0 To buttons.Length - 1
            buttons(index).Text = captions(index)
            buttons(index).ImageAlign = ContentAlignment.TopCenter
            Dim preferred As Size = buttons(index).GetPreferredSize(Size.Empty)
            If buttons(index).ClientSize.Width < preferred.Width OrElse
               buttons(index).ClientSize.Height < preferred.Height Then showCaptions = False
        Next
        For index As Integer = 0 To buttons.Length - 1
            buttons(index).Text = If(showCaptions, captions(index), String.Empty)
        Next
        Dim imageAlignment As ContentAlignment = If(showCaptions,
                                                    ContentAlignment.TopCenter,
                                                    ContentAlignment.MiddleCenter)
        For Each button As Button In buttons
            button.ImageAlign = imageAlignment
        Next
        modernToolbarCaptionsUpdating = False
    End Sub

    Private Function PaletteButtons() As Button()
        Return New Button() {Button1, Button2, Button4, Button6, Button3}
    End Function

    Private Function PaletteCaptions() As String()
        Return New String() {"Обновить", "Пересчёт", "Балки", "Опоры", "Отчёты"}
    End Function

    Private Sub ConfigureTree()
        TreeView1.BackColor = ModernCardColor
        TreeView1.ForeColor = ModernTextColor
        TreeView1.Font = modernRegularFont
        TreeView1.BorderStyle = BorderStyle.None
        TreeView1.ItemHeight = ScaleValue(24)
    End Sub

    Private Sub ConfigurePropertyGrid()
        PropertyGrid1.BackColor = ModernCardColor
        PropertyGrid1.ForeColor = ModernTextColor
        PropertyGrid1.Font = modernRegularFont
        PropertyGrid1.ViewBackColor = ModernCardColor
        PropertyGrid1.ViewForeColor = ModernTextColor
        PropertyGrid1.ViewBorderColor = ModernBorderColor
        PropertyGrid1.CategoryForeColor = ModernTextColor
        PropertyGrid1.CategorySplitterColor = ModernBorderColor
        PropertyGrid1.LineColor = ModernBorderColor
        PropertyGrid1.CommandsBackColor = ModernCardColor
        PropertyGrid1.CommandsForeColor = ModernTextColor
        PropertyGrid1.CommandsBorderColor = ModernBorderColor
        PropertyGrid1.HelpBackColor = Color.FromArgb(248, 250, 252)
        PropertyGrid1.HelpForeColor = ModernMutedColor
        PropertyGrid1.HelpBorderColor = ModernBorderColor
        PropertyGrid1.SelectedItemWithFocusBackColor = ModernAccentColor
        PropertyGrid1.SelectedItemWithFocusForeColor = Color.White
    End Sub

    Private Sub ConfigureMenus()
        For Each menu As ContextMenuStrip In New ContextMenuStrip() {
            ReportMenu, PillarsMenu, MenuBeams, UpdateStructureMenu, MenuSelectedNodeTree, BridgeMenu
        }
            menu.Font = modernRegularFont
            menu.BackColor = ModernCardColor
            menu.ForeColor = ModernTextColor
        Next
    End Sub

    Private Sub ReapplyModernVisualStyle()
        If modernVisualStyleApplying OrElse IsDisposed OrElse modernRegularFont Is Nothing Then Return
        modernVisualStyleApplying = True
        SuspendLayout()
        Try
            Font = modernRegularFont
            BackColor = ModernCanvasColor
            ForeColor = ModernTextColor
            ConfigureDisplayText()

            Dim scrollHost As Control = FindModernControl("ModernPaletteScroll")
            Dim canvas As Control = FindModernControl("ModernPaletteCanvas")
            Dim header As Control = FindModernControl("ModernPaletteHeader")
            Dim title As Label = TryCast(FindModernControl("ModernPaletteTitle"), Label)
            Dim toolbar As TableLayoutPanel = TryCast(FindModernControl("ModernPaletteToolbar"), TableLayoutPanel)
            Dim selectorCard As Control = FindModernControl("ModernModelSelector")
            Dim split As SplitContainer = TryCast(FindModernControl("ModernPaletteSplit"), SplitContainer)

            RestoreControlColors(scrollHost, ModernCanvasColor)
            RestoreControlColors(canvas, ModernCanvasColor)
            RestoreControlColors(header, ModernCardColor)
            RestoreControlColors(toolbar, ModernCardColor)
            RestoreControlColors(selectorCard, ModernCardColor)
            If title IsNot Nothing Then
                title.BackColor = ModernCardColor
                title.ForeColor = ModernTextColor
                title.Font = modernTitleFont
            End If
            If header IsNot Nothing Then
                For Each child As Control In header.Controls
                    Dim accent As Panel = TryCast(child, Panel)
                    If accent IsNot Nothing AndAlso accent.Dock = DockStyle.Left Then
                        accent.Width = ScaleValue(5)
                        accent.BackColor = ModernAccentColor
                    End If
                Next
            End If
            If split IsNot Nothing Then
                RestoreControlColors(split, ModernCanvasColor)
                RestoreControlColors(split.Panel1, ModernCanvasColor)
                RestoreControlColors(split.Panel2, ModernCanvasColor)
            End If

            Label1.BackColor = ModernCardColor
            Label1.ForeColor = ModernTextColor
            Label1.Font = modernBoldFont
            CBox_ListNamesArrProject.BackColor = ModernCardColor
            CBox_ListNamesArrProject.ForeColor = ModernTextColor
            CBox_ListNamesArrProject.Font = modernRegularFont
            CBox_ListNamesArrProject.FlatStyle = FlatStyle.Flat
            RestoreContentCard(TreeView1)
            RestoreContentCard(PropertyGrid1)
            ConfigureButtons()
            ConfigureTree()
            ConfigurePropertyGrid()
            ConfigureMenus()
            If toolbar IsNot Nothing Then UpdateToolbarCaptions(toolbar)
        Finally
            ResumeLayout(False)
            modernVisualStyleApplying = False
        End Try
    End Sub

    Private Sub RestoreControlColors(control As Control, background As Color)
        If control Is Nothing Then Return
        control.BackColor = background
        control.ForeColor = ModernTextColor
        control.Font = modernRegularFont
    End Sub

    Private Sub RestoreContentCard(content As Control)
        Dim card As Control = content.Parent
        If card Is Nothing Then Return
        RestoreControlColors(card, ModernCardColor)
        For Each child As Control In card.Controls
            Dim header As Label = TryCast(child, Label)
            If header IsNot Nothing Then
                header.BackColor = ModernCardColor
                header.ForeColor = ModernTextColor
                header.Font = modernBoldFont
            End If
        Next
    End Sub

    Private Function FindModernControl(name As String) As Control
        Dim matches As Control() = Controls.Find(name, True)
        If matches.Length = 0 Then Return Nothing
        Return matches(0)
    End Function

    Private Sub ConfigureModernStyleEvents(scrollHost As Panel,
                                           canvas As TableLayoutPanel,
                                           split As SplitContainer)
        AddHandler Load, AddressOf ModernVisualStyleLifecycleChanged
        AddHandler ParentChanged, AddressOf ModernVisualStyleLifecycleChanged
        AddHandler VisibleChanged, AddressOf ModernVisualStyleLifecycleChanged
        AddHandler HandleCreated, AddressOf ModernVisualStyleLifecycleChanged
        AddHandler HandleDestroyed, AddressOf ModernVisualStyleHandleDestroyed

        For Each control As Control In New Control() {
            Me, scrollHost, canvas, FindModernControl("ModernPaletteHeader"),
            FindModernControl("ModernPaletteTitle"), FindModernControl("ModernPaletteToolbar"),
            FindModernControl("ModernModelSelector"), split, split.Panel1, split.Panel2,
            TreeView1.Parent, PropertyGrid1.Parent, Label1, CBox_ListNamesArrProject,
            TreeView1, PropertyGrid1, Button1, Button2, Button4, Button6, Button3
        }
            If control IsNot Nothing Then
                AddHandler control.BackColorChanged, AddressOf ModernVisualStyleLifecycleChanged
                AddHandler control.ForeColorChanged, AddressOf ModernVisualStyleLifecycleChanged
                AddHandler control.FontChanged, AddressOf ModernVisualStyleLifecycleChanged
            End If
        Next
    End Sub

    Private Sub ModernVisualStyleLifecycleChanged(sender As Object, e As EventArgs)
        QueueModernVisualStyle()
    End Sub

    Private Sub ModernVisualStyleHandleDestroyed(sender As Object, e As EventArgs)
        modernVisualStyleGeneration += 1
        modernVisualStylePending = False
    End Sub

    Private Sub QueueModernVisualStyle()
        If modernVisualStyleApplying OrElse modernVisualStylePending OrElse IsDisposed OrElse
           Disposing OrElse Not IsHandleCreated Then Return

        modernVisualStylePending = True
        Dim generation As Long = modernVisualStyleGeneration
        Try
            BeginInvoke(New MethodInvoker(
                Sub()
                    Try
                        If generation <> modernVisualStyleGeneration OrElse
                           IsDisposed OrElse Disposing OrElse Not IsHandleCreated Then Return
                        ReapplyModernVisualStyle()
                    Finally
                        If generation = modernVisualStyleGeneration Then modernVisualStylePending = False
                    End Try
                End Sub))
        Catch ex As InvalidOperationException
            If generation = modernVisualStyleGeneration Then modernVisualStylePending = False
        End Try
    End Sub

    Private Sub ResizePaletteCanvas(scrollHost As Panel, canvas As TableLayoutPanel)
        Dim minimumCanvas As Size = ScaleSize(ModernMinimumCanvasSize)
        Dim availableWidth As Integer = scrollHost.ClientSize.Width
        Dim availableHeight As Integer = scrollHost.ClientSize.Height

        Dim targetSize As New Size(Math.Max(minimumCanvas.Width, availableWidth),
                                   Math.Max(minimumCanvas.Height, availableHeight))
        Dim verticalOffset As Integer = Math.Max(0, -scrollHost.AutoScrollPosition.Y)
        canvas.Size = targetSize
        If targetSize.Width <= availableWidth AndAlso scrollHost.AutoScrollPosition.X <> 0 Then
            scrollHost.AutoScrollPosition = New Point(0, verticalOffset)
        End If
    End Sub

    Private Sub QueuePaletteLayout(scrollHost As Panel, canvas As TableLayoutPanel)
        If modernPaletteLayoutPending OrElse IsDisposed OrElse scrollHost.IsDisposed OrElse
           canvas.IsDisposed OrElse Not scrollHost.IsHandleCreated Then Return

        modernPaletteLayoutPending = True
        Try
            scrollHost.BeginInvoke(New MethodInvoker(
                Sub()
                    Try
                        If IsDisposed OrElse scrollHost.IsDisposed OrElse canvas.IsDisposed Then Return
                        scrollHost.PerformLayout()
                        ResizePaletteCanvas(scrollHost, canvas)
                    Finally
                        modernPaletteLayoutPending = False
                    End Try
                End Sub))
        Catch ex As InvalidOperationException
            modernPaletteLayoutPending = False
        End Try
    End Sub

    Private Sub ConfigureSplitResize(split As SplitContainer)
        Dim initialized As Boolean = False
        Dim resizeSplit As EventHandler =
            Sub(sender As Object, e As EventArgs)
                Dim available As Integer = split.Height - split.SplitterWidth
                If available <= 0 Then Return
                split.Panel1MinSize = 0
                split.Panel2MinSize = 0
                Dim paneMinimum As Integer = Math.Min(ScaleValue(80), Math.Max(0, available \ 2))
                Dim distance As Integer
                If initialized Then
                    distance = Math.Max(paneMinimum, Math.Min(split.SplitterDistance, available - paneMinimum))
                Else
                    distance = CInt(Math.Round(available * 0.52R))
                    initialized = True
                End If
                split.SplitterDistance = distance
                split.Panel1MinSize = paneMinimum
                split.Panel2MinSize = paneMinimum
            End Sub
        AddHandler split.SizeChanged, resizeSplit
        resizeSplit(split, EventArgs.Empty)
    End Sub

    Private Function ScaleValue(logicalValue As Integer) As Integer
        Return CInt(Math.Round(logicalValue * DeviceDpi / 96.0R))
    End Function

    Private Function ScaleSize(logicalSize As Size) As Size
        Return New Size(ScaleValue(logicalSize.Width), ScaleValue(logicalSize.Height))
    End Function

    Private Sub DisposeModernResources(sender As Object, e As EventArgs)
        RemoveHandler Disposed, AddressOf DisposeModernResources
        If modernTitleFont IsNot Nothing Then modernTitleFont.Dispose()
        If modernBoldFont IsNot Nothing Then modernBoldFont.Dispose()
        If modernRegularFont IsNot Nothing Then modernRegularFont.Dispose()
        modernTitleFont = Nothing
        modernBoldFont = Nothing
        modernRegularFont = Nothing
    End Sub
End Class
