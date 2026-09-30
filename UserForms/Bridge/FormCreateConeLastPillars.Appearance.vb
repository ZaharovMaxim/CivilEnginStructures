Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class FormCreateConeLastPillars
    Private modernAppearanceApplied As Boolean
    Private modernLayoutInProgress As Boolean
    Private modernRegularFont As Font
    Private modernBoldFont As Font
    Private modernTitleFont As Font
    Private modernHeader As Panel
    Private modernFooter As Panel

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
            ConfigureModernControls()
            ConfigureModernWindowSize()
            AddHandler SizeChanged, AddressOf ModernSizeChanged
            AddHandler Shown, AddressOf ModernShown
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
        Dim preferred As Size = ModernScaleSize(New Size(1120, 800))
        Dim requestedMinimum As Size = ModernScaleSize(New Size(998, 680))
        MinimumSize = New Size(Math.Min(requestedMinimum.Width, safeWidth),
                               Math.Min(requestedMinimum.Height, safeHeight))
        Size = New Size(Math.Max(MinimumSize.Width, Math.Min(preferred.Width, safeWidth)),
                        Math.Max(MinimumSize.Height, Math.Min(preferred.Height, safeHeight)))
    End Sub

    Private Sub CreateModernChrome()
        modernHeader = New Panel() With {.Name = "ModernHeader", .BackColor = ModernCardColor}
        Dim accent As New Panel() With {.Name = "ModernHeaderAccent", .BackColor = ModernAccentColor}
        Dim title As New Label() With {
            .Name = "ModernTitle", .AutoSize = False, .Font = modernTitleFont,
            .ForeColor = ModernTextColor, .Text = "Конус мостового сооружения",
            .TextAlign = ContentAlignment.BottomLeft
        }
        Dim subtitle As New Label() With {
            .Name = "ModernSubtitle", .AutoSize = False, .Font = modernRegularFont,
            .ForeColor = ModernMutedTextColor,
            .Text = "Настройте поверхности, крайнюю опору и параметры построения",
            .TextAlign = ContentAlignment.TopLeft
        }
        modernHeader.Controls.Add(accent)
        modernHeader.Controls.Add(title)
        modernHeader.Controls.Add(subtitle)
        modernFooter = New Panel() With {.Name = "ModernFooter", .BackColor = ModernCardColor}
        Controls.Add(modernHeader)
        Controls.Add(modernFooter)
        modernFooter.Controls.Add(Button2)
        modernFooter.Controls.Add(Button7)
        modernFooter.Controls.Add(Button8)
        modernHeader.SendToBack()
        modernFooter.SendToBack()
    End Sub

    Private Sub ConfigureModernControls()
        ConfigureCard(GroupBox1, "Настройки проекта")
        ConfigureCard(GroupBox2, "Настройки конуса")
        For Each fieldLabel As Label In New Label() {Label12, Label15, Label14, Label2,
                                                      Label13, Label16, Label26, Label4}
            StyleFieldLabel(fieldLabel)
        Next
        For Each input As Control In New Control() {CBox_ListNamesArrProject, CBox_ListNamesBridge,
                                                     CB_NameAlignment, CBox_ListNamesTemplateXML,
                                                     CB_ProjectSurface, CB_EgSurface, CB_NumberPillar,
                                                     NUpD_CountSegmets}
            StyleInput(input)
        Next
        ChB_ProjectSurfaceInAlignment.AutoSize = True
        ChB_ProjectSurfaceInAlignment.Font = modernRegularFont
        ChB_ProjectSurfaceInAlignment.ForeColor = ModernTextColor
        ChB_ProjectSurfaceInAlignment.BackColor = ModernCardColor
        ChB_ProjectSurfaceInAlignment.UseVisualStyleBackColor = False
        Label17.Visible = False
        Label5.Visible = False
        CB_NameSites.Visible = False
        Button1.Visible = False
        ConfigureGrid()
        StyleSecondaryButton(Button2)
        StyleSecondaryButton(Button7)
        StylePrimaryButton(Button8)
        Button2.Text = "Отмена"
        Button7.Text = "Расчёт"
        Button8.Text = "Создать конус"
    End Sub

    Private Sub ConfigureCard(card As GroupBox, caption As String)
        card.Text = caption
        card.Font = modernBoldFont
        card.ForeColor = ModernTextColor
        card.BackColor = ModernCardColor
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
    End Sub

    Private Sub ConfigureGrid()
        DGV_PropertiesCone.Font = modernRegularFont
        DGV_PropertiesCone.BackgroundColor = ModernCardColor
        DGV_PropertiesCone.BorderStyle = BorderStyle.FixedSingle
        DGV_PropertiesCone.GridColor = ModernBorderColor
        DGV_PropertiesCone.EnableHeadersVisualStyles = False
        DGV_PropertiesCone.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        DGV_PropertiesCone.ColumnHeadersDefaultCellStyle.BackColor = ModernAccentPaleColor
        DGV_PropertiesCone.ColumnHeadersDefaultCellStyle.ForeColor = ModernTextColor
        DGV_PropertiesCone.ColumnHeadersDefaultCellStyle.Font = modernBoldFont
        DGV_PropertiesCone.ColumnHeadersDefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        DGV_PropertiesCone.ColumnHeadersDefaultCellStyle.SelectionForeColor = ModernTextColor
        DGV_PropertiesCone.DefaultCellStyle.BackColor = ModernCardColor
        DGV_PropertiesCone.DefaultCellStyle.ForeColor = ModernTextColor
        DGV_PropertiesCone.DefaultCellStyle.SelectionBackColor = ModernAccentPaleColor
        DGV_PropertiesCone.DefaultCellStyle.SelectionForeColor = ModernTextColor
        DGV_PropertiesCone.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DGV_PropertiesCone.ColumnHeadersHeight = ModernScale(36)
        DGV_PropertiesCone.RowTemplate.Height = ModernScale(30)
        DGV_PropertiesCone.RowHeadersVisible = False
        DGV_PropertiesCone.ScrollBars = ScrollBars.Vertical
        DGV_PropertiesCone.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        Номер.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Номер.MinimumWidth = ModernScale(360)
        Column1.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Column1.Width = ModernScale(140)
    End Sub

    Private Sub StylePrimaryButton(button As Button)
        button.Font = modernBoldFont
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.BackColor = ModernAccentColor
        button.ForeColor = Color.White
        button.UseVisualStyleBackColor = False
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
        LayoutModernAppearance()
    End Sub

    Private Sub LayoutModernAppearance()
        If Not modernAppearanceApplied OrElse modernLayoutInProgress OrElse modernHeader Is Nothing Then Return
        modernLayoutInProgress = True
        SuspendLayout()
        Try
            Dim margin As Integer = ModernScale(16)
            Dim clientWidth As Integer = Math.Max(ModernScale(760), ClientSize.Width)
            Dim clientHeight As Integer = Math.Max(ModernScale(600), ClientSize.Height)
            modernHeader.SetBounds(margin, ModernScale(12), clientWidth - margin * 2, ModernScale(72))
            LayoutModernHeader()
            GroupBox1.SetBounds(margin, ModernScale(96), clientWidth - margin * 2, ModernScale(158))
            LayoutProjectControls()
            Dim footerTop As Integer = clientHeight - ModernScale(68)
            modernFooter.SetBounds(0, footerTop, clientWidth, ModernScale(68))
            LayoutFooterActions()
            GroupBox2.SetBounds(margin, ModernScale(266), clientWidth - margin * 2,
                                Math.Max(ModernScale(220), footerTop - ModernScale(278)))
            LayoutConeControls()
            For Each control As Control In New Control() {GroupBox1, GroupBox2, Button2, Button7, Button8}
                control.BringToFront()
            Next
        Finally
            ResumeLayout(False)
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
        Dim columnWidth As Integer = Math.Max(ModernScale(150),
            (GroupBox1.ClientSize.Width - left * 2 - gap * 3) \ 4)
        Dim firstLabels As Label() = {Label12, Label15, Label14, Label2}
        Dim firstInputs As Control() = {CBox_ListNamesArrProject, CBox_ListNamesBridge,
                                        CB_NameAlignment, CBox_ListNamesTemplateXML}
        Dim secondLabels As Label() = {Label13, Label16, Label26}
        Dim secondInputs As Control() = {CB_ProjectSurface, CB_EgSurface, CB_NumberPillar}
        For index As Integer = 0 To 3
            Dim x As Integer = left + index * (columnWidth + gap)
            firstLabels(index).Location = New Point(x, ModernScale(20))
            firstInputs(index).SetBounds(x, ModernScale(44), columnWidth, ModernScale(26))
            If index < secondLabels.Length Then
                secondLabels(index).Location = New Point(x, ModernScale(72))
                secondInputs(index).SetBounds(x, ModernScale(96), columnWidth, ModernScale(26))
            End If
        Next
        ChB_ProjectSurfaceInAlignment.Location = New Point(left, ModernScale(128))
    End Sub

    Private Sub LayoutConeControls()
        Dim left As Integer = ModernScale(12)
        Label4.Location = New Point(left, ModernScale(24))
        NUpD_CountSegmets.SetBounds(ModernScale(190), ModernScale(20), ModernScale(92), ModernScale(26))
        DGV_PropertiesCone.SetBounds(left, ModernScale(60),
                                     Math.Max(ModernScale(380), GroupBox2.ClientSize.Width - left * 2),
                                     Math.Max(ModernScale(180), GroupBox2.ClientSize.Height - ModernScale(72)))
    End Sub

    Private Sub LayoutFooterActions()
        Dim margin As Integer = ModernScale(16)
        Dim y As Integer = ModernScale(14)
        Button2.SetBounds(margin, y, ModernScale(116), ModernScale(40))
        Button8.SetBounds(modernFooter.ClientSize.Width - margin - ModernScale(184), y,
                          ModernScale(184), ModernScale(40))
        Button7.SetBounds(Button8.Left - ModernScale(12) - ModernScale(116), y,
                          ModernScale(116), ModernScale(40))
    End Sub

    Private Function ModernScale(logicalValue As Integer) As Integer
        Return CInt(Math.Round(logicalValue * Math.Max(1.0F, CSng(DeviceDpi) / 96.0F)))
    End Function

    Private Function ModernScaleSize(logicalSize As Size) As Size
        Return New Size(ModernScale(logicalSize.Width), ModernScale(logicalSize.Height))
    End Function

    Private Sub DisposeModernAppearanceResources(sender As Object, e As EventArgs)
        If modernRegularFont IsNot Nothing Then modernRegularFont.Dispose()
        If modernBoldFont IsNot Nothing Then modernBoldFont.Dispose()
        If modernTitleFont IsNot Nothing Then modernTitleFont.Dispose()
    End Sub
End Class
