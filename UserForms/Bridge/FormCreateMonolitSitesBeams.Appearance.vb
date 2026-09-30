Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class FormCreateMonolitSitesBeams
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
        AutoScrollMinSize = ModernScaleSize(New Size(700, 500))
        StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConfigureModernWindowSize()
        Dim workArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim safeWidth As Integer = Math.Max(1, workArea.Width - ModernScale(32))
        Dim safeHeight As Integer = Math.Max(1, workArea.Height - ModernScale(32))
        Dim preferred As Size = ModernScaleSize(New Size(900, 650))
        Dim requestedMinimum As Size = ModernScaleSize(New Size(760, 560))
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
            .ForeColor = ModernTextColor, .Text = "Участки омоноличивания",
            .TextAlign = ContentAlignment.BottomLeft
        }
        Dim subtitle As New Label() With {
            .Name = "ModernSubtitle", .AutoSize = False, .Font = modernRegularFont,
            .ForeColor = ModernMutedTextColor,
            .Text = "Выберите сооружение, форму участка и область построения",
            .TextAlign = ContentAlignment.TopLeft
        }
        modernHeader.Controls.Add(accent)
        modernHeader.Controls.Add(title)
        modernHeader.Controls.Add(subtitle)
        modernFooter = New Panel() With {.Name = "ModernFooter", .BackColor = ModernCardColor}
        Controls.Add(modernHeader)
        Controls.Add(modernFooter)
        modernFooter.Controls.Add(Button1)
        modernFooter.Controls.Add(Button2)
        modernHeader.SendToBack()
        modernFooter.SendToBack()
    End Sub

    Private Sub ConfigureModernControls()
        Panel1.BackColor = ModernCardColor
        Panel1.BorderStyle = BorderStyle.FixedSingle
        Panel2.BackColor = ModernCardColor
        Panel2.BorderStyle = BorderStyle.FixedSingle
        ConfigureCard(GroupBox1, "Форма участка омоноличивания")
        ConfigureCard(GroupBox2, "Выбор пролёта сооружения")
        Label7.Font = modernBoldFont
        Label7.ForeColor = ModernTextColor
        Label7.BackColor = Color.Transparent
        Label8.Font = modernBoldFont
        Label8.ForeColor = ModernTextColor
        Label8.BackColor = Color.Transparent

        For Each fieldLabel As Label In New Label() {Label1, Label2, Label3, Label5}
            StyleFieldLabel(fieldLabel)
        Next
        For Each input As ComboBox In New ComboBox() {CBox_ListNamesArrProject, CBox_ListNamesBridge,
                                                       CBox_ListNamesTemplateXML, CBox_ListProlet}
            input.Font = modernRegularFont
            input.ForeColor = ModernTextColor
            input.BackColor = ModernCardColor
            input.FlatStyle = FlatStyle.Flat
        Next
        For Each radio As RadioButton In New RadioButton() {RadioButton1, RadioButton2, RadioButton3,
                                                             RadioButton4, RadioButton5}
            radio.AutoSize = False
            radio.Font = modernRegularFont
            radio.ForeColor = ModernTextColor
            radio.BackColor = ModernCardColor
            radio.UseVisualStyleBackColor = False
        Next
        StyleSecondaryButton(Button1)
        StylePrimaryButton(Button2)
        Button1.Text = "Отмена"
        Button2.Text = "Создать участки"
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
            Dim gap As Integer = ModernScale(12)
            Dim clientWidth As Integer = Math.Max(ModernScale(700), ClientSize.Width)
            Dim clientHeight As Integer = Math.Max(ModernScale(500), ClientSize.Height)
            modernHeader.SetBounds(margin, ModernScale(12), clientWidth - margin * 2, ModernScale(72))
            LayoutModernHeader()
            Dim footerTop As Integer = clientHeight - ModernScale(68)
            modernFooter.SetBounds(0, footerTop, clientWidth, ModernScale(68))
            Dim bodyTop As Integer = ModernScale(124)
            Dim bodyHeight As Integer = Math.Max(ModernScale(220), footerTop - bodyTop - ModernScale(12))
            Dim availableWidth As Integer = clientWidth - margin * 2
            Dim leftWidth As Integer = Math.Max(ModernScale(300),
                CInt(Math.Floor((availableWidth - gap) * 0.42R)))
            Dim rightWidth As Integer = availableWidth - gap - leftWidth
            Label7.SetBounds(margin, ModernScale(96), leftWidth, ModernScale(24))
            Label8.SetBounds(margin + leftWidth + gap, ModernScale(96), rightWidth, ModernScale(24))
            Panel1.SetBounds(margin, bodyTop, leftWidth, bodyHeight)
            Panel2.SetBounds(margin + leftWidth + gap, bodyTop, rightWidth, bodyHeight)
            LayoutProjectCard()
            LayoutParameterCard()
            LayoutFooterActions()
            For Each control As Control In New Control() {Label7, Label8, Panel1, Panel2, Button1, Button2}
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

    Private Sub LayoutProjectCard()
        Dim left As Integer = ModernScale(12)
        Dim width As Integer = Math.Max(ModernScale(270), Panel1.ClientSize.Width - left * 2)
        Label1.Location = New Point(left, ModernScale(18))
        CBox_ListNamesArrProject.SetBounds(left, ModernScale(42), width, ModernScale(26))
        Label2.Location = New Point(left, ModernScale(82))
        CBox_ListNamesBridge.SetBounds(left, ModernScale(106), width, ModernScale(26))
        Label3.Location = New Point(left, ModernScale(146))
        CBox_ListNamesTemplateXML.SetBounds(left, ModernScale(170), width, ModernScale(26))
    End Sub

    Private Sub LayoutParameterCard()
        Dim left As Integer = ModernScale(12)
        Dim width As Integer = Math.Max(ModernScale(260), Panel2.ClientSize.Width - left * 2)
        GroupBox1.SetBounds(left, ModernScale(12), width, ModernScale(98))
        GroupBox2.SetBounds(left, ModernScale(122), width,
                            Math.Max(ModernScale(120), Panel2.ClientSize.Height - ModernScale(134)))
        Dim radioWidth As Integer = Math.Max(ModernScale(220), GroupBox1.ClientSize.Width - ModernScale(24))
        RadioButton4.SetBounds(ModernScale(12), ModernScale(24), radioWidth, ModernScale(24))
        RadioButton5.SetBounds(ModernScale(12), ModernScale(56), radioWidth, ModernScale(24))
        radioWidth = Math.Max(ModernScale(220), GroupBox2.ClientSize.Width - ModernScale(24))
        RadioButton1.SetBounds(ModernScale(12), ModernScale(24), radioWidth, ModernScale(42))
        RadioButton2.SetBounds(ModernScale(12), ModernScale(72), radioWidth, ModernScale(24))
        Dim comboWidth As Integer = Math.Max(ModernScale(110), CInt(Math.Floor(radioWidth * 0.42R)))
        CBox_ListProlet.SetBounds(ModernScale(12), ModernScale(102), comboWidth, ModernScale(26))
        Label5.Location = New Point(CBox_ListProlet.Right + ModernScale(12), ModernScale(107))
        RadioButton3.SetBounds(ModernScale(12), ModernScale(136), radioWidth, ModernScale(24))
    End Sub

    Private Sub LayoutFooterActions()
        Dim margin As Integer = ModernScale(16)
        Dim y As Integer = ModernScale(14)
        Button1.SetBounds(margin, y, ModernScale(116), ModernScale(40))
        Button2.SetBounds(modernFooter.ClientSize.Width - margin - ModernScale(176), y,
                          ModernScale(176), ModernScale(40))
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
