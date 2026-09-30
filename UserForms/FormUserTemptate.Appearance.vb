Imports System.Drawing
Imports System.Windows.Forms

Partial Public Class FormUserTemptate
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
        FormBorderStyle = FormBorderStyle.Sizable
        DoubleBuffered = True
        MaximumSize = Size.Empty
        Padding = Padding.Empty
        AutoScroll = True
        AutoScrollMinSize = ModernScaleSize(New Size(560, 400))
        StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ConfigureModernWindowSize()
        Dim workArea As Rectangle = Screen.FromControl(Me).WorkingArea
        Dim safeWidth As Integer = Math.Max(1, workArea.Width - ModernScale(32))
        Dim safeHeight As Integer = Math.Max(1, workArea.Height - ModernScale(32))
        Dim preferred As Size = ModernScaleSize(New Size(720, 560))
        Dim requestedMinimum As Size = ModernScaleSize(New Size(600, 460))
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
            .ForeColor = ModernTextColor, .Text = "Библиотека схем",
            .TextAlign = ContentAlignment.BottomLeft
        }
        Dim subtitle As New Label() With {
            .Name = "ModernSubtitle", .AutoSize = False, .Font = modernRegularFont,
            .ForeColor = ModernMutedTextColor,
            .Text = "Выберите сохранённую схему или задайте имя новой",
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
        modernFooter.Controls.Add(Button3)
        modernFooter.Controls.Add(Label3)
        modernHeader.SendToBack()
        modernFooter.SendToBack()
    End Sub

    Private Sub ConfigureModernControls()
        GroupBox1.Text = "Наименование схемы"
        GroupBox1.Font = modernBoldFont
        GroupBox1.ForeColor = ModernTextColor
        GroupBox1.BackColor = ModernCardColor
        For Each fieldLabel As Label In New Label() {Label1, Label2}
            fieldLabel.AutoSize = True
            fieldLabel.Font = modernRegularFont
            fieldLabel.ForeColor = ModernMutedTextColor
            fieldLabel.BackColor = Color.Transparent
        Next
        TextBox1.Font = modernRegularFont
        TextBox1.ForeColor = ModernTextColor
        TextBox1.BackColor = ModernCardColor
        TextBox1.BorderStyle = BorderStyle.FixedSingle
        ListBox1.Font = modernRegularFont
        ListBox1.ForeColor = ModernTextColor
        ListBox1.BackColor = ModernCardColor
        ListBox1.BorderStyle = BorderStyle.FixedSingle
        ListBox1.IntegralHeight = False
        Label3.AutoSize = False
        Label3.AutoEllipsis = True
        Label3.Font = modernRegularFont
        Label3.ForeColor = ModernMutedTextColor
        Label3.BackColor = Color.Transparent
        Label3.TextAlign = ContentAlignment.MiddleLeft
        StylePrimaryButton(Button1)
        StyleSecondaryButton(Button2)
        StyleSecondaryButton(Button3)
        Button1.Text = "Записать"
        Button2.Text = "Удалить"
        Button3.Text = "Закрыть"
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
            Dim clientWidth As Integer = Math.Max(ModernScale(560), ClientSize.Width)
            Dim clientHeight As Integer = Math.Max(ModernScale(400), ClientSize.Height)
            modernHeader.SetBounds(margin, ModernScale(12), clientWidth - margin * 2, ModernScale(72))
            LayoutModernHeader()
            Dim footerTop As Integer = clientHeight - ModernScale(68)
            modernFooter.SetBounds(0, footerTop, clientWidth, ModernScale(68))
            GroupBox1.SetBounds(margin, ModernScale(88), clientWidth - margin * 2,
                                Math.Max(ModernScale(260), footerTop - ModernScale(88)))
            LayoutLibraryCard()
            LayoutFooterActions()
            For Each control As Control In New Control() {GroupBox1, Label3, Button1, Button2, Button3}
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

    Private Sub LayoutLibraryCard()
        Dim left As Integer = ModernScale(12)
        Dim width As Integer = Math.Max(ModernScale(300), GroupBox1.ClientSize.Width - left * 2)
        Label1.Location = New Point(left, ModernScale(26))
        TextBox1.SetBounds(ModernScale(80), ModernScale(20),
                           Math.Max(ModernScale(230), width - ModernScale(68)), ModernScale(28))
        Label2.Location = New Point(left, ModernScale(52))
        ListBox1.SetBounds(left, ModernScale(72), width,
                           Math.Max(ModernScale(180), GroupBox1.ClientSize.Height - ModernScale(84)))
    End Sub

    Private Sub LayoutFooterActions()
        Dim margin As Integer = ModernScale(16)
        Dim gap As Integer = ModernScale(10)
        Dim y As Integer = ModernScale(14)
        Dim primaryWidth As Integer = ModernScale(126)
        Dim secondaryWidth As Integer = ModernScale(108)
        Button1.SetBounds(modernFooter.ClientSize.Width - margin - primaryWidth, y, primaryWidth, ModernScale(40))
        Button2.SetBounds(Button1.Left - gap - secondaryWidth, y, secondaryWidth, ModernScale(40))
        Button3.SetBounds(Button2.Left - gap - secondaryWidth, y, secondaryWidth, ModernScale(40))
        Label3.SetBounds(margin, y, Math.Max(ModernScale(120), Button3.Left - gap - margin), ModernScale(40))
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
