<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormCreateBeam
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormCreateBeam))
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CB_ListNumberRow = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CB_ListNumberProlet = New System.Windows.Forms.ComboBox()
        Me.CB_NameBeams = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CB_NameAlbums = New System.Windows.Forms.ComboBox()
        Me.NUpD_OffsetH = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NUpD_OffsetV = New System.Windows.Forms.NumericUpDown()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.NUpD_LenghtEndMonolit = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.NUpD_LenghtStartMonolit = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NUpD_OffsetH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NUpD_OffsetV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NUpD_LenghtEndMonolit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NUpD_LenghtStartMonolit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(214, 283)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 25)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "OK"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(5, 283)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 25)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Отмена"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.NUpD_LenghtEndMonolit)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.NUpD_LenghtStartMonolit)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.NUpD_OffsetV)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.NUpD_OffsetH)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.CB_ListNumberRow)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CB_ListNumberProlet)
        Me.GroupBox1.Controls.Add(Me.CB_NameBeams)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.CB_NameAlbums)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(325, 327)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Настройки"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(208, 36)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(93, 20)
        Me.Button4.TabIndex = 26
        Me.Button4.Text = "Указать отметку в районе точки опирания В, балки"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(94, 54)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Номер ряда"
        '
        'CB_ListNumberRow
        '
        Me.CB_ListNumberRow.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ListNumberRow.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_ListNumberRow.FormattingEnabled = True
        Me.CB_ListNumberRow.Location = New System.Drawing.Point(7, 46)
        Me.CB_ListNumberRow.Name = "CB_ListNumberRow"
        Me.CB_ListNumberRow.Size = New System.Drawing.Size(76, 21)
        Me.CB_ListNumberRow.TabIndex = 21
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(94, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Номер пролета"
        '
        'CB_ListNumberProlet
        '
        Me.CB_ListNumberProlet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ListNumberProlet.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_ListNumberProlet.FormattingEnabled = True
        Me.CB_ListNumberProlet.Location = New System.Drawing.Point(7, 19)
        Me.CB_ListNumberProlet.Name = "CB_ListNumberProlet"
        Me.CB_ListNumberProlet.Size = New System.Drawing.Size(76, 21)
        Me.CB_ListNumberProlet.TabIndex = 17
        '
        'CB_NameBeams
        '
        Me.CB_NameBeams.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_NameBeams.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_NameBeams.FormattingEnabled = True
        Me.CB_NameBeams.Location = New System.Drawing.Point(5, 247)
        Me.CB_NameBeams.Name = "CB_NameBeams"
        Me.CB_NameBeams.Size = New System.Drawing.Size(309, 21)
        Me.CB_NameBeams.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(122, 231)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Марка"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(133, 192)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Альбом"
        '
        'CB_NameAlbums
        '
        Me.CB_NameAlbums.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_NameAlbums.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_NameAlbums.FormattingEnabled = True
        Me.CB_NameAlbums.Location = New System.Drawing.Point(5, 208)
        Me.CB_NameAlbums.Name = "CB_NameAlbums"
        Me.CB_NameAlbums.Size = New System.Drawing.Size(309, 21)
        Me.CB_NameAlbums.TabIndex = 4
        '
        'NUpD_OffsetH
        '
        Me.NUpD_OffsetH.DecimalPlaces = 3
        Me.NUpD_OffsetH.Location = New System.Drawing.Point(6, 73)
        Me.NUpD_OffsetH.Maximum = New Decimal(New Integer() {10000000, 0, 0, 0})
        Me.NUpD_OffsetH.Minimum = New Decimal(New Integer() {100000000, 0, 0, -2147483648})
        Me.NUpD_OffsetH.Name = "NUpD_OffsetH"
        Me.NUpD_OffsetH.Size = New System.Drawing.Size(77, 20)
        Me.NUpD_OffsetH.TabIndex = 27
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(94, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(110, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Смещение от оси, м"
        '
        'NUpD_OffsetV
        '
        Me.NUpD_OffsetV.DecimalPlaces = 3
        Me.NUpD_OffsetV.Location = New System.Drawing.Point(7, 99)
        Me.NUpD_OffsetV.Maximum = New Decimal(New Integer() {10000000, 0, 0, 0})
        Me.NUpD_OffsetV.Name = "NUpD_OffsetV"
        Me.NUpD_OffsetV.Size = New System.Drawing.Size(76, 20)
        Me.NUpD_OffsetV.TabIndex = 29
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(94, 106)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(119, 13)
        Me.Label6.TabIndex = 30
        Me.Label6.Text = "Толщина покрытия, м"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(94, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(193, 13)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "Участок омоноличивания в конце, м"
        '
        'NUpD_LenghtEndMonolit
        '
        Me.NUpD_LenghtEndMonolit.DecimalPlaces = 3
        Me.NUpD_LenghtEndMonolit.Location = New System.Drawing.Point(7, 153)
        Me.NUpD_LenghtEndMonolit.Maximum = New Decimal(New Integer() {1000000000, 0, 0, 0})
        Me.NUpD_LenghtEndMonolit.Name = "NUpD_LenghtEndMonolit"
        Me.NUpD_LenghtEndMonolit.Size = New System.Drawing.Size(76, 20)
        Me.NUpD_LenghtEndMonolit.TabIndex = 33
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(94, 134)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(198, 13)
        Me.Label8.TabIndex = 32
        Me.Label8.Text = "Участок омоноличивания в начале, м"
        '
        'NUpD_LenghtStartMonolit
        '
        Me.NUpD_LenghtStartMonolit.DecimalPlaces = 3
        Me.NUpD_LenghtStartMonolit.Location = New System.Drawing.Point(6, 127)
        Me.NUpD_LenghtStartMonolit.Maximum = New Decimal(New Integer() {100000000, 0, 0, 0})
        Me.NUpD_LenghtStartMonolit.Name = "NUpD_LenghtStartMonolit"
        Me.NUpD_LenghtStartMonolit.Size = New System.Drawing.Size(77, 20)
        Me.NUpD_LenghtStartMonolit.TabIndex = 31
        '
        'FormCreateBeam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(209, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(344, 351)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(360, 390)
        Me.MinimumSize = New System.Drawing.Size(360, 390)
        Me.Name = "FormCreateBeam"
        Me.Text = "Создать новую балку"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NUpD_OffsetH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NUpD_OffsetV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NUpD_LenghtEndMonolit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NUpD_LenghtStartMonolit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents CB_ListNumberProlet As Windows.Forms.ComboBox
    Friend WithEvents CB_NameBeams As Windows.Forms.ComboBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents CB_NameAlbums As Windows.Forms.ComboBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents CB_ListNumberRow As Windows.Forms.ComboBox
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents NUpD_OffsetV As Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents NUpD_OffsetH As Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents NUpD_LenghtEndMonolit As Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents NUpD_LenghtStartMonolit As Windows.Forms.NumericUpDown
End Class
