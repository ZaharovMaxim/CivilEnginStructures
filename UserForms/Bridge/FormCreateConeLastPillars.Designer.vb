<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormCreateConeLastPillars
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.CB_NumberPillar = New System.Windows.Forms.ComboBox()
        Me.CBox_ListNamesBridge = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.CB_EgSurface = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CBox_ListNamesTemplateXML = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.ChB_ProjectSurfaceInAlignment = New System.Windows.Forms.CheckBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.CB_NameAlignment = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.CB_ProjectSurface = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.CBox_ListNamesArrProject = New System.Windows.Forms.ComboBox()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CB_NameSites = New System.Windows.Forms.ComboBox()
        Me.DGV_PropertiesCone = New System.Windows.Forms.DataGridView()
        Me.Номер = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.NUpD_CountSegmets = New System.Windows.Forms.NumericUpDown()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DGV_PropertiesCone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NUpD_CountSegmets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.GroupBox1.Controls.Add(Me.Label26)
        Me.GroupBox1.Controls.Add(Me.CB_NumberPillar)
        Me.GroupBox1.Controls.Add(Me.CBox_ListNamesBridge)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.CB_EgSurface)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CBox_ListNamesTemplateXML)
        Me.GroupBox1.Controls.Add(Me.Label17)
        Me.GroupBox1.Controls.Add(Me.ChB_ProjectSurfaceInAlignment)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.CB_NameAlignment)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.CB_ProjectSurface)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.CBox_ListNamesArrProject)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(10, 10)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(430, 230)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Настройки проекта"
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(124, 208)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(83, 13)
        Me.Label26.TabIndex = 48
        Me.Label26.Text = "Крайняя опора"
        '
        'CB_NumberPillar
        '
        Me.CB_NumberPillar.FormattingEnabled = True
        Me.CB_NumberPillar.Location = New System.Drawing.Point(5, 204)
        Me.CB_NumberPillar.Name = "CB_NumberPillar"
        Me.CB_NumberPillar.Size = New System.Drawing.Size(114, 21)
        Me.CB_NumberPillar.TabIndex = 47
        '
        'CBox_ListNamesBridge
        '
        Me.CBox_ListNamesBridge.DropDownHeight = 100
        Me.CBox_ListNamesBridge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBox_ListNamesBridge.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CBox_ListNamesBridge.FormattingEnabled = True
        Me.CBox_ListNamesBridge.IntegralHeight = False
        Me.CBox_ListNamesBridge.Location = New System.Drawing.Point(5, 48)
        Me.CBox_ListNamesBridge.Margin = New System.Windows.Forms.Padding(2)
        Me.CBox_ListNamesBridge.Name = "CBox_ListNamesBridge"
        Me.CBox_ListNamesBridge.Size = New System.Drawing.Size(217, 21)
        Me.CBox_ListNamesBridge.TabIndex = 46
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(227, 104)
        Me.Label16.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(184, 13)
        Me.Label16.TabIndex = 44
        Me.Label16.Text = "Фактическая поверхность (земля)"
        '
        'CB_EgSurface
        '
        Me.CB_EgSurface.DropDownHeight = 100
        Me.CB_EgSurface.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_EgSurface.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_EgSurface.FormattingEnabled = True
        Me.CB_EgSurface.IntegralHeight = False
        Me.CB_EgSurface.Location = New System.Drawing.Point(5, 98)
        Me.CB_EgSurface.Margin = New System.Windows.Forms.Padding(2)
        Me.CB_EgSurface.Name = "CB_EgSurface"
        Me.CB_EgSurface.Size = New System.Drawing.Size(217, 21)
        Me.CB_EgSurface.TabIndex = 43
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(227, 55)
        Me.Label15.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(147, 13)
        Me.Label15.TabIndex = 42
        Me.Label15.Text = "Наименование сооружения"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(227, 156)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(164, 13)
        Me.Label2.TabIndex = 40
        Me.Label2.Text = "Шаблон оформления чертежей"
        '
        'CBox_ListNamesTemplateXML
        '
        Me.CBox_ListNamesTemplateXML.DropDownHeight = 100
        Me.CBox_ListNamesTemplateXML.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBox_ListNamesTemplateXML.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CBox_ListNamesTemplateXML.FormattingEnabled = True
        Me.CBox_ListNamesTemplateXML.IntegralHeight = False
        Me.CBox_ListNamesTemplateXML.Location = New System.Drawing.Point(4, 150)
        Me.CBox_ListNamesTemplateXML.Name = "CBox_ListNamesTemplateXML"
        Me.CBox_ListNamesTemplateXML.Size = New System.Drawing.Size(217, 21)
        Me.CBox_ListNamesTemplateXML.TabIndex = 39
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(65, 216)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(0, 13)
        Me.Label17.TabIndex = 34
        '
        'ChB_ProjectSurfaceInAlignment
        '
        Me.ChB_ProjectSurfaceInAlignment.AutoSize = True
        Me.ChB_ProjectSurfaceInAlignment.Location = New System.Drawing.Point(4, 182)
        Me.ChB_ProjectSurfaceInAlignment.Margin = New System.Windows.Forms.Padding(2)
        Me.ChB_ProjectSurfaceInAlignment.Name = "ChB_ProjectSurfaceInAlignment"
        Me.ChB_ProjectSurfaceInAlignment.Size = New System.Drawing.Size(235, 17)
        Me.ChB_ProjectSurfaceInAlignment.TabIndex = 29
        Me.ChB_ProjectSurfaceInAlignment.Text = "Взять проектную поверхность из трассы"
        Me.ChB_ProjectSurfaceInAlignment.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(227, 130)
        Me.Label14.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(187, 13)
        Me.Label14.TabIndex = 28
        Me.Label14.Text = "Ось трассы автомобильной дороги"
        '
        'CB_NameAlignment
        '
        Me.CB_NameAlignment.DropDownHeight = 100
        Me.CB_NameAlignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_NameAlignment.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_NameAlignment.FormattingEnabled = True
        Me.CB_NameAlignment.IntegralHeight = False
        Me.CB_NameAlignment.Location = New System.Drawing.Point(5, 123)
        Me.CB_NameAlignment.Margin = New System.Windows.Forms.Padding(2)
        Me.CB_NameAlignment.Name = "CB_NameAlignment"
        Me.CB_NameAlignment.Size = New System.Drawing.Size(217, 21)
        Me.CB_NameAlignment.TabIndex = 27
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(227, 78)
        Me.Label13.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(129, 13)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "Проектная поверхность"
        '
        'CB_ProjectSurface
        '
        Me.CB_ProjectSurface.DropDownHeight = 100
        Me.CB_ProjectSurface.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ProjectSurface.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CB_ProjectSurface.FormattingEnabled = True
        Me.CB_ProjectSurface.IntegralHeight = False
        Me.CB_ProjectSurface.Location = New System.Drawing.Point(5, 72)
        Me.CB_ProjectSurface.Margin = New System.Windows.Forms.Padding(2)
        Me.CB_ProjectSurface.Name = "CB_ProjectSurface"
        Me.CB_ProjectSurface.Size = New System.Drawing.Size(217, 21)
        Me.CB_ProjectSurface.TabIndex = 25
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(227, 31)
        Me.Label12.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(92, 13)
        Me.Label12.TabIndex = 23
        Me.Label12.Text = "Имя подобъекта"
        '
        'CBox_ListNamesArrProject
        '
        Me.CBox_ListNamesArrProject.DropDownHeight = 100
        Me.CBox_ListNamesArrProject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBox_ListNamesArrProject.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CBox_ListNamesArrProject.FormattingEnabled = True
        Me.CBox_ListNamesArrProject.IntegralHeight = False
        Me.CBox_ListNamesArrProject.Location = New System.Drawing.Point(5, 23)
        Me.CBox_ListNamesArrProject.Margin = New System.Windows.Forms.Padding(2)
        Me.CBox_ListNamesArrProject.Name = "CBox_ListNamesArrProject"
        Me.CBox_ListNamesArrProject.Size = New System.Drawing.Size(217, 21)
        Me.CBox_ListNamesArrProject.TabIndex = 22
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(173, 614)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(110, 37)
        Me.Button7.TabIndex = 10
        Me.Button7.Text = "Расчет"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(331, 614)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(109, 37)
        Me.Button8.TabIndex = 9
        Me.Button8.Text = "OK"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(10, 614)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(119, 37)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.CB_NameSites)
        Me.GroupBox2.Controls.Add(Me.DGV_PropertiesCone)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.NUpD_CountSegmets)
        Me.GroupBox2.Location = New System.Drawing.Point(10, 248)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(430, 370)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Настройки конуса"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(93, 19)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 13)
        Me.Label5.TabIndex = 50
        Me.Label5.Text = "Имя площадки"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(340, 36)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 20)
        Me.Button1.TabIndex = 49
        Me.Button1.Text = "Создать"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CB_NameSites
        '
        Me.CB_NameSites.FormattingEnabled = True
        Me.CB_NameSites.Location = New System.Drawing.Point(4, 37)
        Me.CB_NameSites.Name = "CB_NameSites"
        Me.CB_NameSites.Size = New System.Drawing.Size(329, 21)
        Me.CB_NameSites.TabIndex = 7
        '
        'DGV_PropertiesCone
        '
        Me.DGV_PropertiesCone.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_PropertiesCone.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Номер, Me.Column1})
        Me.DGV_PropertiesCone.Location = New System.Drawing.Point(0, 100)
        Me.DGV_PropertiesCone.Name = "DGV_PropertiesCone"
        Me.DGV_PropertiesCone.RowHeadersVisible = False
        Me.DGV_PropertiesCone.Size = New System.Drawing.Size(432, 239)
        Me.DGV_PropertiesCone.TabIndex = 6
        '
        'Номер
        '
        Me.Номер.HeaderText = "Параметр"
        Me.Номер.Name = "Номер"
        Me.Номер.Width = 300
        '
        'Column1
        '
        Me.Column1.HeaderText = "Значение"
        Me.Column1.Name = "Column1"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(73, 66)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(123, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Количество сегментов"
        '
        'NUpD_CountSegmets
        '
        Me.NUpD_CountSegmets.Location = New System.Drawing.Point(6, 64)
        Me.NUpD_CountSegmets.Name = "NUpD_CountSegmets"
        Me.NUpD_CountSegmets.Size = New System.Drawing.Size(55, 20)
        Me.NUpD_CountSegmets.TabIndex = 4
        Me.NUpD_CountSegmets.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'FormCreateConeLastPillars
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(209, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(448, 662)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FormCreateConeLastPillars"
        Me.Text = "Конус мостового сооружения"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DGV_PropertiesCone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NUpD_CountSegmets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label26 As Windows.Forms.Label
    Friend WithEvents CB_NumberPillar As Windows.Forms.ComboBox
    Friend WithEvents CBox_ListNamesBridge As Windows.Forms.ComboBox
    Friend WithEvents Label16 As Windows.Forms.Label
    Friend WithEvents CB_EgSurface As Windows.Forms.ComboBox
    Friend WithEvents Label15 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents CBox_ListNamesTemplateXML As Windows.Forms.ComboBox
    Friend WithEvents Label17 As Windows.Forms.Label
    Friend WithEvents ChB_ProjectSurfaceInAlignment As Windows.Forms.CheckBox
    Friend WithEvents Label14 As Windows.Forms.Label
    Friend WithEvents CB_NameAlignment As Windows.Forms.ComboBox
    Friend WithEvents Label13 As Windows.Forms.Label
    Friend WithEvents CB_ProjectSurface As Windows.Forms.ComboBox
    Friend WithEvents Label12 As Windows.Forms.Label
    Friend WithEvents CBox_ListNamesArrProject As Windows.Forms.ComboBox
    Friend WithEvents Button7 As Windows.Forms.Button
    Friend WithEvents Button8 As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents NUpD_CountSegmets As Windows.Forms.NumericUpDown
    Friend WithEvents DGV_PropertiesCone As Windows.Forms.DataGridView
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents CB_NameSites As Windows.Forms.ComboBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Номер As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As Windows.Forms.DataGridViewTextBoxColumn
End Class

