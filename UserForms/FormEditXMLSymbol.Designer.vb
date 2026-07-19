<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormEditXMLSymbol
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormEditXMLSymbol))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.КопироватьЭлементToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВставитьЭлементToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьЭлементToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ЗадатьЦветToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗадатьТипЛинииToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗадатьТаблицуPSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗадатьТолщинуЛинииToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗадатьСлойToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыбратьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЗадатьИмяШтриховкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolTip2 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ToolTip3 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TreeView1)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(533, 838)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Список элементов"
        '
        'TreeView1
        '
        Me.TreeView1.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TreeView1.Location = New System.Drawing.Point(6, 19)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(521, 811)
        Me.TreeView1.TabIndex = 0
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.КопироватьЭлементToolStripMenuItem, Me.ВставитьЭлементToolStripMenuItem, Me.УдалитьЭлементToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(189, 70)
        '
        'КопироватьЭлементToolStripMenuItem
        '
        Me.КопироватьЭлементToolStripMenuItem.Name = "КопироватьЭлементToolStripMenuItem"
        Me.КопироватьЭлементToolStripMenuItem.Size = New System.Drawing.Size(188, 22)
        Me.КопироватьЭлементToolStripMenuItem.Text = "Копировать элемент"
        '
        'ВставитьЭлементToolStripMenuItem
        '
        Me.ВставитьЭлементToolStripMenuItem.Name = "ВставитьЭлементToolStripMenuItem"
        Me.ВставитьЭлементToolStripMenuItem.Size = New System.Drawing.Size(188, 22)
        Me.ВставитьЭлементToolStripMenuItem.Text = "Вставить элемент"
        '
        'УдалитьЭлементToolStripMenuItem
        '
        Me.УдалитьЭлементToolStripMenuItem.Name = "УдалитьЭлементToolStripMenuItem"
        Me.УдалитьЭлементToolStripMenuItem.Size = New System.Drawing.Size(188, 22)
        Me.УдалитьЭлементToolStripMenuItem.Text = "Удалить элемент"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.DataGridView1)
        Me.GroupBox2.Location = New System.Drawing.Point(554, 9)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(523, 785)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Свойства"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 560)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Label1"
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button4.Location = New System.Drawing.Point(13, 21)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(53, 52)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "-"
        Me.ToolTip2.SetToolTip(Me.Button4, "Удалить элемент")
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button3.Location = New System.Drawing.Point(72, 21)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(53, 52)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "+"
        Me.ToolTip1.SetToolTip(Me.Button3, "Додавить элемент" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        Me.Button3.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGridView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DataGridView1.Location = New System.Drawing.Point(6, 79)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.Size = New System.Drawing.Size(510, 469)
        Me.DataGridView1.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.HeaderText = "Свойство"
        Me.Column1.MinimumWidth = 8
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 150
        '
        'Column2
        '
        Me.Column2.HeaderText = "Значение"
        Me.Column2.MinimumWidth = 8
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 300
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ЗадатьЦветToolStripMenuItem, Me.ЗадатьТипЛинииToolStripMenuItem, Me.ЗадатьТаблицуPSToolStripMenuItem, Me.ЗадатьТолщинуЛинииToolStripMenuItem, Me.ЗадатьСлойToolStripMenuItem, Me.ВыбратьToolStripMenuItem, Me.ЗадатьИмяШтриховкиToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(202, 158)
        '
        'ЗадатьЦветToolStripMenuItem
        '
        Me.ЗадатьЦветToolStripMenuItem.Name = "ЗадатьЦветToolStripMenuItem"
        Me.ЗадатьЦветToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьЦветToolStripMenuItem.Text = "Задать цвет"
        '
        'ЗадатьТипЛинииToolStripMenuItem
        '
        Me.ЗадатьТипЛинииToolStripMenuItem.Name = "ЗадатьТипЛинииToolStripMenuItem"
        Me.ЗадатьТипЛинииToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьТипЛинииToolStripMenuItem.Text = "Задать тип линии"
        '
        'ЗадатьТаблицуPSToolStripMenuItem
        '
        Me.ЗадатьТаблицуPSToolStripMenuItem.Name = "ЗадатьТаблицуPSToolStripMenuItem"
        Me.ЗадатьТаблицуPSToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьТаблицуPSToolStripMenuItem.Text = "Задать таблицу PS"
        '
        'ЗадатьТолщинуЛинииToolStripMenuItem
        '
        Me.ЗадатьТолщинуЛинииToolStripMenuItem.Name = "ЗадатьТолщинуЛинииToolStripMenuItem"
        Me.ЗадатьТолщинуЛинииToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьТолщинуЛинииToolStripMenuItem.Text = "Задать толщину линии"
        '
        'ЗадатьСлойToolStripMenuItem
        '
        Me.ЗадатьСлойToolStripMenuItem.Name = "ЗадатьСлойToolStripMenuItem"
        Me.ЗадатьСлойToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьСлойToolStripMenuItem.Text = "Задать слой"
        '
        'ВыбратьToolStripMenuItem
        '
        Me.ВыбратьToolStripMenuItem.Name = "ВыбратьToolStripMenuItem"
        Me.ВыбратьToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ВыбратьToolStripMenuItem.Text = "Выбрать объект"
        '
        'ЗадатьИмяШтриховкиToolStripMenuItem
        '
        Me.ЗадатьИмяШтриховкиToolStripMenuItem.Name = "ЗадатьИмяШтриховкиToolStripMenuItem"
        Me.ЗадатьИмяШтриховкиToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.ЗадатьИмяШтриховкиToolStripMenuItem.Text = "Задать имя штриховки"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(554, 814)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(168, 34)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(901, 814)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(168, 34)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Закрыть"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'FormEditXMLSymbol
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(209, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1080, 856)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(1096, 895)
        Me.MinimumSize = New System.Drawing.Size(1096, 895)
        Me.Name = "FormEditXMLSymbol"
        Me.Text = "Редактор библиотеки условных знаков"
        Me.GroupBox1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents TreeView1 As Windows.Forms.TreeView
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents ЗадатьЦветToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЗадатьТипЛинииToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЗадатьТаблицуPSToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЗадатьТолщинуЛинииToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЗадатьСлойToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВыбратьToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Column1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents ToolTip2 As Windows.Forms.ToolTip
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents ToolTip3 As Windows.Forms.ToolTip
    Friend WithEvents ЗадатьИмяШтриховкиToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContextMenuStrip2 As Windows.Forms.ContextMenuStrip
    Friend WithEvents КопироватьЭлементToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВставитьЭлементToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents УдалитьЭлементToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As Windows.Forms.Label
End Class
