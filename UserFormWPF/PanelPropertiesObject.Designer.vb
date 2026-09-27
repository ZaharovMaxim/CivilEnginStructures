<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PanelPropertiesObject
    Inherits System.Windows.Forms.UserControl

    'Пользовательский элемент управления (UserControl) переопределяет метод Dispose для очистки списка компонентов.
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button_LocalProperties = New System.Windows.Forms.Button()
        Me.Button_GlobalProperties = New System.Windows.Forms.Button()
        Me.ComboBox_DataTable = New System.Windows.Forms.ComboBox()
        Me.DataGrid_PropertiesEnt = New System.Windows.Forms.DataGridView()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGrid_PropertiesEnt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button_LocalProperties)
        Me.GroupBox1.Controls.Add(Me.Button_GlobalProperties)
        Me.GroupBox1.Controls.Add(Me.ComboBox_DataTable)
        Me.GroupBox1.Controls.Add(Me.DataGrid_PropertiesEnt)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(289, 317)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Пользовательские свойства"
        '
        'Button_LocalProperties
        '
        Me.Button_LocalProperties.Location = New System.Drawing.Point(228, 19)
        Me.Button_LocalProperties.Name = "Button_LocalProperties"
        Me.Button_LocalProperties.Size = New System.Drawing.Size(25, 25)
        Me.Button_LocalProperties.TabIndex = 3
        Me.Button_LocalProperties.Text = "x"
        Me.Button_LocalProperties.UseVisualStyleBackColor = True
        '
        'Button_GlobalProperties
        '
        Me.Button_GlobalProperties.Location = New System.Drawing.Point(197, 19)
        Me.Button_GlobalProperties.Name = "Button_GlobalProperties"
        Me.Button_GlobalProperties.Size = New System.Drawing.Size(25, 25)
        Me.Button_GlobalProperties.TabIndex = 2
        Me.Button_GlobalProperties.Text = "+"
        Me.Button_GlobalProperties.UseVisualStyleBackColor = True
        '
        'ComboBox_DataTable
        '
        Me.ComboBox_DataTable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_DataTable.FormattingEnabled = True
        Me.ComboBox_DataTable.Location = New System.Drawing.Point(6, 19)
        Me.ComboBox_DataTable.Name = "ComboBox_DataTable"
        Me.ComboBox_DataTable.Size = New System.Drawing.Size(185, 21)
        Me.ComboBox_DataTable.TabIndex = 1
        '
        'DataGrid_PropertiesEnt
        '
        Me.DataGrid_PropertiesEnt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGrid_PropertiesEnt.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGrid_PropertiesEnt.Location = New System.Drawing.Point(6, 50)
        Me.DataGrid_PropertiesEnt.Name = "DataGrid_PropertiesEnt"
        Me.DataGrid_PropertiesEnt.RowHeadersVisible = False
        Me.DataGrid_PropertiesEnt.RowHeadersWidth = 62
        Me.DataGrid_PropertiesEnt.Size = New System.Drawing.Size(274, 254)
        Me.DataGrid_PropertiesEnt.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        Me.ToolStrip1.Size = New System.Drawing.Size(307, 25)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Column1
        '
        Me.Column1.HeaderText = "Имя поля"
        Me.Column1.MinimumWidth = 8
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column1.Width = 110
        '
        'Column2
        '
        Me.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Column2.HeaderText = "Значение"
        Me.Column2.MinimumWidth = 8
        Me.Column2.Name = "Column2"
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'PanelPropertiesObject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "PanelPropertiesObject"
        Me.Size = New System.Drawing.Size(307, 340)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGrid_PropertiesEnt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents DataGrid_PropertiesEnt As Windows.Forms.DataGridView
    Friend WithEvents ComboBox_DataTable As Windows.Forms.ComboBox
    Friend WithEvents Button_LocalProperties As Windows.Forms.Button
    Friend WithEvents Button_GlobalProperties As Windows.Forms.Button
    Friend WithEvents ToolStrip1 As Windows.Forms.ToolStrip
    Friend WithEvents ImageList1 As Windows.Forms.ImageList
    Friend WithEvents Column1 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As Windows.Forms.DataGridViewTextBoxColumn
End Class
