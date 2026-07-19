<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormSelectCheckBox
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormSelectCheckBox))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ContextMenuSelectObject = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ВыбратьВсеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОтменитьВыборToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1.SuspendLayout()
        Me.ContextMenuSelectObject.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckedListBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 9)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(420, 220)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Список доступных объектов"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.ContextMenuStrip = Me.ContextMenuSelectObject
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(11, 24)
        Me.CheckedListBox1.MultiColumn = True
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(400, 184)
        Me.CheckedListBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(18, 241)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(120, 30)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Отмена"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(287, 241)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 30)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "OK"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ContextMenuSelectObject
        '
        Me.ContextMenuSelectObject.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ВыбратьВсеToolStripMenuItem, Me.ОтменитьВыборToolStripMenuItem})
        Me.ContextMenuSelectObject.Name = "ContextMenuSelectObject"
        Me.ContextMenuSelectObject.Size = New System.Drawing.Size(181, 70)
        '
        'ВыбратьВсеToolStripMenuItem
        '
        Me.ВыбратьВсеToolStripMenuItem.Name = "ВыбратьВсеToolStripMenuItem"
        Me.ВыбратьВсеToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ВыбратьВсеToolStripMenuItem.Text = "Выбрать все"
        '
        'ОтменитьВыборToolStripMenuItem
        '
        Me.ОтменитьВыборToolStripMenuItem.Name = "ОтменитьВыборToolStripMenuItem"
        Me.ОтменитьВыборToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.ОтменитьВыборToolStripMenuItem.Text = "Отменить выбор"
        '
        'FormSelectCheckBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(209, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(432, 275)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(448, 314)
        Me.MinimumSize = New System.Drawing.Size(448, 314)
        Me.Name = "FormSelectCheckBox"
        Me.Text = "Выбор объектов"
        Me.GroupBox1.ResumeLayout(False)
        Me.ContextMenuSelectObject.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckedListBox1 As Windows.Forms.CheckedListBox
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents ContextMenuSelectObject As Windows.Forms.ContextMenuStrip
    Friend WithEvents ВыбратьВсеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ОтменитьВыборToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class
