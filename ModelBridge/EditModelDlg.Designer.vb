<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EditModelDlg
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
        Me.tbArrayValues = New System.Windows.Forms.TextBox()
        Me.label4 = New System.Windows.Forms.Label()
        Me.tbIntValue = New System.Windows.Forms.TextBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.tbDoubleValue = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.cbBooleanValue = New System.Windows.Forms.CheckBox()
        Me.tbStringValue = New System.Windows.Forms.TextBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbArrayValues
        '
        Me.tbArrayValues.Location = New System.Drawing.Point(13, 189)
        Me.tbArrayValues.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.tbArrayValues.Multiline = True
        Me.tbArrayValues.Name = "tbArrayValues"
        Me.tbArrayValues.Size = New System.Drawing.Size(556, 122)
        Me.tbArrayValues.TabIndex = 118
        '
        'label4
        '
        Me.label4.AutoSize = True
        Me.label4.Location = New System.Drawing.Point(9, 164)
        Me.label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(213, 20)
        Me.label4.TabIndex = 117
        Me.label4.Text = "Строки с разделителем "";"""
        '
        'tbIntValue
        '
        Me.tbIntValue.Location = New System.Drawing.Point(459, 94)
        Me.tbIntValue.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.tbIntValue.Name = "tbIntValue"
        Me.tbIntValue.Size = New System.Drawing.Size(110, 26)
        Me.tbIntValue.TabIndex = 116
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(9, 98)
        Me.label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(107, 20)
        Me.label3.TabIndex = 115
        Me.label3.Text = "Целое число"
        '
        'tbDoubleValue
        '
        Me.tbDoubleValue.Location = New System.Drawing.Point(459, 54)
        Me.tbDoubleValue.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.tbDoubleValue.Name = "tbDoubleValue"
        Me.tbDoubleValue.Size = New System.Drawing.Size(110, 26)
        Me.tbDoubleValue.TabIndex = 114
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(9, 58)
        Me.label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(219, 20)
        Me.label2.TabIndex = 113
        Me.label2.Text = "Число с плавающей точкой"
        '
        'cbBooleanValue
        '
        Me.cbBooleanValue.AutoSize = True
        Me.cbBooleanValue.Location = New System.Drawing.Point(13, 134)
        Me.cbBooleanValue.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cbBooleanValue.Name = "cbBooleanValue"
        Me.cbBooleanValue.Size = New System.Drawing.Size(295, 24)
        Me.cbBooleanValue.TabIndex = 112
        Me.cbBooleanValue.Text = "Флаг состояния, истина или ложь"
        Me.cbBooleanValue.UseVisualStyleBackColor = True
        '
        'tbStringValue
        '
        Me.tbStringValue.Location = New System.Drawing.Point(213, 14)
        Me.tbStringValue.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.tbStringValue.Name = "tbStringValue"
        Me.tbStringValue.Size = New System.Drawing.Size(356, 26)
        Me.tbStringValue.TabIndex = 111
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(9, 18)
        Me.label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(166, 20)
        Me.label1.TabIndex = 110
        Me.label1.Text = "Строковое значение"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(13, 321)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(150, 35)
        Me.btnCancel.TabIndex = 100
        Me.btnCancel.Text = "Отмена"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(419, 321)
        Me.btnOk.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(150, 35)
        Me.btnOk.TabIndex = 120
        Me.btnOk.Text = "ОК"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'EditModelDlg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(578, 369)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.tbArrayValues)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.tbIntValue)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.tbDoubleValue)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.cbBooleanValue)
        Me.Controls.Add(Me.tbStringValue)
        Me.Controls.Add(Me.label1)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximumSize = New System.Drawing.Size(600, 425)
        Me.MinimumSize = New System.Drawing.Size(600, 425)
        Me.Name = "EditModelDlg"
        Me.Text = "Редактор модели"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents tbArrayValues As Windows.Forms.TextBox
    Private WithEvents label4 As Windows.Forms.Label
    Private WithEvents tbIntValue As Windows.Forms.TextBox
    Private WithEvents label3 As Windows.Forms.Label
    Private WithEvents tbDoubleValue As Windows.Forms.TextBox
    Private WithEvents label2 As Windows.Forms.Label
    Private WithEvents cbBooleanValue As Windows.Forms.CheckBox
    Private WithEvents tbStringValue As Windows.Forms.TextBox
    Private WithEvents label1 As Windows.Forms.Label
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnOk As Windows.Forms.Button
End Class
