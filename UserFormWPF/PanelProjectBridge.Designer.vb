<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class PanelProjectBridge
    Inherits System.Windows.Forms.UserControl

    'Пользовательский элемент управления (UserControl) переопределяет метод Dispose для очистки списка компонентов.
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PanelProjectBridge))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.ReportMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ТочкиОпиранияБалокToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОсиОпорToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ДеформационныеЗазорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ЭкспортЭлементовВDwgToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImageListMenu = New System.Windows.Forms.ImageList(Me.components)
        Me.Button6 = New System.Windows.Forms.Button()
        Me.PillarsMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.СоздатьОсьОпорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьВсеОпорыToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьКрайнююОпоруToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.MenuBeams = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.СоздатьБалкуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПереместитьБалкуToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВосстановитьЭлементыБалкиToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.UpdateStructureMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PropertyGrid1 = New System.Windows.Forms.PropertyGrid()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.MenuSelectedNodeTree = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ПоказатьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ВыбратьЭлементToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоказатьНаПоперечникеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоказатьВсеНаПоперечникеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ОбновитьToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImageListTree = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CBox_ListNamesArrProject = New System.Windows.Forms.ComboBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BridgeMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.СоздатьНовыйПустойМостToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.УдалитьСооружениеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ПоднятьОпуститьСооружениеToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1.SuspendLayout()
        Me.ReportMenu.SuspendLayout()
        Me.PillarsMenu.SuspendLayout()
        Me.MenuBeams.SuspendLayout()
        Me.UpdateStructureMenu.SuspendLayout()
        Me.MenuSelectedNodeTree.SuspendLayout()
        Me.BridgeMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.PropertyGrid1)
        Me.GroupBox1.Controls.Add(Me.TreeView1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.CBox_ListNamesArrProject)
        Me.GroupBox1.Location = New System.Drawing.Point(14, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(580, 926)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Button3
        '
        Me.Button3.ContextMenuStrip = Me.ReportMenu
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.ForeColor = System.Drawing.SystemColors.Control
        Me.Button3.ImageKey = "060_Отчеты.png"
        Me.Button3.ImageList = Me.ImageListMenu
        Me.Button3.Location = New System.Drawing.Point(240, 23)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(45, 46)
        Me.Button3.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.Button3, "Вывести отчет")
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ReportMenu
        '
        Me.ReportMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ReportMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ТочкиОпиранияБалокToolStripMenuItem, Me.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem, Me.ОсиОпорToolStripMenuItem, Me.ДеформационныеЗазорыToolStripMenuItem, Me.ЭкспортЭлементовВDwgToolStripMenuItem})
        Me.ReportMenu.Name = "ReportMenu"
        Me.ReportMenu.Size = New System.Drawing.Size(410, 164)
        '
        'ТочкиОпиранияБалокToolStripMenuItem
        '
        Me.ТочкиОпиранияБалокToolStripMenuItem.Name = "ТочкиОпиранияБалокToolStripMenuItem"
        Me.ТочкиОпиранияБалокToolStripMenuItem.Size = New System.Drawing.Size(409, 32)
        Me.ТочкиОпиранияБалокToolStripMenuItem.Text = "Точки опирания балок"
        '
        'ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem
        '
        Me.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem.Name = "ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem"
        Me.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem.Size = New System.Drawing.Size(409, 32)
        Me.ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem.Text = "Верх плиты балок и толщины покрытия"
        '
        'ОсиОпорToolStripMenuItem
        '
        Me.ОсиОпорToolStripMenuItem.Name = "ОсиОпорToolStripMenuItem"
        Me.ОсиОпорToolStripMenuItem.Size = New System.Drawing.Size(409, 32)
        Me.ОсиОпорToolStripMenuItem.Text = "Оси опор"
        '
        'ДеформационныеЗазорыToolStripMenuItem
        '
        Me.ДеформационныеЗазорыToolStripMenuItem.Name = "ДеформационныеЗазорыToolStripMenuItem"
        Me.ДеформационныеЗазорыToolStripMenuItem.Size = New System.Drawing.Size(409, 32)
        Me.ДеформационныеЗазорыToolStripMenuItem.Text = "Деформационные зазоры"
        '
        'ЭкспортЭлементовВDwgToolStripMenuItem
        '
        Me.ЭкспортЭлементовВDwgToolStripMenuItem.Name = "ЭкспортЭлементовВDwgToolStripMenuItem"
        Me.ЭкспортЭлементовВDwgToolStripMenuItem.Size = New System.Drawing.Size(409, 32)
        Me.ЭкспортЭлементовВDwgToolStripMenuItem.Text = "Экспорт элементов в dwg"
        '
        'ImageListMenu
        '
        Me.ImageListMenu.ImageStream = CType(resources.GetObject("ImageListMenu.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListMenu.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListMenu.Images.SetKeyName(0, "001_Структура проекта.png")
        Me.ImageListMenu.Images.SetKeyName(1, "002_Обновить структуру проекта.png")
        Me.ImageListMenu.Images.SetKeyName(2, "003_Сооружение.png")
        Me.ImageListMenu.Images.SetKeyName(3, "004_Перестроить сооружение.png")
        Me.ImageListMenu.Images.SetKeyName(4, "005_Пролетные строения.png")
        Me.ImageListMenu.Images.SetKeyName(5, "006_Оси сооружения.png")
        Me.ImageListMenu.Images.SetKeyName(6, "007_Прочие элементы.png")
        Me.ImageListMenu.Images.SetKeyName(7, "010_Мостовое полотно.png")
        Me.ImageListMenu.Images.SetKeyName(8, "020_Опоры.png")
        Me.ImageListMenu.Images.SetKeyName(9, "021_Промежуточные опоры.png")
        Me.ImageListMenu.Images.SetKeyName(10, "022_Ось промежуточной опоры.png")
        Me.ImageListMenu.Images.SetKeyName(11, "023_Сваи.png")
        Me.ImageListMenu.Images.SetKeyName(12, "040_Балки.png")
        Me.ImageListMenu.Images.SetKeyName(13, "040_Участки омоноличивания балок.png")
        Me.ImageListMenu.Images.SetKeyName(14, "041_Представление по пролетам.png")
        Me.ImageListMenu.Images.SetKeyName(15, "042_Представление по рядам.png")
        Me.ImageListMenu.Images.SetKeyName(16, "060_Отчеты.png")
        '
        'Button6
        '
        Me.Button6.ContextMenuStrip = Me.PillarsMenu
        Me.Button6.FlatAppearance.BorderSize = 0
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.ForeColor = System.Drawing.SystemColors.Control
        Me.Button6.ImageKey = "020_Опоры.png"
        Me.Button6.ImageList = Me.ImageListMenu
        Me.Button6.Location = New System.Drawing.Point(195, 23)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(45, 46)
        Me.Button6.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.Button6, "Редактировать опоры")
        Me.Button6.UseVisualStyleBackColor = True
        '
        'PillarsMenu
        '
        Me.PillarsMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.PillarsMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СоздатьОсьОпорыToolStripMenuItem, Me.УдалитьВсеОпорыToolStripMenuItem, Me.УдалитьКрайнююОпоруToolStripMenuItem})
        Me.PillarsMenu.Name = "PillarsMenu"
        Me.PillarsMenu.Size = New System.Drawing.Size(341, 100)
        '
        'СоздатьОсьОпорыToolStripMenuItem
        '
        Me.СоздатьОсьОпорыToolStripMenuItem.Name = "СоздатьОсьОпорыToolStripMenuItem"
        Me.СоздатьОсьОпорыToolStripMenuItem.Size = New System.Drawing.Size(340, 32)
        Me.СоздатьОсьОпорыToolStripMenuItem.Text = "Перерассчет элементов опоры"
        '
        'УдалитьВсеОпорыToolStripMenuItem
        '
        Me.УдалитьВсеОпорыToolStripMenuItem.Name = "УдалитьВсеОпорыToolStripMenuItem"
        Me.УдалитьВсеОпорыToolStripMenuItem.Size = New System.Drawing.Size(340, 32)
        Me.УдалитьВсеОпорыToolStripMenuItem.Text = "Удалить элементы всех опор"
        '
        'УдалитьКрайнююОпоруToolStripMenuItem
        '
        Me.УдалитьКрайнююОпоруToolStripMenuItem.Name = "УдалитьКрайнююОпоруToolStripMenuItem"
        Me.УдалитьКрайнююОпоруToolStripMenuItem.Size = New System.Drawing.Size(340, 32)
        Me.УдалитьКрайнююОпоруToolStripMenuItem.Text = "Удалить элементы опоры"
        '
        'Button5
        '
        Me.Button5.FlatAppearance.BorderSize = 0
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.ForeColor = System.Drawing.SystemColors.Control
        Me.Button5.ImageKey = "003_Сооружение.png"
        Me.Button5.ImageList = Me.ImageListMenu
        Me.Button5.Location = New System.Drawing.Point(105, 23)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(45, 46)
        Me.Button5.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.Button5, "Создать\редактировать сооружение")
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.ContextMenuStrip = Me.MenuBeams
        Me.Button4.FlatAppearance.BorderSize = 0
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.ForeColor = System.Drawing.SystemColors.Control
        Me.Button4.ImageKey = "040_Балки.png"
        Me.Button4.ImageList = Me.ImageListMenu
        Me.Button4.Location = New System.Drawing.Point(150, 23)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(45, 46)
        Me.Button4.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.Button4, "Работа с балками сооружения")
        Me.Button4.UseVisualStyleBackColor = True
        '
        'MenuBeams
        '
        Me.MenuBeams.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuBeams.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СоздатьБалкуToolStripMenuItem, Me.СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem, Me.ПереместитьБалкуToolStripMenuItem, Me.УдалитьToolStripMenuItem, Me.ВосстановитьЭлементыБалкиToolStripMenuItem})
        Me.MenuBeams.Name = "ContextMenuStrip3"
        Me.MenuBeams.Size = New System.Drawing.Size(480, 164)
        '
        'СоздатьБалкуToolStripMenuItem
        '
        Me.СоздатьБалкуToolStripMenuItem.Name = "СоздатьБалкуToolStripMenuItem"
        Me.СоздатьБалкуToolStripMenuItem.Size = New System.Drawing.Size(479, 32)
        Me.СоздатьБалкуToolStripMenuItem.Text = "Создать балку по смещению от оси автодороги"
        '
        'СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem
        '
        Me.СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem.Name = "СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem"
        Me.СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem.Size = New System.Drawing.Size(479, 32)
        Me.СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem.Text = "Создать балку по высотным отметкам"
        '
        'ПереместитьБалкуToolStripMenuItem
        '
        Me.ПереместитьБалкуToolStripMenuItem.Name = "ПереместитьБалкуToolStripMenuItem"
        Me.ПереместитьБалкуToolStripMenuItem.Size = New System.Drawing.Size(479, 32)
        Me.ПереместитьБалкуToolStripMenuItem.Text = "Переместить балку вдоль оси автодороги"
        '
        'УдалитьToolStripMenuItem
        '
        Me.УдалитьToolStripMenuItem.Name = "УдалитьToolStripMenuItem"
        Me.УдалитьToolStripMenuItem.Size = New System.Drawing.Size(479, 32)
        Me.УдалитьToolStripMenuItem.Text = "Удалить балку"
        '
        'ВосстановитьЭлементыБалкиToolStripMenuItem
        '
        Me.ВосстановитьЭлементыБалкиToolStripMenuItem.Name = "ВосстановитьЭлементыБалкиToolStripMenuItem"
        Me.ВосстановитьЭлементыБалкиToolStripMenuItem.Size = New System.Drawing.Size(479, 32)
        Me.ВосстановитьЭлементыБалкиToolStripMenuItem.Text = "Восстановить элементы балки"
        '
        'Button2
        '
        Me.Button2.FlatAppearance.BorderSize = 0
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.ForeColor = System.Drawing.SystemColors.Control
        Me.Button2.ImageKey = "004_Перестроить сооружение.png"
        Me.Button2.ImageList = Me.ImageListMenu
        Me.Button2.Location = New System.Drawing.Point(52, 23)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(45, 46)
        Me.Button2.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.Button2, "Перестроить сооружение")
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.ContextMenuStrip = Me.UpdateStructureMenu
        Me.Button1.FlatAppearance.BorderSize = 0
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.SystemColors.Control
        Me.Button1.ImageKey = "002_Обновить структуру проекта.png"
        Me.Button1.ImageList = Me.ImageListMenu
        Me.Button1.Location = New System.Drawing.Point(8, 23)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(45, 46)
        Me.Button1.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.Button1, "Обновить Структуру проекта")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'UpdateStructureMenu
        '
        Me.UpdateStructureMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.UpdateStructureMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripMenuItem2})
        Me.UpdateStructureMenu.Name = "MenuSelectedNodeTree"
        Me.UpdateStructureMenu.Size = New System.Drawing.Size(355, 68)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(354, 32)
        Me.ToolStripMenuItem1.Text = "Обновить список проектов"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(354, 32)
        Me.ToolStripMenuItem2.Text = "Обновить элементы сооружения"
        '
        'PropertyGrid1
        '
        Me.PropertyGrid1.Location = New System.Drawing.Point(8, 554)
        Me.PropertyGrid1.Name = "PropertyGrid1"
        Me.PropertyGrid1.PropertySort = System.Windows.Forms.PropertySort.NoSort
        Me.PropertyGrid1.Size = New System.Drawing.Size(562, 362)
        Me.PropertyGrid1.TabIndex = 3
        '
        'TreeView1
        '
        Me.TreeView1.ContextMenuStrip = Me.MenuSelectedNodeTree
        Me.TreeView1.HideSelection = False
        Me.TreeView1.ImageIndex = 0
        Me.TreeView1.ImageList = Me.ImageListTree
        Me.TreeView1.Location = New System.Drawing.Point(6, 126)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 0
        Me.TreeView1.Size = New System.Drawing.Size(560, 396)
        Me.TreeView1.TabIndex = 2
        '
        'MenuSelectedNodeTree
        '
        Me.MenuSelectedNodeTree.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuSelectedNodeTree.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ПоказатьToolStripMenuItem, Me.ВыбратьЭлементToolStripMenuItem, Me.ПоказатьНаПоперечникеToolStripMenuItem, Me.ПоказатьВсеНаПоперечникеToolStripMenuItem, Me.ОбновитьToolStripMenuItem})
        Me.MenuSelectedNodeTree.Name = "MenuSelectedNodeTree"
        Me.MenuSelectedNodeTree.Size = New System.Drawing.Size(329, 164)
        '
        'ПоказатьToolStripMenuItem
        '
        Me.ПоказатьToolStripMenuItem.Name = "ПоказатьToolStripMenuItem"
        Me.ПоказатьToolStripMenuItem.Size = New System.Drawing.Size(328, 32)
        Me.ПоказатьToolStripMenuItem.Text = "Показать"
        '
        'ВыбратьЭлементToolStripMenuItem
        '
        Me.ВыбратьЭлементToolStripMenuItem.Name = "ВыбратьЭлементToolStripMenuItem"
        Me.ВыбратьЭлементToolStripMenuItem.Size = New System.Drawing.Size(328, 32)
        Me.ВыбратьЭлементToolStripMenuItem.Text = "Выбрать элемент"
        '
        'ПоказатьНаПоперечникеToolStripMenuItem
        '
        Me.ПоказатьНаПоперечникеToolStripMenuItem.Name = "ПоказатьНаПоперечникеToolStripMenuItem"
        Me.ПоказатьНаПоперечникеToolStripMenuItem.Size = New System.Drawing.Size(328, 32)
        Me.ПоказатьНаПоперечникеToolStripMenuItem.Text = "Показать на поперечнике"
        '
        'ПоказатьВсеНаПоперечникеToolStripMenuItem
        '
        Me.ПоказатьВсеНаПоперечникеToolStripMenuItem.Name = "ПоказатьВсеНаПоперечникеToolStripMenuItem"
        Me.ПоказатьВсеНаПоперечникеToolStripMenuItem.Size = New System.Drawing.Size(328, 32)
        Me.ПоказатьВсеНаПоперечникеToolStripMenuItem.Text = "Показать все на поперечнике"
        '
        'ОбновитьToolStripMenuItem
        '
        Me.ОбновитьToolStripMenuItem.Name = "ОбновитьToolStripMenuItem"
        Me.ОбновитьToolStripMenuItem.Size = New System.Drawing.Size(328, 32)
        Me.ОбновитьToolStripMenuItem.Text = "Обновить модель"
        '
        'ImageListTree
        '
        Me.ImageListTree.ImageStream = CType(resources.GetObject("ImageListTree.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListTree.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListTree.Images.SetKeyName(0, "001_Структура проекта.png")
        Me.ImageListTree.Images.SetKeyName(1, "002_Обновить структуру проекта.png")
        Me.ImageListTree.Images.SetKeyName(2, "003_Сооружение.png")
        Me.ImageListTree.Images.SetKeyName(3, "004_Перестроить сооружение.png")
        Me.ImageListTree.Images.SetKeyName(4, "005_Пролетные строения.png")
        Me.ImageListTree.Images.SetKeyName(5, "006_Оси сооружения.png")
        Me.ImageListTree.Images.SetKeyName(6, "007_Прочие элементы.png")
        Me.ImageListTree.Images.SetKeyName(7, "010_Мостовое полотно.png")
        Me.ImageListTree.Images.SetKeyName(8, "020_Опоры.png")
        Me.ImageListTree.Images.SetKeyName(9, "021_Промежуточные опоры.png")
        Me.ImageListTree.Images.SetKeyName(10, "022_Ось промежуточной опоры.png")
        Me.ImageListTree.Images.SetKeyName(11, "023_Сваи.png")
        Me.ImageListTree.Images.SetKeyName(12, "040_Балки.png")
        Me.ImageListTree.Images.SetKeyName(13, "040_Участки омоноличивания балок.png")
        Me.ImageListTree.Images.SetKeyName(14, "041_Представление по пролетам.png")
        Me.ImageListTree.Images.SetKeyName(15, "042_Представление по рядам.png")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(290, 82)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Сооружение"
        '
        'CBox_ListNamesArrProject
        '
        Me.CBox_ListNamesArrProject.BackColor = System.Drawing.Color.White
        Me.CBox_ListNamesArrProject.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CBox_ListNamesArrProject.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CBox_ListNamesArrProject.FormattingEnabled = True
        Me.CBox_ListNamesArrProject.Location = New System.Drawing.Point(6, 77)
        Me.CBox_ListNamesArrProject.Name = "CBox_ListNamesArrProject"
        Me.CBox_ListNamesArrProject.Size = New System.Drawing.Size(276, 28)
        Me.CBox_ListNamesArrProject.TabIndex = 0
        '
        'BridgeMenu
        '
        Me.BridgeMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.BridgeMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.СоздатьНовыйПустойМостToolStripMenuItem, Me.УдалитьСооружениеToolStripMenuItem, Me.ПоднятьОпуститьСооружениеToolStripMenuItem})
        Me.BridgeMenu.Name = "BridgeMenu"
        Me.BridgeMenu.Size = New System.Drawing.Size(381, 100)
        '
        'СоздатьНовыйПустойМостToolStripMenuItem
        '
        Me.СоздатьНовыйПустойМостToolStripMenuItem.Name = "СоздатьНовыйПустойМостToolStripMenuItem"
        Me.СоздатьНовыйПустойМостToolStripMenuItem.Size = New System.Drawing.Size(380, 32)
        Me.СоздатьНовыйПустойМостToolStripMenuItem.Text = "Создать новое (пустое) сооружение"
        '
        'УдалитьСооружениеToolStripMenuItem
        '
        Me.УдалитьСооружениеToolStripMenuItem.Name = "УдалитьСооружениеToolStripMenuItem"
        Me.УдалитьСооружениеToolStripMenuItem.Size = New System.Drawing.Size(380, 32)
        Me.УдалитьСооружениеToolStripMenuItem.Text = "Удалить сооружение (выборочно)"
        '
        'ПоднятьОпуститьСооружениеToolStripMenuItem
        '
        Me.ПоднятьОпуститьСооружениеToolStripMenuItem.Name = "ПоднятьОпуститьСооружениеToolStripMenuItem"
        Me.ПоднятьОпуститьСооружениеToolStripMenuItem.Size = New System.Drawing.Size(380, 32)
        Me.ПоднятьОпуститьСооружениеToolStripMenuItem.Text = "Сдвинуть сооружение вдоль оси"
        '
        'PanelProjectBridge
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "PanelProjectBridge"
        Me.Size = New System.Drawing.Size(609, 971)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ReportMenu.ResumeLayout(False)
        Me.PillarsMenu.ResumeLayout(False)
        Me.MenuBeams.ResumeLayout(False)
        Me.UpdateStructureMenu.ResumeLayout(False)
        Me.MenuSelectedNodeTree.ResumeLayout(False)
        Me.BridgeMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CBox_ListNamesArrProject As Windows.Forms.ComboBox
    Friend WithEvents TreeView1 As Windows.Forms.TreeView
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents PropertyGrid1 As Windows.Forms.PropertyGrid
    Friend WithEvents ImageListMenu As Windows.Forms.ImageList
    Friend WithEvents ImageListTree As Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents MenuBeams As Windows.Forms.ContextMenuStrip
    Friend WithEvents УдалитьToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button5 As Windows.Forms.Button
    Friend WithEvents СоздатьБалкуToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПереместитьБалкуToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button6 As Windows.Forms.Button
    Friend WithEvents PillarsMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents СоздатьОсьОпорыToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents BridgeMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents СоздатьНовыйПустойМостToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents УдалитьСооружениеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПоднятьОпуститьСооружениеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents ReportMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents ТочкиОпиранияБалокToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВерхПлитыБалокИТолщиныПокрытияToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ОсиОпорToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ДеформационныеЗазорыToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ЭкспортЭлементовВDwgToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents УдалитьВсеОпорыToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuSelectedNodeTree As Windows.Forms.ContextMenuStrip
    Friend WithEvents ПоказатьToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВыбратьЭлементToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПоказатьНаПоперечникеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents СоздатьБалкуПоВысотнымОтметкамToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ВосстановитьЭлементыБалкиToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents УдалитьКрайнююОпоруToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ОбновитьToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpdateStructureMenu As Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ПоказатьВсеНаПоперечникеToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class
