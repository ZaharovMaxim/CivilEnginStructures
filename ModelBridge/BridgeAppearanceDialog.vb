Imports System.Drawing
Imports System.Windows.Forms
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View

Public NotInheritable Class BridgeAppearanceApplyEventArgs
    Inherits EventArgs

    Public Sub New(selectionIds As IEnumerable(Of String), patch As BridgeAppearancePatch)
        Me.SelectionIds = selectionIds.ToArray()
        Me.Patch = patch
    End Sub

    Public ReadOnly Property SelectionIds As IReadOnlyList(Of String)
    Public ReadOnly Property Patch As BridgeAppearancePatch
    Public Property ErrorMessage As String
    Public Property AppliedCount As Integer
    Public Property SkippedBodyLabels As IReadOnlyList(Of String) = New String() {}
End Class

Public NotInheritable Class BridgeAppearanceDialog
    Inherits Form

    Private Const MixedText As String = "Разные значения"
    Private ReadOnly _items As Dictionary(Of String, BridgeAppearanceItem)
    Private ReadOnly _tree As TreeView
    Private ReadOnly _cadCheck As CheckBox
    Private ReadOnly _cadValue As ComboBox
    Private ReadOnly _cadColorButton As Button
    Private ReadOnly _layerCheck As CheckBox
    Private ReadOnly _layerValue As ComboBox
    Private ReadOnly _linetypeCheck As CheckBox
    Private ReadOnly _linetypeValue As ComboBox
    Private ReadOnly _scaleCheck As CheckBox
    Private ReadOnly _scaleValue As NumericUpDown
    Private ReadOnly _scaleCurrent As Label
    Private ReadOnly _weightCheck As CheckBox
    Private ReadOnly _weightValue As ComboBox
    Private ReadOnly _widthCheck As CheckBox
    Private ReadOnly _widthValue As NumericUpDown
    Private ReadOnly _widthCurrent As Label
    Private ReadOnly _bodyCheck As CheckBox
    Private ReadOnly _bodyColorButton As Button
    Private ReadOnly _bodyCurrent As Label
    Private ReadOnly _fillCheck As CheckBox
    Private ReadOnly _fillColorButton As Button
    Private ReadOnly _fillCurrent As Label
    Private ReadOnly _hatchPatternCheck As CheckBox
    Private ReadOnly _hatchPatternValue As ComboBox
    Private ReadOnly _hatchScaleCheck As CheckBox
    Private ReadOnly _hatchScaleValue As NumericUpDown
    Private ReadOnly _hatchScaleCurrent As Label
    Private ReadOnly _hatchAngleCheck As CheckBox
    Private ReadOnly _hatchAngleValue As NumericUpDown
    Private ReadOnly _hatchAngleCurrent As Label
    Private ReadOnly _status As Label
    Private _cadCompressedValue As Integer
    Private _bodyArgbValue As Integer
    Private _fillCompressedValue As Integer
    Private _cadChoiceValid As Boolean
    Private _bodyChoiceValid As Boolean
    Private _fillChoiceValid As Boolean
    Private _updatingTree As Boolean
    Private _loadingValues As Boolean

    Public Event ApplyRequested As EventHandler(Of BridgeAppearanceApplyEventArgs)

    Public Sub New(items As IEnumerable(Of BridgeAppearanceItem),
                   Optional layerNames As IEnumerable(Of String) = Nothing,
                   Optional linetypeNames As IEnumerable(Of String) = Nothing)
        If items Is Nothing Then Throw New ArgumentNullException(NameOf(items))
        _items = items.Where(Function(item) item IsNot Nothing).
            ToDictionary(Function(item) item.SelectionId, StringComparer.Ordinal)

        Font = New Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
        BackColor = Color.FromArgb(243, 246, 250)
        ForeColor = Color.FromArgb(32, 50, 77)
        Text = "Оформление элементов моста"
        StartPosition = FormStartPosition.CenterParent
        AutoScaleMode = AutoScaleMode.Dpi
        MinimumSize = New Size(900, 590)
        ClientSize = New Size(1040, 680)

        Dim accent As New Panel With {.Dock = DockStyle.Top, .Height = 5, .BackColor = Color.FromArgb(37, 99, 235)}
        Controls.Add(accent)

        Dim root As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(18),
            .ColumnCount = 2,
            .RowCount = 3,
            .BackColor = BackColor
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 45.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 55.0F))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Controls.Add(root)
        root.BringToFront()

        Dim title As New Label With {
            .AutoSize = True,
            .Font = New Font("Segoe UI Semibold", 17.0F, FontStyle.Bold),
            .Text = "Оформление моста",
            .Margin = New Padding(0, 0, 0, 14)
        }
        root.Controls.Add(title, 0, 0)
        root.SetColumnSpan(title, 2)

        Dim treeCard As Panel = CreateCard()
        treeCard.Margin = New Padding(0, 0, 10, 0)
        _tree = New TreeView With {
            .Dock = DockStyle.Fill,
            .BorderStyle = BorderStyle.None,
            .CheckBoxes = True,
            .HideSelection = False,
            .BackColor = Color.White,
            .ForeColor = ForeColor
        }
        treeCard.Controls.Add(_tree)
        root.Controls.Add(treeCard, 0, 1)

        Dim styleCard As Panel = CreateCard()
        styleCard.Margin = New Padding(10, 0, 0, 0)
        Dim fields As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .AutoScroll = True,
            .ColumnCount = 3,
            .RowCount = 14,
            .Padding = New Padding(2)
        }
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        fields.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        styleCard.Controls.Add(fields)
        root.Controls.Add(styleCard, 1, 1)

        _cadCheck = New CheckBox()
        _cadValue = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Dock = DockStyle.Fill}
        _cadValue.Items.AddRange(New Object() {MixedText & " / Не менять", "По слою", "По блоку", "Выбранный цвет"})
        _cadColorButton = New Button With {.Text = "Цвет…", .AutoSize = True}
        AddField(fields, 0, "Цвет на плане", _cadCheck, _cadValue, _cadColorButton)

        _layerCheck = New CheckBox()
        _layerValue = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Dock = DockStyle.Fill}
        If layerNames IsNot Nothing Then _layerValue.Items.AddRange(layerNames.OrderBy(Function(value) value).Cast(Of Object).ToArray())
        AddField(fields, 1, "Слой", _layerCheck, _layerValue, Nothing)

        _linetypeCheck = New CheckBox()
        _linetypeValue = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown, .Dock = DockStyle.Fill}
        If linetypeNames IsNot Nothing Then _linetypeValue.Items.AddRange(linetypeNames.OrderBy(Function(value) value).Cast(Of Object).ToArray())
        AddField(fields, 2, "Тип линии", _linetypeCheck, _linetypeValue, Nothing)

        _scaleCheck = New CheckBox()
        _scaleValue = CreateNumber(0.000001D, 1000000D, 1D, 6)
        _scaleCurrent = CreateCurrentLabel()
        AddField(fields, 3, "Масштаб типа линии", _scaleCheck, _scaleValue, _scaleCurrent)

        _weightCheck = New CheckBox()
        _weightValue = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Dock = DockStyle.Fill}
        _weightValue.Items.Add(MixedText & " / Не менять")
        For Each value As Integer In {-3, -2, -1, 0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211}
            _weightValue.Items.Add(New LineweightChoice(value))
        Next
        AddField(fields, 4, "Толщина линии", _weightCheck, _weightValue, Nothing)

        _widthCheck = New CheckBox()
        _widthValue = CreateNumber(0D, 1000000D, 0D, 6)
        _widthCurrent = CreateCurrentLabel()
        AddField(fields, 5, "Ширина полилинии", _widthCheck, _widthValue, _widthCurrent)

        _bodyCheck = New CheckBox()
        _bodyColorButton = New Button With {.Text = "Выбрать цвет…", .AutoSize = True}
        _bodyCurrent = CreateCurrentLabel()
        AddField(fields, 6, "Цвет 3D-тела", _bodyCheck, _bodyColorButton, _bodyCurrent)

        _fillCheck = New CheckBox()
        _fillColorButton = New Button With {.Text = "Выбрать цвет…", .AutoSize = True}
        _fillCurrent = CreateCurrentLabel()
        AddField(fields, 7, "Цвет заливки", _fillCheck, _fillColorButton, _fillCurrent)

        _hatchPatternCheck = New CheckBox()
        _hatchPatternValue = New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .Dock = DockStyle.Fill}
        _hatchPatternValue.Items.AddRange(New Object() {"SOLID", "ANSI31"})
        AddField(fields, 8, "Штриховка", _hatchPatternCheck, _hatchPatternValue, Nothing)

        _hatchScaleCheck = New CheckBox()
        _hatchScaleValue = CreateNumber(0.000001D, 1000000D, 1D, 6)
        _hatchScaleCurrent = CreateCurrentLabel()
        AddField(fields, 9, "Масштаб штриховки", _hatchScaleCheck, _hatchScaleValue, _hatchScaleCurrent)

        _hatchAngleCheck = New CheckBox()
        _hatchAngleValue = CreateNumber(-360D, 360D, 0D, 3)
        _hatchAngleCurrent = CreateCurrentLabel()
        AddField(fields, 10, "Угол штриховки", _hatchAngleCheck, _hatchAngleValue, _hatchAngleCurrent)

        Dim hint As New Label With {
            .AutoSize = True,
            .MaximumSize = New Size(470, 0),
            .ForeColor = Color.FromArgb(93, 109, 133),
            .Text = "Структурные линии поверхности оформляются средствами поверхности.",
            .Margin = New Padding(4, 16, 4, 0)
        }
        fields.Controls.Add(hint, 0, 11)
        fields.SetColumnSpan(hint, 3)

        Dim footer As New TableLayoutPanel With {.AutoSize = True, .Dock = DockStyle.Fill, .ColumnCount = 2, .Margin = New Padding(0, 14, 0, 0)}
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        footer.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        _status = New Label With {.AutoSize = True, .ForeColor = Color.FromArgb(93, 109, 133), .Anchor = AnchorStyles.Left}
        footer.Controls.Add(_status, 0, 0)
        Dim buttons As New FlowLayoutPanel With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False}
        Dim applyButton As New Button With {.Text = "Применить", .AutoSize = True}
        Dim okButton As New Button With {.Text = "ОК", .AutoSize = True}
        Dim cancelButton As New Button With {.Text = "Отмена", .AutoSize = True, .DialogResult = DialogResult.Cancel}
        buttons.Controls.Add(applyButton)
        buttons.Controls.Add(okButton)
        buttons.Controls.Add(cancelButton)
        footer.Controls.Add(buttons, 1, 0)
        root.Controls.Add(footer, 0, 2)
        root.SetColumnSpan(footer, 2)
        CancelButton = cancelButton

        PopulateTree()
        AddHandler _tree.AfterCheck, AddressOf TreeAfterCheck
        AddHandler _cadCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _layerCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _linetypeCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _scaleCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _weightCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _widthCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _bodyCheck.CheckedChanged, AddressOf BodyCheckChanged
        AddHandler _fillCheck.CheckedChanged, AddressOf FillCheckChanged
        AddHandler _hatchPatternCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _hatchScaleCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _hatchAngleCheck.CheckedChanged, AddressOf UpdateEditorsEnabled
        AddHandler _cadValue.SelectedIndexChanged, AddressOf CadChoiceChanged
        AddHandler _cadColorButton.Click, AddressOf ChooseCadColor
        AddHandler _bodyColorButton.Click, AddressOf ChooseBodyColor
        AddHandler _fillColorButton.Click, AddressOf ChooseFillColor
        AddHandler applyButton.Click, Sub() ApplyChanges(False)
        AddHandler okButton.Click, Sub() ApplyChanges(True)
        UpdateCommonValues()
    End Sub

    Private Shared Function CreateCard() As Panel
        Return New Panel With {
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(14)
        }
    End Function

    Private Shared Function CreateNumber(minimum As Decimal,
                                         maximum As Decimal,
                                         value As Decimal,
                                         decimalPlaces As Integer) As NumericUpDown
        Return New NumericUpDown With {
            .Minimum = minimum,
            .Maximum = maximum,
            .Value = value,
            .DecimalPlaces = decimalPlaces,
            .Increment = 0.1D,
            .Dock = DockStyle.Fill
        }
    End Function

    Private Shared Function CreateCurrentLabel() As Label
        Return New Label With {.AutoSize = True, .ForeColor = Color.FromArgb(93, 109, 133), .Anchor = AnchorStyles.Left}
    End Function

    Private Shared Sub AddField(layout As TableLayoutPanel,
                                row As Integer,
                                caption As String,
                                check As CheckBox,
                                editor As Control,
                                auxiliary As Control)
        check.Text = caption
        check.AutoSize = True
        check.Margin = New Padding(4, 8, 12, 8)
        editor.Margin = New Padding(4, 5, 4, 5)
        layout.Controls.Add(check, 0, row)
        layout.Controls.Add(editor, 1, row)
        If auxiliary IsNot Nothing Then
            auxiliary.Margin = New Padding(8, 5, 4, 5)
            layout.Controls.Add(auxiliary, 2, row)
        End If
    End Sub

    Private Sub PopulateTree()
        _tree.BeginUpdate()
        Try
            For Each bridgeGroup In _items.Values.GroupBy(
                Function(item) BridgeDrawingGroupManager.GetGroupName(item.BridgeId),
                StringComparer.Ordinal).OrderBy(Function(group) group.First().BridgeLabel)
                Dim bridgeNode As New TreeNode(bridgeGroup.First().BridgeLabel)
                For Each classGroup In bridgeGroup.GroupBy(Function(item) item.ClassLabel).OrderBy(Function(group) group.Key)
                    Dim classNode As New TreeNode(classGroup.Key)
                    For Each typeGroup In classGroup.GroupBy(Function(item) item.TypeLabel).OrderBy(Function(group) group.Key)
                        Dim typeNode As New TreeNode(typeGroup.Key)
                        For Each item As BridgeAppearanceItem In typeGroup.OrderBy(Function(value) value.ItemLabel)
                            typeNode.Nodes.Add(New TreeNode(item.ItemLabel) With {.Tag = item.SelectionId})
                        Next
                        classNode.Nodes.Add(typeNode)
                    Next
                    bridgeNode.Nodes.Add(classNode)
                Next
                _tree.Nodes.Add(bridgeNode)
            Next
            If _tree.Nodes.Count > 0 Then
                SetChecked(_tree.Nodes(0), True)
                _tree.Nodes(0).Expand()
            End If
        Finally
            _tree.EndUpdate()
        End Try
    End Sub

    Private Sub TreeAfterCheck(sender As Object, e As TreeViewEventArgs)
        If _updatingTree Then Return
        _updatingTree = True
        Try
            For Each child As TreeNode In e.Node.Nodes
                SetChecked(child, e.Node.Checked)
            Next
            UpdateParentCheck(e.Node.Parent)
        Finally
            _updatingTree = False
        End Try
        ClearPatch()
        UpdateCommonValues()
    End Sub

    Private Sub SetChecked(node As TreeNode, value As Boolean)
        node.Checked = value
        For Each child As TreeNode In node.Nodes
            SetChecked(child, value)
        Next
    End Sub

    Private Sub UpdateParentCheck(node As TreeNode)
        If node Is Nothing Then Return
        node.Checked = node.Nodes.Cast(Of TreeNode)().All(Function(child) child.Checked)
        UpdateParentCheck(node.Parent)
    End Sub

    Private Function SelectedItems() As List(Of BridgeAppearanceItem)
        Dim result As New List(Of BridgeAppearanceItem)()
        For Each rootNode As TreeNode In _tree.Nodes
            CollectCheckedLeaves(rootNode, result)
        Next
        Return result
    End Function

    Private Sub CollectCheckedLeaves(node As TreeNode, result As List(Of BridgeAppearanceItem))
        If node.Nodes.Count = 0 Then
            Dim id As String = TryCast(node.Tag, String)
            If node.Checked AndAlso id IsNot Nothing AndAlso _items.ContainsKey(id) Then result.Add(_items(id))
            Return
        End If
        For Each child As TreeNode In node.Nodes
            CollectCheckedLeaves(child, result)
        Next
    End Sub

    Private Sub UpdateCommonValues()
        Dim selected As List(Of BridgeAppearanceItem) = SelectedItems()
        _loadingValues = True
        Try
            _status.Text = If(selected.Count = 0,
                              "Выберите элементы слева.",
                              "Отмечено флажками элементов: " & selected.Count)
            If selected.Count = 0 Then Return

            SetComboCurrent(_layerValue, CommonValue(selected, Function(item) item.Values.LayerName))
            SetComboCurrent(_linetypeValue, CommonValue(selected, Function(item) item.Values.LinetypeName))
            Dim scales = selected.Select(Function(item) item.Values.LinetypeScale).Distinct().ToList()
            _scaleCurrent.Text = If(scales.Count = 1, scales(0).ToString("G"), MixedText)
            If scales.Count = 1 Then SetNumericValue(_scaleValue, scales(0))

            Dim widthItems = selected.Where(Function(item) item.SupportsWidth).ToList()
            Dim widths = widthItems.Select(Function(item) item.Values.Width).Distinct().ToList()
            _widthCurrent.Text = CommonDisplay(widthItems,
                                               Function(item) If(item.Values.Width.HasValue, item.Values.Width.Value.ToString("G"), "—"))
            If widths.Count = 1 AndAlso widths(0).HasValue Then SetNumericValue(_widthValue, widths(0).Value)

            Dim bodyItems = selected.Where(Function(item) item.SupportsBody).ToList()
            Dim bodyColors = bodyItems.Select(Function(item) item.Values.BodyArgb).Distinct().ToList()
            _bodyCurrent.Text = CommonDisplay(bodyItems,
                                              Function(item) If(item.Values.BodyArgb.HasValue,
                                                                ColorName(item.Values.BodyArgb.Value), MixedText))
            If bodyColors.Count = 1 AndAlso bodyColors(0).HasValue Then
                _bodyArgbValue = bodyColors(0).Value
                _bodyChoiceValid = True
                _bodyColorButton.BackColor = Color.FromArgb(_bodyArgbValue)
            Else
                _bodyArgbValue = 0
                _bodyChoiceValid = False
            End If

            Dim fillItems = selected.Where(Function(item) item.SupportsFill).ToList()
            Dim fillColors = fillItems.Select(Function(item) item.Values.FillCadColorValue).Distinct().ToList()
            _fillCurrent.Text = CommonDisplay(fillItems,
                                              Function(item) If(item.Values.FillCadColorValue.HasValue,
                                                                CadColorName(CadColor.FromCompressValue(item.Values.FillCadColorValue.Value)),
                                                                MixedText))
            If fillColors.Count = 1 AndAlso fillColors(0).HasValue Then
                _fillCompressedValue = fillColors(0).Value
                _fillChoiceValid = True
                _fillColorButton.BackColor = CadColor.FromCompressValue(_fillCompressedValue).Win32Color
            Else
                _fillCompressedValue = New CadColor(5).ToCompressValue()
                _fillChoiceValid = False
            End If

            Dim hatchPatterns = fillItems.Select(Function(item) item.Values.HatchPatternName).
                Where(Function(value) value IsNot Nothing).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            _hatchPatternValue.SelectedIndex = -1
            If hatchPatterns.Count = 1 Then _hatchPatternValue.SelectedItem = hatchPatterns(0).ToUpperInvariant()

            Dim hatchScales = fillItems.Select(Function(item) item.Values.HatchScale).Distinct().ToList()
            _hatchScaleCurrent.Text = CommonDisplay(fillItems,
                                                    Function(item) If(item.Values.HatchScale.HasValue,
                                                                      item.Values.HatchScale.Value.ToString("G"), MixedText))
            If hatchScales.Count = 1 AndAlso hatchScales(0).HasValue Then
                SetNumericValue(_hatchScaleValue, hatchScales(0).Value)
            End If

            Dim hatchAngles = fillItems.Select(Function(item) item.Values.HatchAngle).Distinct().ToList()
            _hatchAngleCurrent.Text = CommonDisplay(fillItems,
                                                    Function(item) If(item.Values.HatchAngle.HasValue,
                                                                      item.Values.HatchAngle.Value.ToString("G"), MixedText))
            If hatchAngles.Count = 1 AndAlso hatchAngles(0).HasValue Then
                SetNumericValue(_hatchAngleValue, hatchAngles(0).Value)
            End If

            Dim cadValues = selected.Select(Function(item) item.Values.CadColorValue).Distinct().ToList()
            If cadValues.Count = 1 Then
                _cadCompressedValue = cadValues(0)
                _cadChoiceValid = True
                SetCadChoice(cadValues(0))
            Else
                _cadChoiceValid = False
                _cadValue.SelectedIndex = 0
            End If

            Dim weights = selected.Select(Function(item) item.Values.Lineweight).Distinct().ToList()
            _weightValue.SelectedIndex = -1
            If weights.Count = 1 Then
                For index As Integer = 0 To _weightValue.Items.Count - 1
                    Dim choice As LineweightChoice = TryCast(_weightValue.Items(index), LineweightChoice)
                    If choice IsNot Nothing AndAlso choice.Value = weights(0) Then
                        _weightValue.SelectedIndex = index
                        Exit For
                    End If
                Next
            End If
            _widthCheck.Enabled = selected.Any(Function(item) item.SupportsWidth)
            _bodyCheck.Enabled = selected.Any(Function(item) item.SupportsBody)
            Dim supportsFill As Boolean = selected.Any(Function(item) item.SupportsFill)
            _fillCheck.Enabled = supportsFill
            _hatchPatternCheck.Enabled = supportsFill
            _hatchScaleCheck.Enabled = supportsFill
            _hatchAngleCheck.Enabled = supportsFill
        Finally
            _loadingValues = False
            UpdateEditorsEnabled(Nothing, EventArgs.Empty)
        End Try
    End Sub

    Private Shared Function CommonValue(items As List(Of BridgeAppearanceItem), selector As Func(Of BridgeAppearanceItem, String)) As String
        Dim values = items.Select(selector).Distinct(StringComparer.Ordinal).ToList()
        Return If(values.Count = 1, values(0), MixedText)
    End Function

    Private Shared Function CommonDisplay(items As List(Of BridgeAppearanceItem), selector As Func(Of BridgeAppearanceItem, String)) As String
        If items.Count = 0 Then Return "Не применяется"
        Dim values = items.Select(selector).Distinct(StringComparer.Ordinal).ToList()
        Return If(values.Count = 1, values(0), MixedText)
    End Function

    Private Shared Sub SetComboCurrent(combo As ComboBox, value As String)
        combo.SelectedIndex = -1
        combo.Text = If(value, String.Empty)
    End Sub

    Private Shared Sub SetNumericValue(control As NumericUpDown, value As Double)
        If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then Return
        Dim decimalValue As Decimal
        Try
            decimalValue = CDec(value)
        Catch
            Return
        End Try
        control.Value = Math.Max(control.Minimum, Math.Min(control.Maximum, decimalValue))
    End Sub

    Private Sub SetCadChoice(value As Integer)
        Dim cadColor As CadColor = CadColor.FromCompressValue(value)
        If cadColor = CadColor.ByLayer Then
            _cadValue.SelectedIndex = 1
        ElseIf cadColor = CadColor.ByBlock Then
            _cadValue.SelectedIndex = 2
        Else
            _cadValue.SelectedIndex = 3
        End If
        _cadColorButton.BackColor = cadColor.Win32Color
        _cadColorButton.Text = CadColorName(cadColor)
    End Sub

    Private Sub ClearPatch()
        _cadCheck.Checked = False
        _layerCheck.Checked = False
        _linetypeCheck.Checked = False
        _scaleCheck.Checked = False
        _weightCheck.Checked = False
        _widthCheck.Checked = False
        _bodyCheck.Checked = False
        _fillCheck.Checked = False
        _hatchPatternCheck.Checked = False
        _hatchScaleCheck.Checked = False
        _hatchAngleCheck.Checked = False
    End Sub

    Private Sub UpdateEditorsEnabled(sender As Object, e As EventArgs)
        _cadValue.Enabled = _cadCheck.Checked
        _cadColorButton.Enabled = _cadCheck.Checked
        _layerValue.Enabled = _layerCheck.Checked
        _linetypeValue.Enabled = _linetypeCheck.Checked
        _scaleValue.Enabled = _scaleCheck.Checked
        _weightValue.Enabled = _weightCheck.Checked
        _widthValue.Enabled = _widthCheck.Checked AndAlso _widthCheck.Enabled
        _bodyColorButton.Enabled = _bodyCheck.Checked AndAlso _bodyCheck.Enabled
        _fillColorButton.Enabled = _fillCheck.Checked AndAlso _fillCheck.Enabled
        _hatchPatternValue.Enabled = _hatchPatternCheck.Checked AndAlso _hatchPatternCheck.Enabled
        _hatchScaleValue.Enabled = _hatchScaleCheck.Checked AndAlso _hatchScaleCheck.Enabled
        _hatchAngleValue.Enabled = _hatchAngleCheck.Checked AndAlso _hatchAngleCheck.Enabled
    End Sub

    Private Sub BodyCheckChanged(sender As Object, e As EventArgs)
        If Not _loadingValues AndAlso _bodyCheck.Checked AndAlso _bodyArgbValue <> 0 Then
            _bodyChoiceValid = True
        End If
        UpdateEditorsEnabled(sender, e)
    End Sub

    Private Sub FillCheckChanged(sender As Object, e As EventArgs)
        If Not _loadingValues AndAlso _fillCheck.Checked AndAlso _fillCompressedValue <> 0 Then
            _fillChoiceValid = True
        End If
        UpdateEditorsEnabled(sender, e)
    End Sub

    Private Sub CadChoiceChanged(sender As Object, e As EventArgs)
        If _loadingValues Then Return
        If _cadValue.SelectedIndex = 0 Then
            _cadChoiceValid = False
            _cadCheck.Checked = False
            Return
        End If
        _cadCheck.Checked = True
        If _cadValue.SelectedIndex = 1 Then
            _cadCompressedValue = CadColor.ByLayer.ToCompressValue()
            _cadChoiceValid = True
        ElseIf _cadValue.SelectedIndex = 2 Then
            _cadCompressedValue = CadColor.ByBlock.ToCompressValue()
            _cadChoiceValid = True
        End If
    End Sub

    Private Sub ChooseCadColor(sender As Object, e As EventArgs)
        Dim cadColor As CadColor = CadColor.FromCompressValue(_cadCompressedValue)
        If Not ColorsBrowserDlg.Execute(cadColor) Then Return
        _cadCompressedValue = cadColor.ToCompressValue()
        _cadChoiceValid = True
        _cadColorButton.BackColor = cadColor.Win32Color
        SetCadChoice(_cadCompressedValue)
        _cadCheck.Checked = True
    End Sub

    Private Sub ChooseBodyColor(sender As Object, e As EventArgs)
        Using dialog As New ColorDialog With {.FullOpen = True, .Color = _bodyColorButton.BackColor}
            If dialog.ShowDialog(Me) <> DialogResult.OK Then Return
            _bodyArgbValue = dialog.Color.ToArgb()
            _bodyChoiceValid = True
            _bodyColorButton.BackColor = dialog.Color
            _bodyCheck.Checked = True
        End Using
    End Sub

    Private Sub ChooseFillColor(sender As Object, e As EventArgs)
        Dim fillColor As CadColor = CadColor.FromCompressValue(_fillCompressedValue)
        If Not ColorsBrowserDlg.Execute(fillColor) Then Return
        _fillCompressedValue = fillColor.ToCompressValue()
        _fillChoiceValid = True
        _fillColorButton.BackColor = fillColor.Win32Color
        _fillCheck.Checked = True
    End Sub

    Private Sub ApplyChanges(closeAfterApply As Boolean)
        Dim selected As List(Of BridgeAppearanceItem) = SelectedItems()
        If selected.Count = 0 Then
            _status.Text = "Выберите хотя бы один элемент."
            Return
        End If

        Dim patch As New BridgeAppearancePatch()
        If _cadCheck.Checked AndAlso _cadChoiceValid Then patch.CadColorValue = _cadCompressedValue
        If _layerCheck.Checked Then patch.LayerName = _layerValue.Text
        If _linetypeCheck.Checked Then patch.LinetypeName = _linetypeValue.Text
        If _scaleCheck.Checked Then patch.LinetypeScale = CDbl(_scaleValue.Value)
        If _weightCheck.Checked Then
            Dim weightChoice As LineweightChoice = TryCast(_weightValue.SelectedItem, LineweightChoice)
            If weightChoice IsNot Nothing Then patch.Lineweight = weightChoice.Value
        End If
        If _widthCheck.Checked Then patch.Width = CDbl(_widthValue.Value)
        If _bodyCheck.Checked AndAlso _bodyChoiceValid Then patch.BodyArgb = _bodyArgbValue
        If _fillCheck.Checked AndAlso _fillChoiceValid Then patch.FillCadColorValue = _fillCompressedValue
        If _hatchPatternCheck.Checked AndAlso _hatchPatternValue.SelectedItem IsNot Nothing Then
            patch.HatchPatternName = _hatchPatternValue.SelectedItem.ToString()
        End If
        If _hatchScaleCheck.Checked Then patch.HatchScale = CDbl(_hatchScaleValue.Value)
        If _hatchAngleCheck.Checked Then patch.HatchAngle = CDbl(_hatchAngleValue.Value)
        If patch.IsEmpty Then
            _status.Text = "Отметьте свойства, которые нужно изменить."
            Return
        End If

        Dim reason As String = Nothing
        If Not patch.TryValidate(reason) Then
            _status.Text = reason
            Return
        End If

        Dim args As New BridgeAppearanceApplyEventArgs(selected.Select(Function(item) item.SelectionId), patch)
        RaiseEvent ApplyRequested(Me, args)
        If Not String.IsNullOrWhiteSpace(args.ErrorMessage) Then
            _status.Text = args.ErrorMessage
            Return
        End If

        Dim skippedBodyLabels As IReadOnlyList(Of String) = If(args.SkippedBodyLabels, New String() {})
        Dim resultStatus As String
        If skippedBodyLabels.Count > 0 Then
            resultStatus = BodyApplyStatus(args.AppliedCount, skippedBodyLabels)
        Else
            resultStatus = "Изменено элементов: " & args.AppliedCount
        End If
        ClearPatch()
        UpdateCommonValues()
        _status.Text = resultStatus
        If skippedBodyLabels.Count > 0 Then Return
        If closeAfterApply Then
            DialogResult = DialogResult.OK
            Close()
        End If
    End Sub

    Private Shared Function BodyApplyStatus(appliedCount As Integer,
                                            skippedLabels As IReadOnlyList(Of String)) As String
        Dim visibleLabels As String = String.Join(", ", skippedLabels.Take(5))
        If skippedLabels.Count > 5 Then visibleLabels &= ", ещё " & (skippedLabels.Count - 5)
        Return "Изменено элементов: " & appliedCount &
               ". Цвет тела не применён к части выбранных 3D-элементов; пропущено: " &
               skippedLabels.Count & " (" & visibleLabels & ")."
    End Function

    Private Shared Function ColorName(argb As Integer) As String
        Dim color As Color = Color.FromArgb(argb)
        Return String.Format("RGB {0}, {1}, {2}", color.R, color.G, color.B)
    End Function

    Private Shared Function CadColorName(color As CadColor) As String
        If color = CadColor.ByLayer Then Return "По слою"
        If color = CadColor.ByBlock Then Return "По блоку"
        If color = CadColor.Empty Then Return "Пустой"
        If color = CadColor.Background Then Return "Фон"
        If color.ColorIndex >= 0 Then Return "Индекс " & color.ColorIndex
        Return ColorName(color.Win32Color.ToArgb())
    End Function

    Private NotInheritable Class LineweightChoice
        Public Sub New(value As Integer)
            Me.Value = value
        End Sub

        Public ReadOnly Property Value As Integer

        Public Overrides Function ToString() As String
            Select Case Value
                Case -3
                    Return "По умолчанию"
                Case -2
                    Return "По блоку"
                Case -1
                    Return "По слою"
                Case Else
                    Return (Value / 100.0).ToString("0.00") & " мм"
            End Select
        End Function
    End Class
End Class
