Imports System.Drawing
Imports System.Windows.Forms
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class PanelPropertiesObject
    Public userCadView As CadView = Nothing
    Public templateXML As String = ""
    Public boolselectCombo As Boolean = False
    Public clickButtonGlobalProperties As Boolean = True
    Public clickButtonGLocalProperties As Boolean = True
    Public dataStructureElement As StructureElement = Nothing
    Public selectDwgObject As DwgObject = Nothing
    Private cellValueBeforeEdit As Object = Nothing
    Private suppressPropertiesGridEvents As Boolean = False
    Public Sub New()
        InitializeComponent()
        Column1.ReadOnly = True
        DataGrid_PropertiesEnt.AllowUserToAddRows = False
        SetGlobalPropertiesMode()
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим каталог с файлами (c:\Users\mzaharov\AppData\Roaming\Civil3DToolsUtility\TOPOTemplates\)
        templateXML = FuncFiles.getFileToDirectorySupport("Стандарт_ADSK.xml")
        If IsNothing(templateXML) = True Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
    End Sub

    Public Sub BeginPropertiesGridUpdate()
        DataGrid_PropertiesEnt.EndEdit()
        suppressPropertiesGridEvents = True
    End Sub

    Public Sub EndPropertiesGridUpdate()
        suppressPropertiesGridEvents = False
    End Sub

    Public Sub SetGlobalPropertiesMode()
        clickButtonGlobalProperties = True
        Button_GlobalProperties.BackColor = Color.Green
        Button_LocalProperties.BackColor = Color.Red
        Column1.Width = 110
        Column2.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Column2.Width = 160
        Column2.DefaultCellStyle.FormatProvider = Globalization.CultureInfo.CurrentCulture
    End Sub

    Public Sub SetLocalPropertiesMode()
        clickButtonGlobalProperties = False
        Button_GlobalProperties.BackColor = Color.Red
        Button_LocalProperties.BackColor = Color.Green
        Column1.Width = 200
        Column2.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Column2.Width = 70
        Column2.DefaultCellStyle.FormatProvider = Globalization.CultureInfo.InvariantCulture
    End Sub
    'смена таблицы
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_DataTable.SelectedIndexChanged
        If boolselectCombo = False Then
            'Dim nameSelectTable As String = ComboBox1.Text.Trim
            'DataGridView1.Rows.Clear()
            'If IsNothing(userCadView) = False Then
            '    If userCadView.SelectionSet.Count = 1 Then
            '        For Each acEnt As Object In userCadView.SelectionSet
            '            If TypeOf acEnt Is DwgEntity Then
            '                Dim acObj As DwgEntity = acEnt
            '                Dim arrayData As String(,) = Nothing
            '                Dim boolFindProperties As Boolean = FuncXRecords.FuncReadXData(acObj, nameSelectTable, arrayData)
            '                If IsArray(arrayData) = True Then
            '                    Dim countRow As Integer = 0
            '                    For i As Integer = 0 To arrayData.GetUpperBound(1)
            '                        Dim nameF As String = arrayData(0, i)
            '                        Dim valF As String = arrayData(1, i)
            '                        Dim desk As String = arrayData(2, i)
            '                        If IsNothing(nameF) = False Then
            '                            If nameF.Trim.Length > 0 Then
            '                                If IsNothing(valF) = True Then valF = ""
            '                                DataGridView1.Rows.Add()
            '                                DataGridView1.Rows(countRow).Cells(0).Value = nameF
            '                                DataGridView1.Rows(countRow).Cells(1).Value = valF
            '                                DataGridView1.Rows(countRow).Cells(0).Tag = desk
            '                                countRow += 1
            '                            End If
            '                        End If
            '                    Next i
            '                    DataGridView1.Rows(countRow).Cells(0).Value = "ObjectID"
            '                    DataGridView1.Rows(countRow).Cells(1).Value = acObj.ObjectID
            '                    DataGridView1.Rows(countRow).Cells(0).Tag = ""
            '                    Exit For
            '                End If
            '            End If
            '        Next
            '    Else
            '        DataGridView1.Rows.Clear()
            '    End If
            'End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'перезаписать значение в ячейке
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGrid_PropertiesEnt.CellBeginEdit
        If suppressPropertiesGridEvents Then Exit Sub
        If e.RowIndex < 0 OrElse e.ColumnIndex <> 1 Then Exit Sub
        If DataGrid_PropertiesEnt.Rows(e.RowIndex).IsNewRow Then Exit Sub
        cellValueBeforeEdit = DataGrid_PropertiesEnt.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_PropertiesEnt.CellEndEdit
        If suppressPropertiesGridEvents Then Exit Sub
        Dim rowIndex As Integer = e.RowIndex
        If rowIndex < 0 OrElse e.ColumnIndex <> 1 Then Exit Sub
        If DataGrid_PropertiesEnt.Rows(rowIndex).IsNewRow Then Exit Sub
        Dim editedValue As Object = DataGrid_PropertiesEnt.Rows(rowIndex).Cells(e.ColumnIndex).Value
        If Object.Equals(cellValueBeforeEdit, editedValue) Then Exit Sub
        If IsNothing(dataStructureElement) = True Then Exit Sub
        If IsNothing(selectDwgObject) = True Then Exit Sub
        If rowIndex > -1 Then
            'редактируются глобальные свойства
            Dim nameValue As Object = DataGrid_PropertiesEnt.Rows(rowIndex).Cells(0).Value
            If IsNothing(nameValue) Then Exit Sub
            Dim nameF As String = nameValue.ToString()
            Dim valF As String = If(IsNothing(editedValue), "", editedValue.ToString())
            If clickButtonGlobalProperties = True Then
                If nameF Like "Label" Or nameF Like "LABEL" Then
                    dataStructureElement.Label = valF
                ElseIf nameF Like "ClassBridgeObject" Then
                    Try
                        dataStructureElement.ClassBridgeObject = CType([Enum].Parse(GetType(StructureElement.classBridge), valF), StructureElement.classBridge)
                    Catch ex As System.ArgumentException
                    End Try
                ElseIf nameF Like "ClassStructure" OrElse nameF Like "ClassObject" Then
                    Try
                        dataStructureElement.ClassObject = CType([Enum].Parse(GetType(StructureElement.classStructure), valF), StructureElement.classStructure)
                    Catch ex As System.ArgumentException
                    End Try
                ElseIf nameF Like "Name" Or nameF Like "NAME" Then
                    Try
                        dataStructureElement.Name = CType([Enum].Parse(GetType(StructureElement.typeObject), valF), StructureElement.typeObject)
                    Catch ex As System.ArgumentException
                    End Try
                ElseIf nameF Like "Description" Then
                    dataStructureElement.Description = valF
                ElseIf nameF Like "KeyParameter" OrElse nameF Like "KeyParameters" Then
                    dataStructureElement.KeyParameter = valF
                ElseIf nameF Like "IdStructure" Then
                    dataStructureElement.IdStructure = valF
                ElseIf nameF Like "IdElement" Then
                    dataStructureElement.IdElement = valF
                ElseIf nameF Like "Note" Or nameF Like "NOTE" Then
                    dataStructureElement.Note = valF
                End If
            Else
                'редактируются параметры json
                Dim newKeyParameter As String = dataStructureElement.KeyParameter
                Dim boolSetParam As Boolean = FuncGSON.setTypedValue(newKeyParameter, nameF, valF)
                If boolSetParam = True Then
                    dataStructureElement.KeyParameter = newKeyParameter
                Else
                    DataGrid_PropertiesEnt.Rows(rowIndex).Cells(1).Style.ForeColor = Color.Red
                    MsgBox("Не удалось изменить параметр " & nameF & ". Проверьте имя параметра и тип введённого значения.")
                    Exit Sub
                End If
            End If
            Dim boolWriteData As Boolean = FuncXRecords.setXRecords(selectDwgObject, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataStructureElement)
            If boolWriteData = True Then
                DataGrid_PropertiesEnt.Rows(rowIndex).Cells(1).Style.ForeColor = Color.Green
            Else
                DataGrid_PropertiesEnt.Rows(rowIndex).Cells(1).Style.ForeColor = Color.Red
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'посмотреть глобальные свойства объекта
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_GlobalProperties.Click
        BeginPropertiesGridUpdate()
        Try
            SetGlobalPropertiesMode()
            selectDwgObject = Nothing
            dataStructureElement = Nothing
            DataGrid_PropertiesEnt.Rows.Clear()
            If IsNothing(userCadView) = False Then
                For Each acEnt As Object In userCadView.SelectionSet
                    If IsNothing(acEnt) = False Then
                        If TypeOf acEnt Is DwgObject Then
                            Dim selectedDwgObject As DwgObject = acEnt
                            If IsNothing(selectedDwgObject) = False Then
                                Dim selectedDataStructureElement As StructureElement = New StructureElement
                                Dim boolFindData As Boolean = FuncXRecords.getXRecords(selectedDwgObject, selectedDataStructureElement, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                If boolFindData = True Then
                                    selectDwgObject = selectedDwgObject
                                    dataStructureElement = selectedDataStructureElement
                                    DataGrid_PropertiesEnt.Rows.Add(9)
                                    DataGrid_PropertiesEnt.Rows(0).Cells(0).Value = "LABEL"
                                    DataGrid_PropertiesEnt.Rows(0).Cells(1).Value = selectedDataStructureElement.Label

                                    DataGrid_PropertiesEnt.Rows(1).Cells(0).Value = "ClassBridgeObject"
                                    DataGrid_PropertiesEnt.Rows(1).Cells(1).Value = selectedDataStructureElement.ClassBridgeObject

                                    DataGrid_PropertiesEnt.Rows(2).Cells(0).Value = "ClassObject"
                                    DataGrid_PropertiesEnt.Rows(2).Cells(1).Value = selectedDataStructureElement.ClassObject

                                    DataGrid_PropertiesEnt.Rows(3).Cells(0).Value = "Name"
                                    DataGrid_PropertiesEnt.Rows(3).Cells(1).Value = selectedDataStructureElement.Name

                                    DataGrid_PropertiesEnt.Rows(4).Cells(0).Value = "Description"
                                    DataGrid_PropertiesEnt.Rows(4).Cells(1).Value = selectedDataStructureElement.Description

                                    DataGrid_PropertiesEnt.Rows(5).Cells(0).Value = "KeyParameters"
                                    DataGrid_PropertiesEnt.Rows(5).Cells(1).Value = selectedDataStructureElement.KeyParameter

                                    DataGrid_PropertiesEnt.Rows(6).Cells(0).Value = "IdStructure"
                                    DataGrid_PropertiesEnt.Rows(6).Cells(1).Value = selectedDataStructureElement.IdStructure

                                    DataGrid_PropertiesEnt.Rows(7).Cells(0).Value = "IdElement"
                                    DataGrid_PropertiesEnt.Rows(7).Cells(1).Value = selectedDataStructureElement.IdElement

                                    DataGrid_PropertiesEnt.Rows(8).Cells(0).Value = "NOTE"
                                    DataGrid_PropertiesEnt.Rows(8).Cells(1).Value = selectedDataStructureElement.Note
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Finally
            EndPropertiesGridUpdate()
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'посмотреть локальные свойства объекта
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button_LocalProperties.Click
        BeginPropertiesGridUpdate()
        Try
            SetLocalPropertiesMode()
            selectDwgObject = Nothing
            dataStructureElement = Nothing
            DataGrid_PropertiesEnt.Rows.Clear()
            If IsNothing(userCadView) = False Then
                For Each acEnt As Object In userCadView.SelectionSet
                    If IsNothing(acEnt) = False Then
                        If TypeOf acEnt Is DwgObject Then
                            Dim selectedDwgObject As DwgObject = acEnt
                            If IsNothing(selectedDwgObject) = False Then
                                Dim selectedDataStructureElement As StructureElement = New StructureElement
                                Dim boolFindData As Boolean = FuncXRecords.getXRecords(selectedDwgObject, selectedDataStructureElement, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                If boolFindData = True Then
                                    Dim keyParameter As String = selectedDataStructureElement.KeyParameter
                                    If FuncGSON.IsValidJson(keyParameter) = True Then
                                        Dim dictionaryDataGSon As Dictionary(Of String, Object) = Nothing
                                        Try
                                            dictionaryDataGSon = FuncGSON.ParseJsonToDictionary(keyParameter)
                                        Catch ex As ArgumentException
                                        End Try
                                        If dictionaryDataGSon IsNot Nothing AndAlso dictionaryDataGSon.Count > 0 Then
                                            selectDwgObject = selectedDwgObject
                                            dataStructureElement = selectedDataStructureElement
                                            DataGrid_PropertiesEnt.Rows.Add(dictionaryDataGSon.Count)
                                            For i As Integer = 0 To dictionaryDataGSon.Count - 1
                                                Dim nameField As String = dictionaryDataGSon.ElementAt(i).Key
                                                Dim value As Object = dictionaryDataGSon.ElementAt(i).Value
                                                DataGrid_PropertiesEnt.Rows(i).Cells(0).Value = nameField
                                                DataGrid_PropertiesEnt.Rows(i).Cells(1).Value = value
                                            Next i
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Finally
            EndPropertiesGridUpdate()
        End Try
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGrid_PropertiesEnt.CellFormatting
        If (e.ColumnIndex = 0 AndAlso e.Value IsNot Nothing) Then
            Dim cell As DataGridViewCell = Me.DataGrid_PropertiesEnt.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim tagStr As String = cell.Tag
            Dim strVal As String = cell.Value
            If IsNothing(tagStr) = False Then
                cell.ToolTipText = tagStr
            Else
                cell.ToolTipText = strVal
            End If
        End If
    End Sub

    Private Sub PanelPropertiesObject_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim sizeForm As System.Drawing.Size = Me.Size
        Dim formWight As Double = sizeForm.Width
        Dim formHeight As Double = sizeForm.Height
        If formHeight < 200 Then formHeight = 200
        If formWight < 200 Then formWight = 200
        'рамка
        Dim sizeGroupBox As System.Drawing.Size = New System.Drawing.Size(formWight - 10, formHeight - 10)
        GroupBox1.Size = sizeGroupBox
        'листищч
        Dim sizeDatagrid As System.Drawing.Size = New System.Drawing.Size(formWight - 25, formHeight - 25)
        DataGrid_PropertiesEnt.Size = sizeDatagrid
    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub
End Class
