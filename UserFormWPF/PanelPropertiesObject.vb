Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports System.Xml
Imports Topomatic.Alg
Imports Topomatic.Alg.Crs
Imports Topomatic.Alg.Kilometres
Imports Topomatic.Alg.Plan
Imports Topomatic.Alg.Road
Imports Topomatic.Alg.Road.Crossing
Imports Topomatic.Alg.Road.Urb
Imports Topomatic.Alg.Road.Urb.Border
Imports Topomatic.Alg.Road.Urb.Descent
Imports Topomatic.Alg.Road.Urb.UserStrips
Imports Topomatic.Alg.Runtime
Imports Topomatic.Alg.Runtime.ServiceClasses
Imports Topomatic.Alg.Stationing
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Planchet.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Visualization.Runtime.Tools

Public Class PanelPropertiesObject
    Public userCadView As CadView = Nothing
    Public templateXML As String = ""
    Public boolselectCombo As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим каталог с файлами (c:\Users\mzaharov\AppData\Roaming\Civil3DToolsUtility\TOPOTemplates\)
        templateXML = FuncFiles.getFileToDirectorySupport("Стандарт_ADSK.xml")
        If IsNothing(templateXML) = True Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
    End Sub
    'смена таблицы
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
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
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim rowIndex As Integer = e.RowIndex
        If rowIndex > -1 Then
            Dim userDict As String = ComboBox1.Text
            If IsNothing(userDict) = True Then Exit Sub
            If userDict.Trim.Length > 0 Then
                Dim xData As StructureElement = New StructureElement
                For i As Integer = 0 To DataGridView1.RowCount - 1
                    Dim nameF As String = DataGridView1.Rows(i).Cells(0).Value
                    Dim valF As String = DataGridView1.Rows(i).Cells(1).Value
                    If IsNothing(nameF) = True Then
                        Continue For
                    End If
                    If nameF.Trim.Length = 0 Then
                        Continue For
                    End If
                    If IsNothing(valF) = True Then
                        valF = ""
                    End If
                    If nameF Like "Label" Or nameF Like "LABEL" Then
                        xData.Label = valF
                    ElseIf nameF Like "ClassStructure" Then
                        Try
                            xData.ClassObject = CType([Enum].Parse(GetType(StructureElement.classStructure), valF), StructureElement.classStructure)
                        Catch ex As System.ArgumentException
                            xData.ClassObject = StructureElement.classStructure.OtherObject
                        End Try
                    ElseIf nameF Like "Name" Or nameF Like "NAME" Then
                        Try
                            xData.Name = CType([Enum].Parse(GetType(StructureElement.typeObject), valF), StructureElement.typeObject)
                        Catch ex As System.ArgumentException
                            xData.ClassObject = StructureElement.typeObject.OtherElement
                        End Try
                    ElseIf nameF Like "Description" Then
                        xData.Description = valF
                    ElseIf nameF Like "KeyParameter" Then
                        xData.KeyParameter = valF
                    ElseIf nameF Like "IdStructure" Then
                        xData.IdStructure = valF
                    ElseIf nameF Like "IdElement" Then
                        xData.IdElement = valF
                    ElseIf nameF Like "Note" Or nameF Like "NOTE" Then
                        xData.Note = valF
                    End If
                Next i
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf acEnt Is DwgEntity Then
                        Dim acObj As DwgEntity = acEnt
                        If IsNothing(acObj) = False Then
                            xData.IdObject = acObj.ObjectID.ToString

                            Dim boolWriteData As Boolean = FuncXRecords.setXRecords(acObj, StructureElement.tableXRecords.PROJECT_STRUCTURES, xData)
                            If boolWriteData = True Then
                                DataGridView1.Rows(rowIndex).Cells(1).Style.ForeColor = Color.Green
                            Else
                                DataGridView1.Rows(rowIndex).Cells(1).Style.ForeColor = Color.Red
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'подключить таблицу
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim arrayTablwsCombo As String() = ComboBox1.Items.Cast(Of String).ToArray()
        Dim count As Integer = 0
        If IsArray(arrayTablwsCombo) = True Then
            count = arrayTablwsCombo.Length
        End If
        If IO.File.Exists(templateXML) = True Then
            Dim arrayTables As String() = Nothing
            'забираем в массив все таблицы
            Dim boolFindTables As Boolean = FuncXML.FuncReadPSTablesToXML(templateXML, arrayTables)
            'массив с выбранными таблицами
            Dim arraySelectTables As String() = Nothing
            Dim countarraySelectTables As Integer = 0
            If IsArray(arrayTables) = True Then
                Dim formSelect As FormSelectCheckBox = New FormSelectCheckBox
                formSelect.CheckedListBox1.DataSource = arrayTables
                formSelect.ShowDialog()
                If formSelect.boolShow = True Then
                    For i As Integer = 0 To formSelect.CheckedListBox1.Items.Count - 1
                        If formSelect.CheckedListBox1.GetItemChecked(i) Then
                            Dim item = formSelect.CheckedListBox1.Items.Item(i)
                            ReDim Preserve arraySelectTables(countarraySelectTables)
                            arraySelectTables(countarraySelectTables) = item
                            countarraySelectTables += 1
                        End If
                    Next i
                End If
                If IsArray(arraySelectTables) = True Then
                    For i As Integer = 0 To arraySelectTables.Length - 1
                        Dim nameTablePS As String = arraySelectTables(i)
                        ReDim Preserve arrayTablwsCombo(count)
                        arrayTablwsCombo(count) = nameTablePS
                        count += 1
                        For Each acEnt As Object In userCadView.SelectionSet
                            If TypeOf acEnt Is DwgObject Then
                                Dim acObj As DwgObject = acEnt
                                Dim arrayTempTablePS As String() = Nothing
                                Dim boolFindData As Boolean = FuncXRecords.FuncReadTablesPS(acEnt, arrayTempTablePS)
                                If IsArray(arrayTempTablePS) = True Then
                                    Dim boolFindTable As Boolean = False
                                    For j As Integer = 0 To arrayTempTablePS.Count - 1
                                        Dim tb As String = arrayTempTablePS(j)
                                        If IsNothing(tb) = False Then
                                            If tb.Trim Like nameTablePS Then
                                                boolFindTable = True
                                                Exit For
                                            End If
                                        End If
                                    Next j
                                    If boolFindTable = False Then
                                        Dim arrayData As String(,) = Nothing
                                        Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, nameTablePS, arrayData, templateXML)
                                    End If
                                Else
                                    Dim arrayData As String(,) = Nothing
                                    Dim boolWriteData As Boolean = FuncXRecords.FuncCreateXData(acEnt, nameTablePS, arrayData, templateXML)
                                End If
                            End If
                        Next
                    Next i
                End If
                If IsArray(arrayTablwsCombo) = True Then
                    ComboBox1.DataSource = arrayTablwsCombo
                End If
            Else
                MsgBox("Файл-шаблон " & templateXML & " не найден!!! Макро прервано.")
                Exit Sub
            End If
        End If
        boolselectCombo = False
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'исключить таблицу
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim arrayTablwsCombo As String() = ComboBox1.Items.Cast(Of String).ToArray()
        Dim collTables As Windows.Forms.ComboBox.ObjectCollection = ComboBox1.Items
        Dim count As Integer = 0
        Dim arrayDeleteTables As String() = Nothing
        Dim countArrayDeleteTables As Integer = 0
        If collTables.Count > 0 Then
            Dim arrayTables As String() = Nothing
            Dim countArrayTables As Integer = 0
            For i As Integer = 0 To collTables.Count - 1
                ReDim Preserve arrayTables(countArrayTables)
                arrayTables(countArrayTables) = collTables.Item(i).ToString
                countArrayTables += 1
            Next i
            If IsArray(arrayTables) = True Then
                Dim formSelect As FormSelectCheckBox = New FormSelectCheckBox
                formSelect.CheckedListBox1.DataSource = arrayTables
                formSelect.ShowDialog()
                If formSelect.boolShow = True Then
                    For i As Integer = 0 To formSelect.CheckedListBox1.Items.Count - 1
                        If formSelect.CheckedListBox1.GetItemChecked(i) Then
                            Dim item = formSelect.CheckedListBox1.Items.Item(i)
                            ReDim Preserve arrayDeleteTables(countArrayDeleteTables)
                            arrayDeleteTables(countArrayDeleteTables) = item
                            countArrayDeleteTables += 1
                        End If
                    Next i
                End If
            End If
        End If
        If IsArray(arrayDeleteTables) = True Then
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is DwgObject Then
                    Dim acObj As DwgObject = acEnt
                    Dim boolDeleteDict As Boolean = FuncXRecords.FuncdeleteDictionaryXData(acObj, arrayDeleteTables)
                    If boolDeleteDict = True Then
                        Dim newArrayCombo As String() = Nothing
                        Dim countNewArrayCombo As Integer = 0
                        For i As Integer = 0 To arrayTablwsCombo.Length - 1
                            Dim tb1 As String = arrayTablwsCombo(i)
                            Dim boolFindtb As Boolean = False
                            For j As Integer = 0 To arrayDeleteTables.Length - 1
                                Dim tb2 As String = arrayDeleteTables(j)
                                If tb1.Trim Like tb2.Trim Then
                                    boolFindtb = True
                                End If
                                If boolFindtb = False Then
                                    ReDim Preserve newArrayCombo(countNewArrayCombo)
                                    newArrayCombo(countNewArrayCombo) = tb1
                                    countNewArrayCombo += 1
                                End If
                            Next j
                        Next i
                        If IsArray(newArrayCombo) = False Then
                            ReDim Preserve newArrayCombo(0)
                            newArrayCombo(0) = ""
                        End If
                        ComboBox1.DataSource = newArrayCombo
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If (e.ColumnIndex = 0 AndAlso e.Value IsNot Nothing) Then
            Dim cell As DataGridViewCell = Me.DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim tagStr As String = cell.Tag
            Dim strVal As String = cell.Value
            If IsNothing(tagStr) = False Then
                cell.ToolTipText = tagStr
            Else
                cell.ToolTipText = strVal
            End If
        End If
    End Sub

    Private Sub PanelPropertiesObject_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        DataGridView1.Size = sizeDatagrid
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.ColumnIndex = 1 Then
            Dim nameTablePS As String = ComboBox1.Text.Trim
            Dim nameField As String = Me.DataGridView1.Rows(e.RowIndex).Cells(0).Value
            'читаем xml
            Dim arrayDataValue As String() = Nothing
            Dim boolFindData As Boolean = FuncXML.FuncReadFieldValueByPSTablesToXML(templateXML, nameTablePS, nameField, arrayDataValue)
            If IsArray(arrayDataValue) = True Then
                Dim arrayRezultData As String() = Nothing
                Dim countArrayRezultData As Integer = 0
                For i As Integer = 0 To arrayDataValue.Length - 1
                    If arrayDataValue(i).Trim.Length > 0 Then
                        ReDim Preserve arrayRezultData(countArrayRezultData)
                        arrayRezultData(countArrayRezultData) = arrayDataValue(i).Trim
                        countArrayRezultData += 1
                    End If
                Next
                If IsArray(arrayRezultData) = True Then
                    Dim formSelectObject As FormSelectObject = New FormSelectObject
                    formSelectObject.ListBox1.DataSource = arrayRezultData
                    formSelectObject.ShowDialog()
                    If formSelectObject.boolShow = True Then
                        Me.DataGridView1.Rows(e.RowIndex).Cells(1).Value = formSelectObject.ListBox1.Text
                    End If
                End If
            End If
        End If
    End Sub
End Class
