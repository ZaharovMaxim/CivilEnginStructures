Imports System.Windows.Forms

Partial Public Class FormPlacementBeams
    Private Function FuncCreateTables() As Boolean
        FuncCreateTables = False
        Dim numberProlet As Integer = NUpD_CountProlet.Value

        Dim countRowsLeftBeam As Integer = NUpD_CountLeftRows.Value 'количество
        Dim offsetLeftFirstBeam As Integer = NumericUpDown3.Value  'смещение первой оси влево
        Dim offsetLeftLastBeam As Integer = NumericUpDown5.Value 'смещение последней оси влево

        Dim countRowsRightBeam As Integer = NUpD_CountRightRows.Value 'количество
        Dim offsetRightFirstBeam As Integer = NumericUpDown8.Value  'смещение первой оси
        Dim offsetRightLastBeam As Integer = NumericUpDown9.Value  'смещение последней оси
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создаем таблицу опор
        If numberProlet > 0 Then
            Dim countRow As Integer = DG_PillarsProperties.RowCount
            If countRow > 0 Then
                DG_PillarsProperties.Rows.Clear()
            End If
            countRow = DG_ProletListBeams.RowCount
            If countRow > 0 Then
                DG_ProletListBeams.Rows.Clear()
            End If
            countRow = DG_RowProperties.RowCount
            If countRow > 0 Then
                DG_RowProperties.Rows.Clear()
            End If
            While DG_ProletListBeams.ColumnCount > numberProlet
                DG_ProletListBeams.Columns.RemoveAt(DG_ProletListBeams.ColumnCount - 1)
            End While
            DG_PillarsProperties.Rows.Add(numberProlet + 1)
            For i As Integer = 0 To numberProlet - 1
                'заполняем таблицу во вкладке пролеты
                If i > DG_ProletListBeams.ColumnCount - 1 Then
                    Dim newcol = New DataGridViewComboBoxColumn()
                    newcol.HeaderText = "Пролет " & i + 1
                    newcol.FlatStyle = FlatStyle.Flat
                    newcol.Name = "Prolet" & i
                    newcol.Width = 120
                    newcol.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                    Dim intColl As Integer = DG_ProletListBeams.Columns.Add(newcol)
                    DG_ProletListBeams.Columns(intColl).Tag = i + 1
                Else
                    DG_ProletListBeams.Columns.Item(i).HeaderText = "Пролет " & i + 1
                    DG_ProletListBeams.Columns.Item(i).Name = "Prolet" & i
                    DG_ProletListBeams.Columns.Item(i).Tag = i + 1
                End If
                'заполняем таблицу во вкладке опоры
                DG_PillarsProperties.Rows(i).HeaderCell.Value = "Опора №" & i + 1
                DG_PillarsProperties.Rows(i).Tag = i + 1
                If i = 0 Then
                    Dim chCell As DataGridViewCheckBoxCell = DG_PillarsProperties.Rows(i).Cells(0)
                    chCell.Value = True
                End If
                DG_PillarsProperties.Rows(i).Cells(1).Value = 0
                DG_PillarsProperties.Rows(i).Cells(2).Value = 0
            Next i
            DG_PillarsProperties.Rows(numberProlet).HeaderCell.Value = "Опора №" & numberProlet + 1
            DG_PillarsProperties.Rows(numberProlet).Tag = numberProlet + 1
            DG_PillarsProperties.Rows(numberProlet).Cells(1).Value = 0
            DG_PillarsProperties.Rows(numberProlet).Cells(2).Value = 0
            DG_PillarsProperties.Rows(numberProlet).Cells(3).Value = 0
        End If
        Dim countRows As Integer = countRowsLeftBeam + countRowsRightBeam
        If ChB_CenterBeam.Checked = True Then
            countRows += 1
        End If
        If countRows > 0 Then
            DG_ProletListBeams.Rows.Add(countRows)
            DG_RowProperties.Rows.Add(countRows)
        End If
        Dim count As Integer = 0
        Dim count2 As Integer = 0
        If countRowsLeftBeam > 0 Then
            For i As Integer = countRowsLeftBeam - 1 To 0 Step -1
                DG_ProletListBeams.Rows(count).HeaderCell.Value = "Ряд Л-" & i + 1
                DG_ProletListBeams.Rows(count).Tag = -1 * (i + 1)
                count += 1
            Next i

            Dim deltaLenghtRow As Double = 0.0R
            If countRowsLeftBeam > 1 Then
                deltaLenghtRow = (offsetLeftLastBeam - offsetLeftFirstBeam) / (countRowsLeftBeam - 1)
            End If
            count2 = countRowsLeftBeam
            For i As Integer = 0 To countRowsLeftBeam - 1
                Dim tempDist As Integer = If(countRowsLeftBeam = 1,
                                             offsetLeftFirstBeam,
                                             offsetLeftLastBeam - deltaLenghtRow * i)
                DG_RowProperties.Rows(i).HeaderCell.Value = "Ряд Л-" & countRowsLeftBeam - i
                DG_RowProperties.Rows(i).Tag = -1 * (countRowsLeftBeam - i)
                DG_RowProperties.Rows(i).Cells(0).Value = tempDist
                DG_RowProperties.Rows(i).Cells(1).Value = 100
                DG_RowProperties.Rows(i).Cells(2).Value = "Ось трассы"
            Next
        End If

        If ChB_CenterBeam.Checked = True Then
            DG_ProletListBeams.Rows(count).HeaderCell.Value = "Ось"
            DG_ProletListBeams.Rows(count).Tag = 0
            count += 1
            DG_RowProperties.Rows(count2).HeaderCell.Value = "Ось"
            DG_RowProperties.Rows(count2).Tag = 0
            DG_RowProperties.Rows(count2).Cells(0).Value = 0
            DG_RowProperties.Rows(count2).Cells(1).Value = 100
            DG_RowProperties.Rows(count2).Cells(2).Value = "Ось трассы"
            count2 += 1
        End If

        If countRowsRightBeam > 0 Then
            For i As Integer = 0 To countRowsRightBeam - 1
                DG_ProletListBeams.Rows(count).HeaderCell.Value = "Ряд П-" & i + 1
                DG_ProletListBeams.Rows(count).Tag = i + 1
                count += 1
            Next i
            Dim deltaLenghtRow As Double = 0.0R
            If countRowsRightBeam > 1 Then
                deltaLenghtRow = (offsetRightLastBeam - offsetRightFirstBeam) / (countRowsRightBeam - 1)
            End If
            For i As Integer = 0 To countRowsRightBeam - 1
                Dim tempDist As Integer = If(countRowsRightBeam = 1,
                                             offsetRightFirstBeam,
                                             offsetRightFirstBeam + deltaLenghtRow * i)
                DG_RowProperties.Rows(count2).HeaderCell.Value = "Ряд П-" & i + 1
                DG_RowProperties.Rows(count2).Tag = i + 1
                DG_RowProperties.Rows(count2).Cells(0).Value = tempDist
                DG_RowProperties.Rows(count2).Cells(1).Value = 100
                DG_RowProperties.Rows(count2).Cells(2).Value = "Ось трассы"
                count2 += 1
            Next
        End If
        Return True
    End Function
End Class
