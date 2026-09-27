Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Topomatic.Acax.Export
Imports Topomatic.Alg.Prf
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime

Public Class BridgeReports
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отчет по точкам опирания
    Public Sub ReportPointPrBeams(ByVal dataBridge As StructureElement)
        If IsNothing(dataBridge) = True Then Exit Sub
        If IsNothing(dataBridge.DWGEntity) = True Then Exit Sub
        If FuncGSON.IsValidJson(dataBridge.KeyParameter) = False Then Exit Sub
        Dim userBridge As Bridges = dataBridge.getBridge
        If Not (TypeOf dataBridge.DWGEntity Is DwgPolyline) Then Exit Sub
        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
        Dim idBridge As String = dataBridge.IdStructure
        Dim nameAlign As String = userBridge.AlignmentName
        Dim nameSurface As String = userBridge.projectSurfaceName
        Dim projectAlign As Topomatic.Alg.Alignment = Nothing
        Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlign)
        Dim projectSurface As Surface = FuncSurface.getSurfaceByName(nameSurface)
        If IsNothing(projectSurface) = True Then
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
        End If
        If IsNothing(projectAlign) = True Then
            MsgBox("Ось трассы не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        'фильтруем все объектв сооружения
        Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(axisLineBridge)
        'фильтруем балки
        Dim axisBeamsDictuionary As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If axisBeamsDictuionary.Count > 0 Then
            'запускаем Excel
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Try
                Dim StartRows As Integer = 0
                Dim startColumns As Integer = 0
                xlApp = New Excel.Application()
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.Sheets.Item(1)
                xlApp.Visible = True
                xlApp.DisplayAlerts = False ' Отключить предупреждения
                Dim cell1 As Excel.Range 'первая ячейка в строке
                Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                Dim CellRng As Excel.Range 'массив ячеек
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'заголовок
                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 1) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 1) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ пролета "
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 2) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 2) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Балки"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 3) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 3) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Марка"
                CellRng.ColumnWidth = 20

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 4) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 4) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "ПК+"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 5) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 5) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Опоры"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 6) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 7) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Смещение"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 6) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Лево"
                CellRng.ColumnWidth = 10

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 7) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Право"
                CellRng.ColumnWidth = 10

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 8) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 10) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "т. 1"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 15) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 17) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "т. 2"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 8) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "X"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 9) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Y"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 10) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Z"
                CellRng.ColumnWidth = 12
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 11) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 11) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "ПК+"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 12) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 12) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Опоры"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 13) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 14) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Смещение"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 13) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Лево"
                CellRng.ColumnWidth = 10
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 14) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Право"
                CellRng.ColumnWidth = 10
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 15) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "X"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 16) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Y"
                CellRng.ColumnWidth = 15
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 17) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Z"
                CellRng.ColumnWidth = 12
                CellRng.NumberFormat = "0.000"

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 18) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 18) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Уклон в промилле"
                CellRng.ColumnWidth = 12
                CellRng.NumberFormat = "0.000"
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 19) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 19) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Примечание"
                CellRng.ColumnWidth = 12
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                StartRows = 4
                For i As Integer = 0 To axisBeamsDictuionary.Count - 1
                    Dim dictBeamsProlet As Dictionary(Of Integer, StructureElement) = axisBeamsDictuionary.ElementAt(i).Value
                    Dim numProlet As Integer = axisBeamsDictuionary.ElementAt(i).Key
                    If dictBeamsProlet.Count > 0 Then
                        For j As Integer = 0 To dictBeamsProlet.Count - 1
                            Dim dataBeam As StructureElement = dictBeamsProlet.ElementAt(j).Value
                            Dim userBeam As BeamI = dataBeam.getBeamI()
                            Dim acLineBeam As DwgLine = dataBeam.DWGEntity
                            If IsNothing(acLineBeam) = False Then
                                Dim PkStart As Double = 0
                                Dim offPKStart As Double = 0
                                Dim PkEnd As Double = 0
                                Dim offPKEnd As Double = 0
                                If IsNothing(projectAlign) = False Then
                                    Dim boolPk1 As Boolean = projectAlign.Plan.CompoundLine.PosToStaOffset(acLineBeam.StartPoint.Pos, PkStart, offPKStart)
                                    Dim boolPk2 As Boolean = projectAlign.Plan.CompoundLine.PosToStaOffset(acLineBeam.EndPoint.Pos, PkEnd, offPKEnd)
                                End If
                                If IsNothing(userBeam) = False Then
                                    Dim numBeam As String = FuncFormatZn.getConditionalRow(userBeam.numberRow)
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 1) 'ссылка на первую ячейку
                                    cell2 = xlSheets.Cells(StartRows + j, startColumns + 19) 'ссылка на 8 ячейку в строке
                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Font.Name = "Times New Roman"
                                    CellRng.Font.Bold = 0
                                    CellRng.Font.Size = 12
                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                    'номер пролета
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 1)
                                    cell1.Value = numProlet
                                    'номер балки
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 2)
                                    cell1.Value = numBeam
                                    'марка
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 3)
                                    cell1.Value = userBeam.model
                                    'пикет
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 4)
                                    cell1.Value = PkStart
                                    cell1.NumberFormat = "0.000"

                                    'номер опоры
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 5)
                                    cell1.Value = numProlet
                                    'смещение
                                    If offPKStart < 0 Then
                                        cell1 = xlSheets.Cells(StartRows + j, startColumns + 6)
                                        cell1.Value = Math.Abs(offPKStart)
                                        cell1.NumberFormat = "0.000"
                                    Else
                                        cell1 = xlSheets.Cells(StartRows + j, startColumns + 7)
                                        cell1.Value = Math.Abs(offPKStart)
                                        cell1.NumberFormat = "0.000"
                                    End If
                                    'координаты
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 8)
                                    cell1.Value = acLineBeam.StartPoint.Y
                                    cell1.NumberFormat = "0.000"
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 9)
                                    cell1.Value = acLineBeam.StartPoint.X
                                    cell1.NumberFormat = "0.000"
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 10)
                                    cell1.Value = acLineBeam.StartPoint.Z
                                    cell1.NumberFormat = "0.000"
                                    '2 точка
                                    'пикет
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 11)
                                    cell1.Value = PkEnd
                                    cell1.NumberFormat = "0.000"
                                    'номер балки
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 12)
                                    cell1.Value = numProlet + 1
                                    'смещение
                                    If offPKEnd < 0 Then
                                        cell1 = xlSheets.Cells(StartRows + j, startColumns + 13)
                                        cell1.Value = Math.Abs(offPKEnd)
                                        cell1.NumberFormat = "0.000"
                                    Else
                                        cell1 = xlSheets.Cells(StartRows + j, startColumns + 14)
                                        cell1.Value = Math.Abs(offPKEnd)
                                        cell1.NumberFormat = "0.000"
                                    End If
                                    'координаты
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 15)
                                    cell1.Value = acLineBeam.EndPoint.Y
                                    cell1.NumberFormat = "0.000"
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 16)
                                    cell1.Value = acLineBeam.EndPoint.X
                                    cell1.NumberFormat = "0.000"
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 17)
                                    cell1.Value = acLineBeam.EndPoint.Z
                                    cell1.NumberFormat = "0.000"
                                    'уклон
                                    Dim h As Double = acLineBeam.StartPoint.Z - acLineBeam.EndPoint.Z
                                    Dim a As Double = MathFunction.funcCalcDistanceByToPoints2d(acLineBeam.StartPoint.Pos, acLineBeam.EndPoint.Pos)
                                    Dim iOBJ As Double = (h / a) * 1000
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 18)
                                    cell1.Value = iOBJ
                                    cell1.NumberFormat = "0.000"
                                    'примечание
                                    cell1 = xlSheets.Cells(StartRows + j, startColumns + 19)
                                    cell1.Value = "-"
                                End If
                            End If
                        Next j
                        StartRows = StartRows + dictBeamsProlet.Count
                    End If
                Next
            Catch ex As System.Exception
            Finally
                If xlBook IsNot Nothing Then
                    Marshal.ReleaseComObject(xlBook)
                End If

                If xlApp IsNot Nothing Then
                    Marshal.ReleaseComObject(xlApp)
                End If

                ' Очистка COM-объектов
                If xlSheets IsNot Nothing Then Marshal.ReleaseComObject(xlSheets)

                ' Сборка мусора для гарантии освобождения ресурсов
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End If
    End Sub
    '===========================================================================================================================
    'Ведомость верха плиты балок и толщины покрытия
    Public Sub ReportUpPokr(ByVal dataBridge As StructureElement)
        If IsNothing(dataBridge) = True Then Exit Sub
        If IsNothing(dataBridge.DWGEntity) = True Then Exit Sub
        If FuncGSON.IsValidJson(dataBridge.KeyParameter) = False Then Exit Sub
        Dim userBridge As Bridges = dataBridge.getBridge
        If Not (TypeOf dataBridge.DWGEntity Is DwgPolyline) Then Exit Sub
        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
        Dim idBridge As String = dataBridge.IdStructure
        Dim nameAlign As String = userBridge.AlignmentName
        Dim nameSurface As String = userBridge.projectSurfaceName
        Dim projectAlign As Topomatic.Alg.Alignment = Nothing
        Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlign)
        Dim projectSurface As Surface = FuncSurface.getSurfaceByName(nameSurface)
        If IsNothing(projectSurface) = True Then
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
        End If
        If IsNothing(projectAlign) = True Then
            MsgBox("Ось трассы не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        'фильтруем все объектв сооружения
        Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(axisLineBridge)
        'фильтруем балки
        Dim axisBeamsDictuionary As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If axisBeamsDictuionary.Count > 0 Then
            'запускаем Excel
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Try
                Dim StartRows As Integer = 0
                Dim startColumns As Integer = 0
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.Sheets.Item(1)
                xlApp.Visible = True
                Dim cell1 As Excel.Range 'первая ячейка в строке
                Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                Dim CellRng As Excel.Range 'массив ячеек

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 1) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 4, startColumns + 1) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ пролета "
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 2) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 4, startColumns + 2) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Опоры"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 3) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 4, startColumns + 3) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Балки"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 4) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 9) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Координаты крайних точек в зоне опирания балки"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 4) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 6) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Лево"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 7) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 9) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Право"

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 4) 'ссылка на первую ячейку
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "X"
                cell1.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 5) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Y"
                cell1.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 6) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Z"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 7) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "X"
                cell1.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 8) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Y"
                cell1.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 9) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Z"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 10) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 11) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Отметки покрытия, м"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 10) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Лево"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 11) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Право"

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 10) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "H"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 11) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "H"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 12) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 13) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Толщина покрытия, м"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 12) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Лево"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 13) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Право"

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 12) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "dh"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 4, startColumns + 13) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "dh"
                cell1.ColumnWidth = 12
                StartRows = 5
                If axisBeamsDictuionary.Count > 0 Then
                    For i As Integer = 0 To axisBeamsDictuionary.Count - 1
                        Dim dictBeamsProlet As Dictionary(Of Integer, StructureElement) = axisBeamsDictuionary.ElementAt(i).Value
                        Dim numProlet As Integer = axisBeamsDictuionary.ElementAt(i).Key
                        Dim tempArrayZn As String(,) = {}
                        Dim countTempArrayZn As Integer = 0
                        If dictBeamsProlet.Count > 0 Then
                            For j As Integer = 0 To dictBeamsProlet.Count - 1
                                Dim dataBeam As StructureElement = dictBeamsProlet.ElementAt(j).Value
                                Dim userBeam As BeamI = dataBeam.getBeamI()
                                Dim acLineBeam As DwgLine = dataBeam.DWGEntity
                                If IsNothing(acLineBeam) = False Then
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 1) 'ссылка на первую ячейку
                                    cell2 = xlSheets.Cells(StartRows, startColumns + 13) 'ссылка на 13 ячейку в строке
                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Font.Name = "Times New Roman"
                                    CellRng.Font.Bold = 0
                                    CellRng.Font.Size = 12
                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                    'номер пролета балки на первой опоре
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 1)
                                    cell1.Value = userBeam.numberProlet
                                    'сохраняем пролет
                                    ReDim Preserve tempArrayZn(12, countTempArrayZn)
                                    tempArrayZn(0, countTempArrayZn) = userBeam.numberProlet
                                    'номер опоры в начале балки
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 2)
                                    cell1.Value = userBeam.numberProlet
                                    'записываем номер второй опоры
                                    tempArrayZn(1, countTempArrayZn) = userBeam.numberProlet + 1
                                    'номер балки
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 3)
                                    Dim numBeam As String = FuncFormatZn.getConditionalRow(userBeam.numberRow)
                                    cell1.Value = numBeam
                                    tempArrayZn(2, countTempArrayZn) = numBeam
                                    'ищем отметки верха
                                    Dim listStartTopPoint As List(Of Vector3D) = New List(Of Vector3D)
                                    Dim listEndTopPoint As List(Of Vector3D) = New List(Of Vector3D)
                                    'удлинняем балку на величину участков опирания и обрезаем на величину омоноличивания
                                    Dim VBeam As Boolean = CalculationBeams.restoreElementsBeam(acLineBeam, userBeam, listStartTopPoint, listEndTopPoint, CalculationBeams.rectoreBeam.siteMonolit)
                                    'точка1 начало балки лево
                                    If listStartTopPoint.Count = 4 Then
                                        cell1 = xlSheets.Cells(StartRows, startColumns + 4)
                                        cell1.Value = listStartTopPoint.Item(0).Y
                                        cell1.NumberFormat = "0.000"

                                        cell1 = xlSheets.Cells(StartRows, startColumns + 5)
                                        cell1.Value = listStartTopPoint.Item(0).X
                                        cell1.NumberFormat = "0.000"

                                        cell1 = xlSheets.Cells(StartRows, startColumns + 6)
                                        cell1.Value = listStartTopPoint.Item(0).Z
                                        cell1.NumberFormat = "0.000"

                                        'точка2 начало балки право
                                        cell1 = xlSheets.Cells(StartRows, startColumns + 7)
                                        cell1.Value = listStartTopPoint.Item(1).Y
                                        cell1.NumberFormat = "0.000"

                                        cell1 = xlSheets.Cells(StartRows, startColumns + 8)
                                        cell1.Value = listStartTopPoint.Item(1).X
                                        cell1.NumberFormat = "0.000"

                                        cell1 = xlSheets.Cells(StartRows, startColumns + 9)
                                        cell1.Value = listStartTopPoint.Item(1).Z
                                        cell1.NumberFormat = "0.000"
                                    End If
                                    'абсолютные отметки начало балки лево
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 10)
                                    Try
                                        Dim h As Double = projectSurface.GetElevation(listStartTopPoint.Item(0).Pos)
                                        cell1.Value = h
                                        cell1.NumberFormat = "0.000"
                                        Dim dh As Double = h - listStartTopPoint.Item(1).Z
                                        cell1 = xlSheets.Cells(StartRows, startColumns + 12)
                                        cell1.Value = dh
                                        cell1.NumberFormat = "0.000"
                                    Catch ex As System.Exception
                                    End Try

                                    'абсолютные отметки начало балки право
                                    cell1 = xlSheets.Cells(StartRows, startColumns + 11)
                                    Try
                                        Dim h As Double = projectSurface.GetElevation(listStartTopPoint.Item(1).Pos)
                                        cell1.Value = h
                                        cell1.NumberFormat = "0.000"
                                        Dim dh As Double = h - listStartTopPoint.Item(0).Z
                                        cell1 = xlSheets.Cells(StartRows, startColumns + 13)
                                        cell1.Value = dh
                                        cell1.NumberFormat = "0.000"
                                    Catch ex As System.Exception
                                    End Try

                                    'третья точка
                                    tempArrayZn(3, countTempArrayZn) = listEndTopPoint.Item(0).Y
                                    tempArrayZn(4, countTempArrayZn) = listEndTopPoint.Item(0).X
                                    tempArrayZn(5, countTempArrayZn) = listEndTopPoint.Item(0).Z
                                    Try
                                        Dim h As Double = projectSurface.GetElevation(listEndTopPoint.Item(0).Pos)
                                        tempArrayZn(9, countTempArrayZn) = h
                                        Dim dh As Double = h - listEndTopPoint.Item(0).Z
                                        tempArrayZn(11, countTempArrayZn) = dh
                                    Catch ex As System.Exception
                                    End Try

                                    'точка 4
                                    tempArrayZn(6, countTempArrayZn) = listEndTopPoint.Item(1).Y
                                    tempArrayZn(7, countTempArrayZn) = listEndTopPoint.Item(1).X
                                    tempArrayZn(8, countTempArrayZn) = listEndTopPoint.Item(1).Z
                                    Try
                                        Dim h As Double = projectSurface.GetElevation(listEndTopPoint.Item(1).Pos)
                                        tempArrayZn(10, countTempArrayZn) = h
                                        Dim dh As Double = h - listEndTopPoint.Item(1).Z
                                        tempArrayZn(12, countTempArrayZn) = dh
                                    Catch ex As System.Exception
                                    End Try
                                    countTempArrayZn += 1
                                    StartRows = StartRows + 1
                                End If
                            Next j
                        End If
                        If IsArray(tempArrayZn) = True Then
                            For j As Integer = 0 To tempArrayZn.GetUpperBound(1)
                                cell1 = xlSheets.Cells(StartRows, startColumns + 1) 'ссылка на первую ячейку
                                cell2 = xlSheets.Cells(StartRows, startColumns + 13) 'ссылка на 13 ячейку в строке
                                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                CellRng.Font.Name = "Times New Roman"
                                CellRng.Font.Bold = 0
                                CellRng.Font.Size = 12
                                CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                'номер пролета балки на первой опоре
                                cell1 = xlSheets.Cells(StartRows, startColumns + 1)
                                cell1.Value = tempArrayZn(0, j)

                                cell1 = xlSheets.Cells(StartRows, startColumns + 2)
                                cell1.Value = tempArrayZn(1, j)

                                cell1 = xlSheets.Cells(StartRows, startColumns + 3)
                                cell1.Value = tempArrayZn(2, j)

                                cell1 = xlSheets.Cells(StartRows, startColumns + 4)
                                cell1.Value = tempArrayZn(3, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 5)
                                cell1.Value = tempArrayZn(4, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 6)
                                cell1.Value = tempArrayZn(5, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 7)
                                cell1.Value = tempArrayZn(6, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 8)
                                cell1.Value = tempArrayZn(7, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 9)
                                cell1.Value = tempArrayZn(8, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 10)
                                cell1.Value = tempArrayZn(9, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 11)
                                cell1.Value = tempArrayZn(10, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 12)
                                cell1.Value = tempArrayZn(11, j)
                                cell1.NumberFormat = "0.000"

                                cell1 = xlSheets.Cells(StartRows, startColumns + 13)
                                cell1.Value = tempArrayZn(12, j)
                                cell1.NumberFormat = "0.000"

                                StartRows += 1
                            Next j
                        End If
                    Next i
                End If
            Catch ex As System.Exception
            Finally
                If xlBook IsNot Nothing Then
                    Marshal.ReleaseComObject(xlBook)
                End If

                If xlApp IsNot Nothing Then
                    Marshal.ReleaseComObject(xlApp)
                End If

                ' Очистка COM-объектов
                If xlSheets IsNot Nothing Then Marshal.ReleaseComObject(xlSheets)

                ' Сборка мусора для гарантии освобождения ресурсов
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End If
    End Sub
    '==========================================================================================================================
    'Ведомость по осям опор
    Public Sub ReportAxisPillars(ByVal dataBridge As StructureElement)
        If IsNothing(dataBridge) = True Then Exit Sub
        If IsNothing(dataBridge.DWGEntity) = True Then Exit Sub
        If FuncGSON.IsValidJson(dataBridge.KeyParameter) = False Then Exit Sub
        Dim userBridge As Bridges = dataBridge.getBridge
        If Not (TypeOf dataBridge.DWGEntity Is DwgPolyline) Then Exit Sub
        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
        Dim idBridge As String = dataBridge.IdStructure
        Dim nameAlign As String = userBridge.AlignmentName
        Dim nameSurface As String = userBridge.projectSurfaceName
        Dim projectAlign As Topomatic.Alg.Alignment = Nothing
        Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlign)
        Dim projectSurface As Surface = FuncSurface.getSurfaceByName(nameSurface)
        If IsNothing(projectSurface) = True Then
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
        End If
        If IsNothing(projectAlign) = True Then
            MsgBox("Ось трассы не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        Dim alignProjectPolyline3d As Polyline3D = New Polyline3D()
        projectAlign.Plan.CompoundLine.ToPolyLine(alignProjectPolyline3d)
        'фильтруем все объектв сооружения
        Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(axisLineBridge)
        'фильтруем балки
        Dim axisPillarsDictionary As Dictionary(Of Integer, StructureElement) = userBridge.getAxisPillars(dictionaryBridgeElements)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If axisPillarsDictionary.Count > 0 Then
            'запускаем Excel
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Dim StartRows As Integer = 0
            Dim startColumns As Integer = 0
            Try
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.Sheets.Item(1)
                xlApp.Visible = True
                Dim cell1 As Excel.Range 'первая ячейка в строке
                Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                Dim CellRng As Excel.Range 'массив ячеек
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'формируем отчет
                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 1) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 1) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ опоры"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 2) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 2) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Пикет+"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 3) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 3) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Угол пересечения, град"
                CellRng.ColumnWidth = 20
                CellRng.NumberFormat = "0.000000"
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 4) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 4) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Дирекционный угол, град"
                CellRng.ColumnWidth = 20
                CellRng.NumberFormat = "0.000000"
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 5) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 7) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Координаты пересечения, м"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 5) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "X"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 6) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Y"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 7) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "H"
                CellRng.ColumnWidth = 15

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 8) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 8) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Отметка земли, м"
                CellRng.ColumnWidth = 15
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 9) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 9) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Примечание"

                StartRows = 4
                Dim count As Integer = 0
                For i As Integer = 0 To axisPillarsDictionary.Count - 1
                    Dim numberPillars As Integer = axisPillarsDictionary.ElementAt(i).Key
                    Dim dataPillar As StructureElement = axisPillarsDictionary.ElementAt(i).Value
                    If IsNothing(dataPillar) = False Then
                        Dim userPillar As Pillar = dataPillar.getPillar
                        Dim acLinePillar As DwgLine = dataPillar.DWGEntity
                        If IsNothing(acLinePillar) = False Then
                            If IsNothing(userPillar) = False Then
                                '=====================================================================================
                                Dim PKStart As Double = -1
                                Dim PKOffset As Double = -1
                                Dim dirAngle As Double = 0
                                Dim angleRezult As Double = 0
                                Dim HEarth As Double = 0

                                Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(alignProjectPolyline3d, acLinePillar.StartPoint.Pos, acLinePillar.EndPoint.Pos)
                                If pointIntersectCollection.Count = 1 Then
                                    'начальная точка раскладки балок
                                    Dim startAlignPoint As Vector2D = pointIntersectCollection.ElementAt(0)
                                    Dim elevST As Double = 0
                                    Try
                                        elevST = projectSurface.GetElevation(startAlignPoint)
                                    Catch ex As Exception
                                        Continue For
                                    End Try
                                    Dim boolPK As Boolean = projectAlign.Plan.CompoundLine.PosToStaOffset(startAlignPoint, PKStart, PKOffset)
                                    Dim angle As Double = 0
                                    If boolPK = True Then
                                        Dim pos1 As Vector2D = Nothing
                                        Dim pos2 As Vector2D = Nothing
                                        Dim boolAngle As Boolean = projectAlign.Plan.CompoundLine.GetTangentPosition(PKStart, pos1, pos2)
                                        For Each acTransition As Transition In projectAlign.Transitions
                                            Dim userBlackProfile As Profile = acTransition.EgProfile
                                            If userBlackProfile.MinStation <= PKStart And userBlackProfile.MaxStation >= PKStart Then
                                                userBlackProfile.GetY(PKStart, HEarth)
                                                Exit For
                                            End If
                                        Next
                                        If boolAngle = True Then
                                            Dim anglePointAlign As Double = (pos1 - pos2).Angle
                                            Dim anglePointAxis As Double = (acLinePillar.EndPoint.Pos - acLinePillar.StartPoint.Pos).Angle
                                            angleRezult = anglePointAxis - anglePointAlign
                                            angleRezult = (angleRezult * 180) / Math.PI
                                        End If
                                        dirAngle = MathFunction.funcCalcDirectionAngleByToPoints2d(acLinePillar.StartPoint.Pos, acLinePillar.EndPoint.Pos, 8)
                                    End If
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 1) 'ссылка на первую ячейку
                                    cell2 = xlSheets.Cells(StartRows + count, startColumns + 9) 'ссылка на 7 ячейку в строке
                                    CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                    CellRng.Font.Name = "Times New Roman"
                                    CellRng.Font.Bold = 0
                                    CellRng.Font.Size = 12
                                    CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                    CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                    'номер опоры
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 1)
                                    cell1.Value = userPillar.Number
                                    'пикет
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 2)
                                    cell1.Value = PKStart
                                    cell1.NumberFormat = "0.000"

                                    'угол
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 3)
                                    cell1.Value = angleRezult
                                    cell1.NumberFormat = "0.000000"
                                    'дирекционный угол
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 4)
                                    cell1.Value = dirAngle
                                    cell1.NumberFormat = "0.000000"
                                    'X
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 5)
                                    cell1.Value = startAlignPoint.Y
                                    cell1.NumberFormat = "0.000"
                                    'Y
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 6)
                                    cell1.Value = startAlignPoint.X
                                    cell1.NumberFormat = "0.000"
                                    'H
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 7)
                                    cell1.Value = elevST
                                    cell1.NumberFormat = "0.000"
                                    'отметка земли
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 8)
                                    cell1.Value = HEarth
                                    cell1.NumberFormat = "0.000"
                                    'Примечание
                                    cell1 = xlSheets.Cells(StartRows + count, startColumns + 9)
                                    cell1.Value = "-"
                                End If

                            End If
                        End If
                    End If
                    count += 1
                Next i
            Catch ex As Exception
            Finally
                If xlBook IsNot Nothing Then
                    Marshal.ReleaseComObject(xlBook)
                End If

                If xlApp IsNot Nothing Then
                    Marshal.ReleaseComObject(xlApp)
                End If

                ' Очистка COM-объектов
                If xlSheets IsNot Nothing Then Marshal.ReleaseComObject(xlSheets)

                ' Сборка мусора для гарантии освобождения ресурсов
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End If
    End Sub
    '===========================================================================================================================
    'Отчет по деформационным зазорам
    Public Sub ReportDefZazor(ByVal dataBridge As StructureElement)
        If IsNothing(dataBridge) = True Then Exit Sub
        If IsNothing(dataBridge.DWGEntity) = True Then Exit Sub
        If FuncGSON.IsValidJson(dataBridge.KeyParameter) = False Then Exit Sub
        Dim userBridge As Bridges = dataBridge.getBridge
        If Not (TypeOf dataBridge.DWGEntity Is DwgPolyline) Then Exit Sub
        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
        Dim idBridge As String = dataBridge.IdStructure
        Dim nameAlign As String = userBridge.AlignmentName
        Dim nameSurface As String = userBridge.projectSurfaceName
        Dim projectAlign As Topomatic.Alg.Alignment = Nothing
        Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlign)
        Dim projectSurface As Surface = FuncSurface.getSurfaceByName(nameSurface)
        If IsNothing(projectSurface) = True Then
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
        End If
        If IsNothing(projectAlign) = True Then
            MsgBox("Ось трассы не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена!!! Отчет не сформирован.")
            Exit Sub
        End If
        'фильтруем все объектв сооружения
        Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(axisLineBridge)
        If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisBeam) = False Then
            Exit Sub
        End If
        Dim listBeams As List(Of StructureElement) = dictionaryBridgeElements.Item(StructureElement.typeObject.axisBeam)
        'фильтруем балки
        Dim axisBeamsDictuionary As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If axisBeamsDictuionary.Count > 0 Then
            'запускаем Excel
            Dim xlApp As Excel.Application = Nothing
            Dim xlBook As Excel.Workbook = Nothing
            Dim xlSheets As Excel.Worksheet = Nothing
            Dim StartRows As Integer = 0
            Dim startColumns As Integer = 0
            Try
                xlApp = CreateObject("Excel.Application")
                xlBook = xlApp.Workbooks.Add()
                xlSheets = xlBook.Sheets.Item(1)
                xlApp.Visible = True
                Dim cell1 As Excel.Range 'первая ячейка в строке
                Dim cell2 As Excel.Range 'последняя ячейка в строке(6)
                Dim CellRng As Excel.Range 'массив ячеек

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 1) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 1) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ пролета "
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 2) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 2) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ Балки"
                CellRng.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 3) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 3) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Длина балки, м"
                CellRng.ColumnWidth = 12
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 4) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 4) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Высота балки, м"
                CellRng.ColumnWidth = 12
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 5) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 5) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Ширина балки по низу, м"
                CellRng.ColumnWidth = 17
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 6) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 6) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Ширина балки по верху, м"
                CellRng.ColumnWidth = 17
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 7) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 7) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Отметка балки в точке опирания А, м"
                CellRng.ColumnWidth = 22
                CellRng.WrapText = True

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 8) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 8) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Отметка балки в точке опирания Б, м"
                CellRng.ColumnWidth = 22
                CellRng.WrapText = True
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 9) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 2, startColumns + 14) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Зазоры, м"

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 9) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ пролета"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 10) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "№ балки"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 11) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Верх лево"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 12) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Верх право"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 13) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Низ лево"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 3, startColumns + 14) 'ссылка на первую ячейку
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                cell1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Низ право"
                cell1.ColumnWidth = 12

                cell1 = xlSheets.Cells(StartRows + 2, startColumns + 15) 'ссылка на первую ячейку
                cell2 = xlSheets.Cells(StartRows + 3, startColumns + 15) 'ссылка на 8 ячейку в строке
                CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                CellRng.Merge(Type.Missing)
                cell1.Font.Name = "Times New Roman"
                cell1.Font.Bold = 1
                cell1.Font.Size = 12
                cell1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                cell1.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                cell1.Value = "Примечание"
                CellRng.ColumnWidth = 15

                StartRows = 4
                For i As Integer = 0 To axisBeamsDictuionary.Count - 1
                    Dim dictBeamsProlet As Dictionary(Of Integer, StructureElement) = axisBeamsDictuionary.ElementAt(i).Value
                    Dim numProlet As Integer = axisBeamsDictuionary.ElementAt(i).Key
                    If dictBeamsProlet.Count > 0 Then
                        For j As Integer = 0 To dictBeamsProlet.Count - 1
                            Dim dataBeam As StructureElement = dictBeamsProlet.ElementAt(j).Value
                            Dim userBeam As BeamI = dataBeam.getBeamI()
                            Dim acLineBeam As DwgLine = dataBeam.DWGEntity
                            If IsNothing(acLineBeam) = False Then
                                If IsNothing(userBeam) = False Then
                                    Dim numbProlet As Integer = userBeam.numberProlet
                                    Dim dataPrevBeam As StructureElement = CalculationBeams.getPreviousBeam(acLineBeam, listBeams)
                                    If IsNothing(dataPrevBeam) = True Then Continue For
                                    Dim acPrevLine As DwgLine = dataPrevBeam.DWGEntity
                                    Dim userPrevBeam As BeamI = dataPrevBeam.getBeamI
                                    If numbProlet > 1 Then
                                        'считаем зазор
                                        Dim listZazor As List(Of Double) = New List(Of Double)
                                        Dim zazor As Double = userBridge.zazorPreviousBeam(acPrevLine, acLineBeam, listZazor, userBeam)
                                        If listZazor.Count > 0 Then
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 1) 'ссылка на первую ячейку
                                            cell2 = xlSheets.Cells(StartRows, startColumns + 15) 'ссылка на 8 ячейку в строке
                                            CellRng = xlSheets.Range(cell1, cell2) 'выбираем строку
                                            CellRng.Font.Name = "Times New Roman"
                                            CellRng.Font.Bold = 0
                                            CellRng.Font.Size = 12
                                            CellRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                                            CellRng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                            CellRng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                            'номер пролета
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 1)
                                            cell1.Value = numbProlet - 1
                                            'номер балки
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 2)
                                            Dim numberPrevBeam As String = FuncFormatZn.getConditionalRow(userPrevBeam.numberRow)
                                            Dim numberBeam As String = FuncFormatZn.getConditionalRow(userBeam.numberRow)
                                            cell1.Value = numberPrevBeam
                                            'длина балки
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 3)
                                            cell1.Value = acPrevLine.Length
                                            cell1.NumberFormat = "0.000"
                                            'высота балки
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 4)
                                            cell1.Value = userPrevBeam.height
                                            cell1.NumberFormat = "0.000"
                                            'ширина по низу
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 5)
                                            cell1.Value = userPrevBeam.widthBottom
                                            cell1.NumberFormat = "0.000"
                                            'ширина по верху
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 6)
                                            cell1.Value = userPrevBeam.WidthTop
                                            cell1.NumberFormat = "0.000"
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'высота
                                            If IsNothing(projectSurface) = False Then
                                                Try
                                                    Dim H1 As Double = projectSurface.GetElevation(acPrevLine.StartPoint.Pos)
                                                    Dim H2 As Double = projectSurface.GetElevation(acPrevLine.EndPoint.Pos)
                                                    'ширина по низу
                                                    cell1 = xlSheets.Cells(StartRows, startColumns + 7)
                                                    cell1.Value = H1 - userPrevBeam.height
                                                    cell1.NumberFormat = "0.000"
                                                    'ширина по верху
                                                    cell1 = xlSheets.Cells(StartRows, startColumns + 8)
                                                    cell1.Value = H2 - userPrevBeam.height
                                                    cell1.NumberFormat = "0.000"
                                                Catch ex As System.Exception
                                                End Try
                                            End If
                                            'номер пролета балки 2
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 9)
                                            cell1.Value = numbProlet
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 10)
                                            cell1.Value = numberBeam
                                            'зазоры
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 11)
                                            cell1.Value = listZazor.Item(0)
                                            cell1.NumberFormat = "0.000"
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 12)
                                            cell1.Value = listZazor.Item(1)
                                            cell1.NumberFormat = "0.000"
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 13)
                                            cell1.Value = listZazor.Item(2)
                                            cell1.NumberFormat = "0.000"
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 14)
                                            cell1.Value = listZazor.Item(3)
                                            cell1.NumberFormat = "0.000"
                                            'примечание
                                            cell1 = xlSheets.Cells(StartRows, startColumns + 15)
                                            cell1.Value = "-"
                                            StartRows += 1
                                        End If
                                    End If
                                Else
                                    'шкафная стенка
                                End If
                            End If
                        Next j
                    End If
                Next i
            Catch ex As Exception
            Finally
                If xlBook IsNot Nothing Then
                    Marshal.ReleaseComObject(xlBook)
                End If

                If xlApp IsNot Nothing Then
                    Marshal.ReleaseComObject(xlApp)
                End If

                ' Очистка COM-объектов
                If xlSheets IsNot Nothing Then Marshal.ReleaseComObject(xlSheets)

                ' Сборка мусора для гарантии освобождения ресурсов
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End If
    End Sub
    '=========================================================================================================================
    'экспорт в автокад
    Public Sub ExportToAutucad(ByVal dataBridge As StructureElement)
        If IsNothing(dataBridge) = True Then Exit Sub
        If IsNothing(dataBridge.DWGEntity) = True Then Exit Sub
        If FuncGSON.IsValidJson(dataBridge.KeyParameter) = False Then Exit Sub
        Dim userBridge As Bridges = dataBridge.getBridge
        If Not (TypeOf dataBridge.DWGEntity Is DwgPolyline) Then Exit Sub
        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
        Dim idBridge As String = dataBridge.IdStructure
        'фильтруем все объектв сооружения
        Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(axisLineBridge)
        Dim rng As Random = New Random
        If dictionaryBridgeElements.Count > 0 Then
            Dim strFilter As String = "AutoCAD files (*.dwg)|*.dwg|All files (*.*)|*.*"
            Dim strFileName As String = FuncFiles.FuncSaveFilePathName("Экспорт данных", strFilter)
            If strFileName.Trim.Length > 0 Then
                Dim drawing As Topomatic.Dwg.Drawing = New Topomatic.Dwg.Drawing()
                Dim blockEnt As DwgBlock = drawing.ActiveSpace
                For i As Integer = 0 To dictionaryBridgeElements.Count - 1
                    Dim listBridgeElements As List(Of StructureElement) = dictionaryBridgeElements.ElementAt(i).Value
                    If IsNothing(listBridgeElements) = False Then
                        If listBridgeElements.Count > 0 Then
                            For j As Integer = 0 To listBridgeElements.Count - 1
                                Dim dataElement As StructureElement = listBridgeElements.Item(j)
                                If IsNothing(dataElement) = True Then Continue For
                                Dim entity As DwgEntity = dataElement.DWGEntity
                                If IsNothing(entity) = True Then Continue For
                                If TypeOf entity Is DwgModel3DElement Then
                                    Dim AcObject As DwgModel3DElement = entity
                                    Dim element As ImElement = AcObject.Element
                                    If IsNothing(element) = False Then
                                        Dim viewElement As ImViewElement = element
                                        Dim blockName = $"{"TLC_Obj"}_{rng.Next}_prf"
                                        Dim elementBlock As DwgBlock = drawing.Blocks.Add(blockName)
                                        Dim Res As Boolean = viewElement.GetView(Model3DView.Top, elementBlock, Matrix.Identity, 1.0)
                                        Dim generalPropertiesObject1 As ImProperties = element.GetProperties()
                                        If generalPropertiesObject1.Count > 0 Then
                                            For k As Integer = 0 To generalPropertiesObject1.Count - 1
                                                Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k)
                                                Dim g As Topomatic.Visualization.Geometry.GeometryModel3D = AcObject.Element.GetModel
                                                Dim l As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                                                Dim de As ObjectsDisjointerArgs = New ObjectsDisjointerArgs(AcObject.Bounds)
                                                AcObject.GetNodePoint(de, l)
                                            Next k
                                        End If
                                        Dim m As Matrix = AcObject.Matrix
                                        If Res = True Then
                                            blockEnt.AddInsert(
                                        New Vector3D(m.Translation.Pos, 0),
                                        Vector3D.One, 0.0, blockName)
                                        End If
                                    End If
                                Else
                                    Dim destAcObject As DwgEntity = entity.Clone
                                    blockEnt.Add(destAcObject)
                                End If
                            Next j
                        End If
                    End If
                Next i
                Dim provider As DrawingExportProvider = DrawingExportProvider.GetProvider("DWG")
                provider.SaveToFile(strFileName, drawing)
                Exit Sub
            End If
        End If
    End Sub
End Class
