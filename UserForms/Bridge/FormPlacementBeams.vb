Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Controls.Common
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Sfc
Imports Topomatic.Sites.Core
Imports Topomatic.Visualization.Runtime
Public Class FormPlacementBeams
    Public boolShowDlg As Boolean = False
    Public ActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public arrangementProject As ArrangementModel = Nothing
    Public civilStructuresProject As ProjectCivilStructures = Nothing
    Public civilBridgeProject As ProjectBridge = Nothing

    Public dictNamesProjectBridge As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public dictionaryBridge As Dictionary(Of String, StructureElement) = Nothing 'словарь с мостами активного чертежа
    Public dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = New Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) 'словарь с элементами чертежа
    Public dictionaryBridgePillars As Dictionary(Of Integer, List(Of StructureElement)) = New Dictionary(Of Integer, List(Of StructureElement)) 'словарь для осей опор и осей опирания балок
    Public dictionaryBridgeBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) 'словарь для  балок
    Public dictionaryBridgeTraectoryBeams As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement) 'стоварь для траектории раскладки балок

    Public listArragmentStructures As List(Of ProjectBridge) = New List(Of ProjectBridge)
    Public metodPlacmentBeams As Bridges.typePlacementBeam = Bridges.typePlacementBeam.fixed '(0-обычная раскладка, 1-раскладка балками индивидуального проектирования, 2-раскладка балками инливидуального проектирования с вычислением минимального зазора)
    Public objectBridgeDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,)) 'словарь со всеми элементами мостового сооружения
    Public brigeGeneralAxisDictionary As Dictionary(Of String, String()) = New Dictionary(Of String, String())
    Public dictionaryAlbumBeams As Dictionary(Of String, BeamI()) = New Dictionary(Of String, BeamI()) 'словарь с балками (имя альбома,балки)
    Public axisPillarsDictionary As Dictionary(Of Integer, DwgLine()) = New Dictionary(Of Integer, DwgLine()) 'словарь с уже существующими осями опор (номер опоры, 0-ось опирания1,1-ось опоры,2-ось опирания2
    Public dictProjectModel As Dictionary(Of IProjectModel, String) = New Dictionary(Of IProjectModel, String) 'своварь с проектами

    Public axisBeamsDictionary As Dictionary(Of String, DwgLine) = New Dictionary(Of String, DwgLine) 'словарь с существующими балками 
    Public elementsBeamDictionary As Dictionary(Of String, DwgEntity()) = New Dictionary(Of String, DwgEntity()) 'словарь со вспомогательными элементами балки
    Public elementsBridgeDictionary As Dictionary(Of String, DwgEntity) = New Dictionary(Of String, DwgEntity) 'словарь со вспомогательными элементами путепровода

    Public dictionaryProlet As Dictionary(Of Integer, String(,)) = New Dictionary(Of Integer, String(,)) 'словарь с пролетами, номер пролета, массив 0-номер ряда,1-альбом, 2-балка, 3-длина балки
    Public dictiondryPillars As Dictionary(Of Integer, Double()) = New Dictionary(Of Integer, Double()) 'словарь для характеристик опор, номер опоры, массив 0-левый зазор, 1-правый зазор
    Public dictiondryRows As Dictionary(Of Integer, String()) = New Dictionary(Of Integer, String()) 'словарь с рядами' номер ряда, массив 0-смещение от оси,1-смещение от поверхности,2-id трассы или полилинии
    Public arrayNameAlbom As String() = {""} 'массив с альбомаи балок
    Public arrayNameBridges As String() = {""} 'массив с именами сооружений
    Public dirTemplate As String() = Nothing
    Public idBridge As String = ""
    Public generalDir As String = ""
    Public axisLineBridge As DwgPolyline = Nothing
    Public formUserTempl As FormUserTemptate = New FormUserTemptate
    Public boolWriteData As Boolean = True
    Public boolButtonRows As Boolean = False
    Public boolButtonSelectPillar As Boolean = False
    Public numberSelectRows As Integer = -1
    Private lastInvalidPKText As String = Nothing

    Private Function FuncCreateTables() As Boolean
        FuncCreateTables = False
        Dim numberProlet As Integer = NUpD_CountProlet.Value

        Dim countRowsLeftBeam As Integer = NUpD_CountLeftRows.Value 'количество
        Dim offsetLeftFirstBeam As Integer = NumericUpDown3.Value  'смещение первой оси влево
        Dim offsetLeftLastBeam As Integer = NumericUpDown5.Value 'смещение последней оси влево

        Dim countRowsRightBeam As Double = NUpD_CountRightRows.Value 'количество
        Dim offsetRightFirstBeam As Integer = NumericUpDown8.Value  'смещение первой оси
        Dim offsetRightLastBeam As Integer = NumericUpDown9.Value  'смещение последней оси
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        DG_ProletListBeams.Rows.Add(countRows)
        DG_RowProperties.Rows.Add(countRows)
        Dim count As Integer = 0
        Dim count2 As Integer = 0
        If countRowsLeftBeam > 0 Then
            For i As Integer = countRowsLeftBeam - 1 To 0 Step -1
                DG_ProletListBeams.Rows(count).HeaderCell.Value = "Ряд Л-" & i + 1
                DG_ProletListBeams.Rows(count).Tag = -1 * (i + 1)
                count += 1
            Next i

            Dim deltaLenghtRow As Double = (offsetLeftLastBeam - offsetLeftFirstBeam) / (countRowsLeftBeam - 1)
            count2 = countRowsLeftBeam
            For i As Integer = 0 To countRowsLeftBeam - 1
                Dim tempDist As Integer = offsetLeftLastBeam - deltaLenghtRow * i
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
            Dim deltaLenghtRow As Double = (offsetRightLastBeam - offsetRightFirstBeam) / (countRowsRightBeam - 1)
            For i As Integer = 0 To countRowsRightBeam - 1
                Dim tempDist As Integer = offsetRightFirstBeam + deltaLenghtRow * i
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
    Public Sub New()
        InitializeComponent()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.readDirectoriesSupport(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "InfrastradaToolsUtility"
                    Exit For
                End If
            Next
        End If
        'Dim supportFile As String = FuncFiles.getFindDirectorySupport("InfrastradaToolsUtility")
        'If Directory.Exists(supportFile) = True Then
        '    generalDir = FuncFiles.TrimPathAfter(supportFile)
        'End If
        Dim templateDir As String = generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\"
        If Directory.Exists(templateDir) = True Then
            Dim allFiles As String() = Directory.GetFiles(templateDir)
            Dim arrayFilesTemplate As String() = {""}
            Dim count As Integer = 1
            For i As Integer = 0 To allFiles.Length - 1
                ReDim Preserve arrayFilesTemplate(count)
                arrayFilesTemplate(count) = IO.Path.GetFileNameWithoutExtension(allFiles(i))
                count += 1
            Next i
            If IsArray(allFiles) = True Then
                CBox_ShemaPlacementBeams.DataSource = arrayFilesTemplate
                CBox_ShemaPlacementBeams.Tag = templateDir
            End If
        Else
            'создаем новый каталог
            Try
                Directory.CreateDirectory(generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\")
                CBox_ShemaPlacementBeams.Tag = templateDir
            Catch ex As Exception
                MsgBox("Не удалось создать каталог: " & generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\")
            End Try
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayTypePlastmentBeams As String() = {"Фиксированная балка", "С расчетом длины балки", "С расчетом максимального зазора"}
        CBox_ListPlacementBeams.DataSource = arrayTypePlastmentBeams
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем все альбомы
        If Directory.Exists(generalDir) = True Then
            Dim directoryBridge As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\"
            Dim countArrayNameAlbom As Integer = 0
            'проверяем наличие директории
            If Directory.Exists(directoryBridge) = True Then
                'получаем директории с альбомами
                If metodPlacmentBeams = Bridges.typePlacementBeam.fixed Then
                    Dim allfolders As String() = Directory.GetDirectories(directoryBridge)
                    If IsArray(allfolders) = True Then
                        For i As Integer = 0 To allfolders.Length - 1
                            Dim tempFolder As String = allfolders(i)
                            Dim nameAlbum As String = New DirectoryInfo(tempFolder).Name
                            '============================================================================================
                            '1 ищем файл xml
                            Dim fullPatchFiles As String = directoryBridge & nameAlbum & "\"
                            Dim xmlBeamsFile As String() = Directory.GetFiles(fullPatchFiles, "*.xml")
                            If xmlBeamsFile.Length = 0 Then
                                Dim xlsBeamsFile As String() = Directory.GetFiles(fullPatchFiles, "*.xlsx")
                                If xlsBeamsFile.Length > 0 Then
                                    'создаем новый файл xml
                                    Dim newNameXml As String = tempFolder & "\" & nameAlbum & ".xml"
                                    Dim boolCreateFiles As Boolean = FuncXML.createXMLFileBeamsByExcel(xlsBeamsFile, newNameXml)
                                End If
                            End If
                            ReDim Preserve arrayNameAlbom(countArrayNameAlbom)
                            arrayNameAlbom(countArrayNameAlbom) = nameAlbum
                            countArrayNameAlbom += 1
                        Next
                    End If
                Else
                    Dim allfolders As String() = Directory.GetDirectories(directoryBridge)
                    If IsArray(allfolders) = True Then
                        For i As Integer = 0 To allfolders.Length - 1
                            Dim tempFolder As String = allfolders(i)
                            Dim nameAlbum As String = New DirectoryInfo(tempFolder).Name
                            If nameAlbum Like "Балки индивидуального проектирования" Then
                                '============================================================================================
                                '1 ищем файл xml
                                Dim fullPatchFiles As String = directoryBridge & nameAlbum & "\"
                                Dim xmlBeamsFile As String() = Directory.GetFiles(fullPatchFiles, "*.xml")
                                If xmlBeamsFile.Length = 0 Then
                                    Dim xlsBeamsFile As String() = Directory.GetFiles(fullPatchFiles, "*.xlsx")
                                    If xlsBeamsFile.Length > 0 Then
                                        'создаем новый файл xml
                                        Dim newNameXml As String = tempFolder & "\" & nameAlbum & ".xml"
                                        Dim boolCreateFiles As Boolean = FuncXML.createXMLFileBeamsByExcel(xlsBeamsFile, newNameXml)
                                    End If
                                End If
                                ReDim Preserve arrayNameAlbom(countArrayNameAlbom)
                                arrayNameAlbom(countArrayNameAlbom) = nameAlbum
                                Exit For
                            End If
                        Next
                    End If
                End If
            Else
                'создаем новый каталог
                Try
                    Directory.CreateDirectory(generalDir & "\TopomaticRobur\DesignBridge\Beams\")
                    MsgBox("Создан новый каталог мостовых балок: " & generalDir & "\TopomaticRobur\DesignBridge\Beams\")
                    Me.Hide()
                Catch ex As Exception
                    MsgBox("Не удалось создать каталог: " & generalDir & "\TopomaticRobur\DesignBridge\Beams\")
                End Try
            End If
            If IsArray(arrayNameAlbom) = True Then
                CBox_AlbumsBeams.DataSource = arrayNameAlbom
                CBox_AlbumsBeams.Tag = directoryBridge
            End If
            'заполняем таблицы
            Dim boolAddProlet As Boolean = FuncCreateTables()
            'метод раскладки
            If metodPlacmentBeams = Bridges.typePlacementBeam.float Or metodPlacmentBeams = Bridges.typePlacementBeam.maxClearence Then
                ButtonCreateArr.Enabled = False
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка =создать новый проект=
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles ButtonCreateArr.Click
        arrangementProject = civilBridgeProject.createArrangementModel()
        If IsNothing(arrangementProject) = False Then
            ActivDocument = arrangementProject.Drawing
            'получаем имя созданного подобъекта
            Dim nameProject = ApplicationHost.Current.Plugins.Execute("getname", New Object() {arrangementProject})
            CBox_ListModelStructures.DataSource = civilBridgeProject.listNameArrangementModels()
            CBox_ListModelStructures.Text = nameProject
        Else
            MsgBox("Ошибка при создании навого проекта. Попробуйте создать новый проект самостоятельно.")
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор проекта arrangement
    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles CBox_ListModelStructures.SelectedValueChanged
        Dim nameArrangementProject As String = CBox_ListModelStructures.SelectedText
        If nameArrangementProject.Trim.Length > 0 Then
            '=========================================================================================
            'имя мостового сооружения
            arrangementProject = civilStructuresProject.getArrangementModelByIndex(CBox_ListModelStructures.SelectedIndex)
            'получаем все мостовые сооружения текущей модели+создаем нове пустое сооружение
            civilBridgeProject.getBridges(True)
            'получаем славарь с сооружениями
            dictNamesProjectBridge = civilBridgeProject.getDictionaryNamesBridge()
            'заполняем список
            If IsNothing(dictNamesProjectBridge) = False Then
                If dictNamesProjectBridge.Count > 0 Then
                    idBridge = dictNamesProjectBridge.ElementAt(0).Key
                    CBox_ListNameStructures.DataSource = dictNamesProjectBridge.Values.ToList
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор сооружения
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles CBox_ListNameStructures.SelectedValueChanged
        Dim nameBridge As String = CBox_ListNameStructures.SelectedValue
        Dim indexBridge As Integer = CBox_ListNameStructures.SelectedIndex
        'если список имен мостов пустой, заново пытаемся его получить
        If dictNamesProjectBridge.Count = 0 Then
            If IsNothing(civilBridgeProject) = False Then
                dictNamesProjectBridge = civilBridgeProject.getDictionaryNamesBridge()
            End If
        End If
        'выбор корректен
        If indexBridge > -1 And dictNamesProjectBridge.Count > 0 Then
            idBridge = dictNamesProjectBridge.ElementAt(indexBridge).Key
            If idBridge.Trim.Length > 0 Then
                If dictNamesProjectBridge.ContainsKey(idBridge) = True Then
                    Dim dataStructuresBridge As StructureElement = civilBridgeProject.getDataBridgeByID(idBridge)
                    If IsNothing(dataStructuresBridge) = False Then
                        Dim userBridge As Bridges = dataStructuresBridge.getBridge()
                        If IsNothing(userBridge) = False Then
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'заполняем форму
                            boolWriteData = False
                            If userBridge.centerAxis = True Then
                                ChB_CenterBeam.Checked = True
                            Else
                                ChB_CenterBeam.Checked = False
                            End If
                            If userBridge.projectSurfaceName.Trim.Length = 0 Then
                                CB_SurfaceFromAlign.Checked = True
                            Else
                                Dim listSurface As List(Of String) = CBox_ListProjectSurfaces.DataSource
                                If IsNothing(listSurface) = False Then
                                    If listSurface.Contains(userBridge.projectSurfaceName) = True Then
                                        CBox_ListProjectSurfaces.Text = userBridge.projectSurfaceName
                                    End If
                                End If
                            End If
                            If userBridge.AlignmentName.Trim.Length > 0 Then
                                Dim arrayAlign As List(Of String) = CBox_ListAxisRoads.DataSource
                                If arrayAlign.Count > 0 Then
                                    If arrayAlign.Contains(userBridge.AlignmentName) = True Then
                                        CBox_ListAxisRoads.Text = userBridge.AlignmentName
                                    End If
                                End If
                            End If
                            boolWriteData = False
                            'DataGridView1.Rows.Clear()
                            If userBridge.ProletCount > 1 Then
                                NUpD_CountProlet.Value = userBridge.ProletCount
                            End If
                            If userBridge.LeftRowsCount > 0 Then
                                NUpD_CountLeftRows.Value = userBridge.LeftRowsCount
                            End If
                            If userBridge.RightRowsCount > 0 Then
                                NUpD_CountRightRows.Value = userBridge.RightRowsCount
                            End If
                            If userBridge.LeftStructureWidth > 0 Then
                                NUpD_dimLeftBridge.Value = CInt(userBridge.LeftStructureWidth * 1000)
                            End If
                            If userBridge.RightStructureWidth > 0 Then
                                NUpD_dimRightBridge.Value = CInt(userBridge.RightStructureWidth * 1000)
                            End If
                            Dim countRows As Integer = userBridge.LeftRowsCount + userBridge.RightRowsCount
                            If userBridge.centerAxis = True Then
                                countRows += 1
                            End If
                            FuncCreateTables()
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'получаем все оси опор
                            Dim dictionaryAlbum As Dictionary(Of String, String()) = New Dictionary(Of String, String())
                            'получаем все элементы мостового сооружения
                            dictionaryBridgeElements = userBridge.getBridgeObjects(dataStructuresBridge.DWGEntity)
                            'получаем все оси опор
                            dictionaryBridgePillars = userBridge.getPillars(dictionaryBridgeElements)
                            'получаем все балки
                            dictionaryBridgeBeams = userBridge.getBeams(dictionaryBridgeElements)
                            'получаем траектории раскладки балок
                            dictionaryBridgeTraectoryBeams = userBridge.getTrajectoryPlacementBeams(dictionaryBridgeElements)
                            '=========================================================================================
                            '1. заполняем таблицу опор
                            If dictionaryBridgePillars.Count > 0 Then
                                For i As Integer = 0 To dictionaryBridgePillars.Count - 1
                                    Dim boolFindPillar As Boolean = False
                                    Dim userListPillar As List(Of StructureElement) = dictionaryBridgePillars.ElementAt(i).Value
                                    If userListPillar.Count > 0 Then
                                        Dim dataElement As StructureElement = userListPillar(1)
                                        If IsNothing(dataElement) = False Then
                                            Dim userPillar As Pillar = dataElement.getPillar()
                                            If IsNothing(userPillar) = False Then
                                                If userPillar.Number > 0 Then
                                                    DG_PillarsProperties.Rows(i).Tag = userPillar.Number
                                                    If userPillar.Defining = True Then
                                                        Dim chbox As DataGridViewCheckBoxCell = DG_PillarsProperties.Rows(i).Cells(0)
                                                        chbox.Value = True
                                                    End If
                                                    DG_PillarsProperties.Rows(i).Cells(1).Value = Val(userPillar.Clearence) * 1000
                                                    DG_PillarsProperties.Rows(i).Cells(2).Value = Val(userPillar.RightClearence) * 1000
                                                    DG_PillarsProperties.Rows(i).Cells(3).Value = Val(userPillar.SiteMonolit) * 1000
                                                    DG_PillarsProperties.Rows(i).Cells(4).Value = dataElement.DWGEntity.ObjectID
                                                    boolFindPillar = True
                                                End If
                                            Else
                                            End If
                                        End If
                                    End If
                                    If boolFindPillar = False Then
                                        DG_PillarsProperties.Rows(i).Tag = i + 1
                                        DG_PillarsProperties.Rows(i).Cells(1).Value = 0
                                        DG_PillarsProperties.Rows(i).Cells(2).Value = 0
                                        DG_PillarsProperties.Rows(i).Cells(3).Value = 0
                                    End If
                                Next i
                            End If
                            '=============================================================================================================
                            '2. заполняем таблицу марками балок
                            If IsNothing(dictionaryBridgeBeams) = False Then
                                If dictionaryBridgeBeams.Count > 0 Then
                                    For i As Integer = 0 To dictionaryBridgeBeams.Count - 1
                                        Dim numberProlet As Integer = dictionaryBridgeBeams.ElementAt(i).Key
                                        Dim dictBeamsProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.ElementAt(i).Value
                                        If IsNothing(dictBeamsProlet) = True Then Continue For
                                        If dictBeamsProlet.Count > 0 Then
                                            For j As Integer = 0 To dictBeamsProlet.Count - 1
                                                Dim numberRow As Integer = dictBeamsProlet.ElementAt(j).Key
                                                Dim dataElement As StructureElement = dictBeamsProlet.ElementAt(j).Value
                                                If IsNothing(dataElement) = False Then
                                                    Dim userBeam As BeamI = dataElement.getBeamI()
                                                    If IsNothing(userBeam) = False Then
                                                        Dim modelBeam As String = userBeam.model
                                                        Dim nameAlbum As String = userBeam.nameAlbum
                                                        'загоняем все балки которые есть в альбоме
                                                        If dictionaryAlbum.ContainsKey(nameAlbum) = False Then
                                                            Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
                                                            Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
                                                            'находим файлы xls в директории альбома
                                                            Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
                                                            If allFilesCSV.Length > 0 Then
                                                                Dim arrayModel As String() = {""}
                                                                Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                                                                If IsArray(arrayModel) = True Then
                                                                    dictionaryAlbum.Add(nameAlbum, arrayModel)
                                                                End If
                                                            End If
                                                        End If
                                                        If dictionaryAlbum.ContainsKey(nameAlbum) = True Then
                                                            Dim arrayModelBeams As String() = dictionaryAlbum.Item(nameAlbum)
                                                            If IsArray(arrayModelBeams) = True Then
                                                                Dim ind As Integer = MathFunction.FuncFindValueToFArray(modelBeam, arrayModelBeams)
                                                                If ind > -1 Then
                                                                    'заполняем таблицу с марками балок
                                                                    If DG_ProletListBeams.RowCount > 0 Then
                                                                        For k As Integer = 0 To DG_ProletListBeams.RowCount - 1
                                                                            Dim numRow As Integer = DG_ProletListBeams.Rows(k).Tag
                                                                            If numRow = numberRow Then
                                                                                If DG_ProletListBeams.ColumnCount >= numberProlet Then
                                                                                    Dim cell2 As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(k).Cells(numberProlet - 1)
                                                                                    cell2.DataSource = arrayModelBeams
                                                                                    cell2.Value = modelBeam
                                                                                    Dim arrayTag As String() = {userBeam.nameAlbum}
                                                                                    DG_ProletListBeams.Rows(k).Cells(numberProlet - 1).Tag = arrayTag
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        Next k
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        'заполняем таблицу с рядами
                                                        If i = 0 Then
                                                            For k As Integer = 0 To DG_RowProperties.RowCount - 1
                                                                Dim numRow As Integer = DG_RowProperties.Rows(k).Tag
                                                                If numRow = numberRow Then
                                                                    DG_RowProperties.Rows(k).Cells(0).Value = Math.Abs(userBeam.axisOffset * 1000)
                                                                    DG_RowProperties.Rows(k).Cells(1).Value = userBeam.offsetSurface * 1000
                                                                    Dim dataPlacementBeams As StructureElement = TrajectoryPlacementBeams.getTrajectoryPlacementBeams(dictionaryBridgeElements, numRow)
                                                                    If IsNothing(dataPlacementBeams) = True Then
                                                                        DG_RowProperties.Rows(k).Cells(2).Value = "Ось трассы"
                                                                    Else
                                                                        Dim plineAxis As DwgPolyline = dataPlacementBeams.DWGEntity
                                                                        If plineAxis.ObjectID > 0 Then
                                                                            DG_RowProperties.Rows(k).Cells(2).Value = Math.Round(plineAxis.Length, 3) & " м."
                                                                            DG_RowProperties.Rows(k).Cells(2).Tag = plineAxis
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next k
                                                            If userBeam.numberRow = 1 Then
                                                                'отступ от оси трассы до первой правой балки
                                                                Dim offsetAxisBeam As Integer = Math.Abs(userBeam.axisOffset * 1000)
                                                                NumericUpDown8.Value = offsetAxisBeam
                                                            ElseIf userBeam.numberRow = -1 Then
                                                                'отступ от оси трассы до первой левой балки
                                                                Dim offsetAxisBeam As Integer = Math.Abs(userBeam.axisOffset * 1000)
                                                                NumericUpDown3.Value = offsetAxisBeam
                                                            ElseIf userBeam.numberRow = userBridge.RightRowsCount Then
                                                                'отступ от оси трассы до первой правой балки
                                                                Dim offsetAxisBeam As Integer = Math.Abs(userBeam.axisOffset * 1000)
                                                                NumericUpDown9.Value = offsetAxisBeam
                                                            ElseIf userBeam.numberRow = -1 * userBridge.LeftRowsCount Then
                                                                'отступ от оси трассы до первой левой балки
                                                                Dim offsetAxisBeam As Integer = Math.Abs(userBeam.axisOffset * 1000)
                                                                NumericUpDown5.Value = offsetAxisBeam
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'взять поверхность из трассы
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CB_SurfaceFromAlign.CheckedChanged
        If CB_SurfaceFromAlign.Checked = True Then
            CBox_ListProjectSurfaces.Enabled = False
        Else
            CBox_ListProjectSurfaces.Enabled = True
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор марки балки во второй таблице
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_ProletListBeams.CellContentClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        Dim nameAlbum As String = CBox_AlbumsBeams.Text
        Dim cellSelect As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(e.RowIndex).Cells(e.ColumnIndex)
        Dim tagCell As Object = cellSelect.Tag
        If IsNothing(tagCell) = False Then
            If IsArray(tagCell) = True Then
                Dim arrayTagCell As String() = cellSelect.Tag
                If arrayTagCell(0) Like nameAlbum Then
                    If cellSelect.Items.Count > 0 Then
                        DG_ProletListBeams.BeginEdit(False)
                        TryCast(DG_ProletListBeams.EditingControl, DataGridViewComboBoxEditingControl).DroppedDown = True
                        Exit Sub
                    End If
                Else
                    'заполняем датадрид
                    Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
                    Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
                    'находим файлы xls в директории альбома
                    Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
                    If allFilesCSV.Length > 0 Then
                        'загружаем из файла xls балки
                        Dim arrayModel As String() = {""}
                        Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                        If boolFindBeams = True Then
                            If IsArray(arrayModel) = True Then
                                cellSelect.DataSource = arrayModel
                                arrayTagCell(0) = nameAlbum
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
            Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
            'находим файлы xls в директории альбома
            Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
            If allFilesCSV.Length > 0 Then
                'загружаем из файла xls балки
                Dim arrayModel As String() = {""}
                Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                If boolFindBeams = True Then
                    If IsArray(arrayModel) = True Then
                        cellSelect.DataSource = arrayModel
                        Dim arrayTagCell As String() = {nameAlbum}
                        cellSelect.Tag = arrayTagCell
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор марки балки во второй таблице
    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_ProletListBeams.CellClick
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex < 0 Then Exit Sub
        Dim nameAlbum As String = CBox_AlbumsBeams.Text
        Dim cellSelect As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(e.RowIndex).Cells(e.ColumnIndex)
        Dim tagCell As Object = cellSelect.Tag
        If IsNothing(tagCell) = False Then
            If IsArray(tagCell) = True Then
                Dim arrayTagCell As String() = cellSelect.Tag
                If arrayTagCell(0) Like nameAlbum Then
                    If cellSelect.Items.Count > 0 Then
                        DG_ProletListBeams.BeginEdit(False)
                        TryCast(DG_ProletListBeams.EditingControl, DataGridViewComboBoxEditingControl).DroppedDown = True
                        Exit Sub
                    End If
                Else
                    'заполняем датадрид
                    Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
                    Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
                    'находим файлы xls в директории альбома
                    Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
                    If allFilesCSV.Length > 0 Then
                        'загружаем из файла xls балки
                        Dim arrayModel As String() = {""}
                        Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                        If boolFindBeams = True Then
                            If IsArray(arrayModel) = True Then
                                cellSelect.DataSource = arrayModel
                                arrayTagCell(0) = nameAlbum
                                cellSelect.Tag = arrayTagCell
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
            Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
            'находим файлы xls в директории альбома
            Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
            If allFilesCSV.Length > 0 Then
                'загружаем из файла xls балки
                Dim arrayModel As String() = {""}
                Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                If boolFindBeams = True Then
                    If IsArray(arrayModel) = True Then
                        cellSelect.DataSource = arrayModel
                        Dim arrayTagCell As String() = {nameAlbum}
                        cellSelect.Tag = arrayTagCell
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор определяющей опоры
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_PillarsProperties.CellContentClick
        ' Проверяем, что клик был по CheckBox ячейке
        Dim indexSelectColumn As Integer = e.ColumnIndex
        If indexSelectColumn = 0 Then
            If e.RowIndex >= 0 Then
                Dim column As DataGridViewColumn = DG_PillarsProperties.Columns(e.ColumnIndex)
                If TypeOf column Is DataGridViewCheckBoxColumn Then
                    For i As Integer = 0 To DG_PillarsProperties.RowCount - 1
                        If i <> e.RowIndex Then
                            Dim chBool As DataGridViewCheckBoxCell = DG_PillarsProperties.Rows(i).Cells(0)
                            chBool.Value = False
                        End If
                    Next
                End If
            End If
        ElseIf indexSelectColumn = 4 Then
            boolButtonSelectPillar = True
            numberSelectRows = e.RowIndex
            Me.Hide()
        End If
    End Sub
    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DG_PillarsProperties.CellMouseUp
        ' Проверяем, что клик был по CheckBox ячейке
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = DG_PillarsProperties.Columns(e.ColumnIndex)
            If TypeOf column Is DataGridViewCheckBoxColumn Then
                For i As Integer = 0 To DG_PillarsProperties.RowCount - 1
                    If i <> e.RowIndex Then
                        Dim chBool As DataGridViewCheckBoxCell = DG_PillarsProperties.Rows(i).Cells(0)
                        chBool.Value = False
                    End If
                Next
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ввод числа пролетов
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NUpD_CountProlet.ValueChanged
        If boolWriteData = False Then
            Dim countProlet As Integer = NUpD_CountProlet.Value
            Dim countColumnsProlet As Integer = DG_ProletListBeams.ColumnCount
            If countProlet > countColumnsProlet Then
                If countProlet > 1 Then
                    For i As Integer = countColumnsProlet To countProlet - 1
                        Dim oldcol = DG_ProletListBeams.Columns(countColumnsProlet - 1)
                        Dim newcol = New DataGridViewComboBoxColumn()
                        newcol.HeaderText = "Пролет " & countProlet
                        newcol.FlatStyle = FlatStyle.Flat
                        newcol.Name = "Prolet" & countProlet - 1
                        newcol.DataPropertyName = oldcol.DataPropertyName
                        newcol.Width = oldcol.Width
                        DG_ProletListBeams.Columns.Add(newcol)

                        Dim posRow As Integer = DG_PillarsProperties.Rows.Add()
                        'заполняем таблицу во вкладке опоры
                        DG_PillarsProperties.Rows(posRow).HeaderCell.Value = "Опора №" & posRow + 1
                        DG_PillarsProperties.Rows(posRow).Tag = posRow + 1
                        DG_PillarsProperties.Rows(posRow).Cells(1).Value = 0
                        DG_PillarsProperties.Rows(posRow).Cells(2).Value = 0
                        DG_PillarsProperties.Rows(posRow).Cells(3).Value = 1000
                    Next
                End If
            ElseIf countProlet < countColumnsProlet Then
                For i As Integer = countProlet To countColumnsProlet - 1
                    If DG_ProletListBeams.ColumnCount > 1 Then
                        DG_ProletListBeams.Columns.RemoveAt(countProlet)
                    End If
                    If DG_PillarsProperties.RowCount > 2 Then
                        DG_PillarsProperties.Rows.RemoveAt(DG_PillarsProperties.RowCount - 1)
                    End If
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'количество рядов слева
    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NUpD_CountLeftRows.ValueChanged
        If boolWriteData = False Then
            If DG_ProletListBeams.RowCount = 0 Then Exit Sub
            Dim countNewRows As Integer = NUpD_CountLeftRows.Value
            Dim countRows As Integer = 0
            For i As Integer = 0 To DG_ProletListBeams.RowCount - 1
                Dim row As DataGridViewRow = DG_ProletListBeams.Rows(i)
                Dim tagRow As Integer = Val(row.Tag)
                If tagRow < 0 Then
                    countRows += 1
                End If
            Next i
            If countRows > 0 Then
                If DG_ProletListBeams.RowCount = 0 Then Exit Sub
                If countRows < countNewRows Then
                    For i As Integer = countRows To countNewRows - 1
                        DG_ProletListBeams.Rows.Insert(0)
                        DG_ProletListBeams.Rows(0).HeaderCell.Value = "Ряд Л-" & countRows + 1
                        DG_ProletListBeams.Rows(0).Tag = -1 * (countRows + 1)

                        DG_RowProperties.Rows.Insert(0)
                        DG_RowProperties.Rows(0).HeaderCell.Value = "Ряд Л-" & countRows + 1
                        DG_RowProperties.Rows(0).Tag = -1 * (countRows + 1)
                        DG_RowProperties.Rows(0).Cells(1).Value = 100
                        DG_RowProperties.Rows(0).Cells(2).Value = "Ось трассы"
                    Next
                ElseIf countRows > countNewRows Then
                    If DG_ProletListBeams.RowCount = 0 Then Exit Sub
                    For i As Integer = countNewRows To countRows - 1
                        DG_ProletListBeams.Rows.RemoveAt(0)
                        DG_RowProperties.Rows.RemoveAt(0)
                    Next
                End If
            End If
            If NUpD_CountLeftRows.Value = 0 Then
                NumericUpDown5.Enabled = False
            Else
                NumericUpDown5.Enabled = True
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'количество рядов справа
    Private Sub NumericUpDown10_ValueChanged(sender As Object, e As EventArgs) Handles NUpD_CountRightRows.ValueChanged
        If boolWriteData = False Then
            If DG_ProletListBeams.RowCount = 0 Then Exit Sub
            Dim countNewRows As Integer = NUpD_CountRightRows.Value
            Dim countRows As Integer = 0
            For i As Integer = 0 To DG_ProletListBeams.RowCount - 1
                Dim row As DataGridViewRow = DG_ProletListBeams.Rows(i)
                Dim tagRow As Integer = Val(row.Tag)
                If tagRow > 0 Then
                    countRows += 1
                End If
            Next i
            If countRows > 0 Then
                If countRows < countNewRows Then
                    If DG_ProletListBeams.RowCount = 0 Then Exit Sub
                    For i As Integer = countRows To countNewRows - 1
                        Dim posRow As Integer = DG_ProletListBeams.Rows.Add()
                        DG_ProletListBeams.Rows(posRow).HeaderCell.Value = "Ряд П-" & countRows + 1
                        DG_ProletListBeams.Rows(posRow).Tag = (countRows + 1)

                        posRow = DG_RowProperties.Rows.Add()
                        DG_RowProperties.Rows(posRow).HeaderCell.Value = "Ряд П-" & countRows + 1
                        DG_RowProperties.Rows(posRow).Tag = (countRows + 1)
                        DG_RowProperties.Rows(posRow).Cells(1).Value = 100
                        DG_RowProperties.Rows(posRow).Cells(2).Value = "Ось трассы"
                    Next
                ElseIf countRows > countNewRows Then
                    If DG_ProletListBeams.RowCount = 0 Then Exit Sub
                    For i As Integer = countNewRows To countRows - 1
                        DG_ProletListBeams.Rows.RemoveAt(DG_ProletListBeams.RowCount - 1)
                        DG_RowProperties.Rows.RemoveAt(DG_RowProperties.RowCount - 1)
                    Next
                End If
            End If
        End If
        If NUpD_CountRightRows.Value = 0 Then
            NumericUpDown9.Enabled = False
        Else
            NumericUpDown9.Enabled = True
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'поставить или убрать осевой ряд
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_CenterBeam.CheckedChanged
        If ChB_CenterBeam.Checked = True Then
            Dim countRows As Integer = 0
            For i As Integer = 0 To DG_ProletListBeams.RowCount - 1
                Dim row As DataGridViewRow = DG_ProletListBeams.Rows(i)
                Dim tagRow As Integer = Val(row.Tag)
                If tagRow < 0 Then
                    countRows += 1
                End If
            Next i
            DG_ProletListBeams.Rows.Insert(countRows)
            DG_ProletListBeams.Rows(countRows).HeaderCell.Value = "Ось"
            DG_ProletListBeams.Rows(countRows).Tag = 0

            DG_RowProperties.Rows.Insert(countRows)
            DG_RowProperties.Rows(countRows).HeaderCell.Value = "Ось"
            DG_RowProperties.Rows(countRows).Tag = 0
            DG_RowProperties.Rows(countRows).Cells(0).Value = 0
            DG_RowProperties.Rows(countRows).Cells(1).Value = 100
            DG_RowProperties.Rows(countRows).Cells(2).Value = "Ось трассы"
            NumericUpDown3.Enabled = False
            NumericUpDown8.Enabled = False
        Else
            For i As Integer = 0 To DG_ProletListBeams.RowCount - 1
                Dim row As DataGridViewRow = DG_ProletListBeams.Rows(i)
                Dim tagRow As Integer = Val(row.Tag)
                If tagRow = 0 Then
                    DG_ProletListBeams.Rows.RemoveAt(i)
                    DG_RowProperties.Rows.RemoveAt(i)
                    NumericUpDown3.Enabled = True
                    NumericUpDown8.Enabled = True
                    Exit Sub
                End If
            Next i
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'перерасчет отступа от оси
    Private Sub РасчитатьОтступОтОсиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РасчитатьОтступОтОсиToolStripMenuItem.Click
        Dim countRowsLeftBeam As Integer = NUpD_CountLeftRows.Value 'количество рядов слева от оси
        Dim offsetLeftFirstBeam As Integer = NumericUpDown3.Value 'смещение первого ряда от оси влево
        Dim offsetLeftLastBeam As Integer = NumericUpDown5.Value 'смещение последнего пяда от оси влево

        Dim countRowsRightBeam As Integer = NUpD_CountRightRows.Value 'количество рядов справа от оси
        Dim offsetRightFirstBeam As Integer = NumericUpDown8.Value 'смещение первого ряда от оси вправо
        Dim offsetRightLastBeam As Integer = NumericUpDown9.Value  'смещение последнего пяда от оси вправо

        If ChB_CenterBeam.Checked = False Then
            'разница расстояний
            Dim deltaLenghtLeftRow As Double = offsetLeftLastBeam - offsetLeftFirstBeam
            deltaLenghtLeftRow = deltaLenghtLeftRow / (countRowsLeftBeam - 1)

            Dim deltaLenghtRightRow As Double = offsetRightLastBeam - offsetRightFirstBeam
            deltaLenghtRightRow = deltaLenghtRightRow / (countRowsRightBeam - 1)

            Dim count As Integer = 1
            For i As Integer = 0 To countRowsLeftBeam - 1
                Dim tagRow As Integer = Val(DG_RowProperties.Rows(i).Tag)
                Dim tempDist As Double = offsetLeftLastBeam - deltaLenghtLeftRow * i
                If tagRow < 0 And DG_RowProperties.RowCount >= i Then
                    'DataGridView3.Rows(i).HeaderCell.Value = "Ряд Л-" & countRowsLeftBeam - i
                    'DataGridView3.Rows(i).Tag = -1 * (countRowsLeftBeam - i)
                    DG_RowProperties.Rows(i).Cells(0).Value = Math.Round(tempDist, 0)
                    count += 1
                End If
            Next

            For i As Integer = 0 To countRowsRightBeam - 1
                Dim tempDist As Double = offsetRightFirstBeam + deltaLenghtRightRow * i
                Dim tagRow As Integer = Val(DG_RowProperties.Rows(count - 1).Tag)
                If tagRow > 0 And DG_RowProperties.RowCount >= i Then
                    'DataGridView3.Rows(count).HeaderCell.Value = "Ряд П-" & i + 1
                    'DataGridView3.Rows(count).Tag = i + 1
                    DG_RowProperties.Rows(count - 1).Cells(0).Value = Math.Round(tempDist, 0)
                    'DataGridView3.Rows(count).Cells(1).Value = 100
                    'DataGridView3.Rows(count).Cells(2).Value = "Ось трассы"
                    count += 1
                End If
            Next i
        Else
            Dim deltaLenghtLeftRow As Double = offsetLeftLastBeam / countRowsLeftBeam
            Dim deltaLenghtRightRow As Double = offsetRightLastBeam / countRowsRightBeam
            offsetRightFirstBeam = deltaLenghtRightRow
            Dim count As Integer = 1
            For i As Integer = 0 To countRowsLeftBeam - 1
                Dim tagRow As Integer = Val(DG_RowProperties.Rows(i).Tag)
                Dim tempDist As Double = offsetLeftLastBeam - deltaLenghtLeftRow * i
                If tagRow < 0 And DG_RowProperties.RowCount >= i Then
                    'DataGridView3.Rows(i).HeaderCell.Value = "Ряд Л-" & countRowsLeftBeam - i
                    'DataGridView3.Rows(i).Tag = -1 * (countRowsLeftBeam - i)
                    DG_RowProperties.Rows(i).Cells(0).Value = Math.Round(tempDist, 0)
                    'DataGridView3.Rows(i).Cells(1).Value = 100
                    'DataGridView3.Rows(i).Cells(2).Value = "Ось трассы"
                    count += 1
                End If
            Next
            count += 1
            For i As Integer = 0 To countRowsRightBeam - 1
                Dim tempDist As Double = offsetRightFirstBeam + deltaLenghtRightRow * i
                Dim tagRow As Integer = Val(DG_RowProperties.Rows(count - 1).Tag)
                If tagRow > 0 And DG_RowProperties.RowCount >= i Then
                    'DataGridView3.Rows(count).HeaderCell.Value = "Ряд П-" & i + 1
                    'DataGridView3.Rows(count).Tag = i + 1
                    DG_RowProperties.Rows(count - 1).Cells(0).Value = Math.Round(tempDist, 0)
                    'DataGridView3.Rows(count).Cells(1).Value = 100
                    'DataGridView3.Rows(count).Cells(2).Value = "Ось трассы"
                    count += 1
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка записать настройки (Библиотека)
    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        formUserTempl.templateDir = CBox_ShemaPlacementBeams.Tag
        If formUserTempl.ListBox1.Items.Count > 0 Then
            formUserTempl.ListBox1.Items.Clear()
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim oldArray As String() = CBox_ShemaPlacementBeams.DataSource
        Dim arraySheme As String() = Nothing
        Dim countArraySheme As Integer = 0
        If IsArray(oldArray) = True Then
            For i As Integer = 0 To oldArray.Length - 1
                Dim zn As String = oldArray(i)
                If IsNothing(zn) = False Then
                    If zn.Trim.Length > 0 Then
                        ReDim Preserve arraySheme(countArraySheme)
                        arraySheme(countArraySheme) = zn
                        countArraySheme += 1
                    End If
                End If
            Next i
        End If
        If IsArray(arraySheme) = True Then
            formUserTempl.ListBox1.Items.AddRange(arraySheme)
            formUserTempl.Label3.Text = CBox_ShemaPlacementBeams.Tag
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем новую форму
        formUserTempl.ShowDialog()
        If formUserTempl.boolWriteFile = True Then 'запись разрешена
            Dim nameFile As String = formUserTempl.TextBox1.Text
            Dim FileNameGeo As String = CBox_ShemaPlacementBeams.Tag & nameFile & ".txt"
            If FileNameGeo.Trim.Length > 0 Then
                If DG_PillarsProperties.Rows.Count > 1 Then
                    Try
                        Dim input As StreamWriter = New StreamWriter(FileNameGeo, False)
                        input.WriteLine("BeamPlactmentBridgeVer=2")
                        input.WriteLine("<Data")
                        input.WriteLine("  DistFirstLeft=" & NumericUpDown3.Value)

                        input.WriteLine("  DistFirstLeft=" & NumericUpDown3.Value)
                        input.WriteLine("  DistLastLeft=" & NumericUpDown5.Value)
                        input.WriteLine("  CountRowsLeft=" & NUpD_CountLeftRows.Value)

                        input.WriteLine("  DistFirstRight=" & NumericUpDown8.Value)
                        input.WriteLine("  DistLastRight=" & NumericUpDown9.Value)
                        input.WriteLine("  CountRowsRight=" & NUpD_CountRightRows.Value)

                        input.WriteLine("  AxisBearm=" & ChB_CenterBeam.Checked.ToString)

                        input.WriteLine("  DimBridgeLeft=" & NUpD_dimLeftBridge.Value)
                        input.WriteLine("  DimBridgeRight=" & NUpD_dimRightBridge.Value)

                        input.WriteLine("  CountProlet=" & NUpD_CountProlet.Value)
                        input.WriteLine("Data/>")

                        input.WriteLine("<DataProlet")
                        For i As Integer = 0 To DG_ProletListBeams.Rows.Count - 1
                            Dim strWrite As String = "  " & DG_ProletListBeams.Rows(i).Tag & ";" & DG_ProletListBeams.Rows(i).HeaderCell.Value & ";"
                            For j As Integer = 0 To DG_ProletListBeams.Columns.Count - 1
                                Dim data1 As String = DG_ProletListBeams.Rows(i).Cells(j).Value
                                Dim data2 As String() = DG_ProletListBeams.Rows(i).Cells(j).Tag
                                If IsNothing(data1) = True Then
                                    Continue For
                                End If
                                If j = DG_ProletListBeams.Columns.Count - 1 Then
                                    strWrite = strWrite & data2(0) & ";" & data1
                                Else
                                    strWrite = strWrite & data2(0) & ";" & data1 & ";"
                                End If
                            Next
                            input.WriteLine(strWrite)
                        Next
                        input.WriteLine("DataProlet/>")

                        input.WriteLine("<DataRows")
                        For i As Integer = 0 To DG_RowProperties.Rows.Count - 1
                            Dim strWrite As String = "  " & DG_RowProperties.Rows(i).Tag & ";" & DG_RowProperties.Rows(i).HeaderCell.Value & ";"
                            For j As Integer = 0 To DG_RowProperties.Columns.Count - 1
                                Dim data1 As String = DG_RowProperties.Rows(i).Cells(j).Value
                                If j = DG_RowProperties.Columns.Count - 1 Then
                                    strWrite = strWrite & data1
                                Else
                                    strWrite = strWrite & data1 & ";"
                                End If
                            Next j
                            input.WriteLine(strWrite)
                        Next
                        input.WriteLine("DataRows/>")

                        input.WriteLine("<DataPillars")
                        For i As Integer = 0 To DG_PillarsProperties.Rows.Count - 1
                            Dim strWrite As String = "  " & DG_PillarsProperties.Rows(i).Tag & ";" & DG_PillarsProperties.Rows(i).HeaderCell.Value & ";"
                            For j As Integer = 0 To DG_PillarsProperties.Columns.Count - 1
                                Dim data1 As String = DG_PillarsProperties.Rows(i).Cells(j).Value
                                If j = DG_PillarsProperties.Columns.Count - 1 Then
                                    strWrite = strWrite & data1
                                Else
                                    strWrite = strWrite & data1 & ";"
                                End If
                            Next j
                            input.WriteLine(strWrite)
                        Next
                        input.WriteLine("DataPillars/>")
                        input.Close()
                        MsgBox("Текущая схема успешно сохранена!")
                    Catch ex As System.Exception
                        MsgBox(ex.Message)
                    End Try
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim templateDir As String = generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\"
                If Directory.Exists(templateDir) = True Then
                    Dim allFiles As String() = Directory.GetFiles(templateDir)
                    Dim arrayFilesTemplate As String() = {""}
                    Dim count As Integer = 1
                    For i As Integer = 0 To allFiles.Length - 1
                        ReDim Preserve arrayFilesTemplate(count)
                        arrayFilesTemplate(count) = IO.Path.GetFileNameWithoutExtension(allFiles(i))
                        count += 1
                    Next i
                    If IsArray(allFiles) = True Then
                        CBox_ShemaPlacementBeams.DataSource = arrayFilesTemplate
                        CBox_ShemaPlacementBeams.Text = nameFile
                        CBox_ShemaPlacementBeams.Tag = templateDir
                    End If
                Else
                    'создаем новый каталог
                    Try
                        Directory.CreateDirectory(generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\")
                    Catch ex As Exception
                        MsgBox("Не удалось создать каталог: " & generalDir & "\TopomaticRobur\DesignBridge\UserProperties\BeamsSchema\")
                    End Try
                End If
            End If
        Else
            Dim templateDir As String = CBox_ShemaPlacementBeams.Tag
            If Directory.Exists(templateDir) = True Then
                Dim allFiles As String() = Directory.GetFiles(templateDir)
                Dim arrayFilesTemplate As String() = {""}
                Dim count As Integer = 1
                For i As Integer = 0 To allFiles.Length - 1
                    ReDim Preserve arrayFilesTemplate(count)
                    arrayFilesTemplate(count) = IO.Path.GetFileNameWithoutExtension(allFiles(i))
                    count += 1
                Next i
                If IsArray(allFiles) = True Then
                    CBox_ShemaPlacementBeams.DataSource = arrayFilesTemplate
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'загрузка шаблона
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ShemaPlacementBeams.SelectedIndexChanged
        Dim putshTemplate As String = CBox_ShemaPlacementBeams.Tag & "\" & CBox_ShemaPlacementBeams.Text & ".txt"
        Dim indRow As Integer = 0
        Dim count As Integer = 0
        boolWriteData = False
        If File.Exists(putshTemplate) = True Then
            Try
                Me.Cursor = Cursors.WaitCursor
                DG_PillarsProperties.Rows.Clear()
                DG_ProletListBeams.Rows.Clear()
                DG_RowProperties.Rows.Clear()
                Dim input As StreamReader = New StreamReader(putshTemplate, True)
                Dim flagProletData As Boolean = False
                Dim flagRowData As Boolean = False
                Dim flagPillarsData As Boolean = False
                boolWriteData = True
                Do Until input.EndOfStream
                    count += 1
                    Dim line1 As String = input.ReadLine() 'считываем строку
                    line1 = line1.Trim
                    If line1.Trim.Length = 0 Then Continue Do
                    If line1.IndexOf("DistFirstLeft") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 15))
                        NumericUpDown3.Value = RezDbl
                    ElseIf line1.IndexOf("DistLastLeft") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 14))
                        NumericUpDown5.Value = RezDbl
                    ElseIf line1.IndexOf("CountRowsLeft") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 15))
                        NUpD_CountLeftRows.Value = RezDbl
                    ElseIf line1.IndexOf("DistFirstRight") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 16))
                        NumericUpDown8.Value = RezDbl
                    ElseIf line1.IndexOf("DistLastRight") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 15))
                        NumericUpDown9.Value = RezDbl
                    ElseIf line1.IndexOf("CountRowsRight") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 16))
                        NUpD_CountRightRows.Value = RezDbl
                    ElseIf line1.IndexOf("DimBridgeLeft") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 15))
                        NUpD_dimLeftBridge.Value = RezDbl
                    ElseIf line1.IndexOf("DimBridgeRight") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 16))
                        NUpD_dimRightBridge.Value = RezDbl
                    ElseIf line1.IndexOf("CountProlet") = 0 Then
                        Dim RezDbl As Integer = Val(Mid(line1, 13))
                        NUpD_CountProlet.Value = RezDbl
                    ElseIf line1.IndexOf("AxisBearm") = 0 Then
                        Dim boolRez As String = Mid(line1, 11)
                        If boolRez Like "True" Then
                            ChB_CenterBeam.Checked = True
                        Else
                            ChB_CenterBeam.Checked = False
                        End If
                    ElseIf line1.IndexOf("Data/>") = 0 Then
                        'заполняем таблицы
                        Dim boolAddProlet As Boolean = FuncCreateTables()
                    ElseIf line1.IndexOf("<DataRows") = 0 Then
                        flagRowData = True
                        indRow = 0
                    ElseIf line1.IndexOf("DataRows/>") = 0 Then
                        flagRowData = False
                    ElseIf flagRowData = True Then
                        Dim Blok As String() = line1.Split(";")
                        If IsArray(Blok) = True Then
                            If Blok.Length = DG_RowProperties.Columns.Count + 2 Then
                                DG_RowProperties.Rows(indRow).Cells(0).Value = Val(Blok(2))
                                DG_RowProperties.Rows(indRow).Cells(1).Value = Val(Blok(3))
                                DG_RowProperties.Rows(indRow).Cells(2).Value = Blok(4)
                                indRow += 1
                            End If
                        End If
                    ElseIf line1.IndexOf("<DataPillars") = 0 Then
                        flagPillarsData = True
                        indRow = 0
                    ElseIf line1.IndexOf("DataPillars/>") = 0 Then
                        flagPillarsData = False
                    ElseIf flagPillarsData = True Then
                        Dim Blok As String() = line1.Split(";")
                        If IsArray(Blok) = True Then
                            If Blok.Length > 2 Then
                                Dim chbox As DataGridViewCheckBoxCell = DG_PillarsProperties.Rows(indRow).Cells(0)
                                For i = 2 To Blok.Length - 1
                                    If i = 2 Then
                                        Dim Rez As String = Blok(2)
                                        If Rez Like "True" Then
                                            chbox.Value = True
                                        Else
                                            chbox.Value = False
                                        End If
                                    Else
                                        DG_PillarsProperties.Rows(indRow).Cells(i - 2).Value = Val(Blok(i))
                                    End If
                                Next i
                                indRow += 1
                            End If
                        End If
                    ElseIf line1.IndexOf("<DataProlet") = 0 Then
                        flagProletData = True
                        indRow = 0
                    ElseIf line1.IndexOf("DataProlet/>") = 0 Then
                        flagPillarsData = False
                    ElseIf flagProletData = True Then
                        Dim Blok As String() = line1.Split(";")
                        If IsArray(Blok) = True Then
                            Dim countCells As Integer = 0
                            If Blok.Length > 2 Then
                                For j As Integer = 2 To Blok.Length - 1 Step 2
                                    Dim nameAlbum As String = Blok(j)
                                    Dim nameMark As String = Blok(j + 1)
                                    Dim cmbbox As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(indRow).Cells(countCells)
                                    If dictionaryAlbumBeams.ContainsKey(nameAlbum) = True Then
                                        Dim userBeams As BeamI() = dictionaryAlbumBeams.Item(nameAlbum)
                                        If IsArray(userBeams) = True Then
                                            Dim arrayModel As String() = Nothing
                                            For k As Integer = 0 To userBeams.Length - 1
                                                Dim userBeam As BeamI = userBeams(k)
                                                ReDim Preserve arrayModel(k)
                                                arrayModel(k) = userBeam.model
                                            Next k
                                            cmbbox.DataSource = arrayModel
                                            Dim arrayDataTag As String() = cmbbox.Tag
                                            If IsArray(arrayDataTag) = True Then
                                                arrayDataTag(0) = nameAlbum
                                            Else
                                                ReDim arrayDataTag(0)
                                                arrayDataTag(0) = nameAlbum
                                            End If
                                            cmbbox.Tag = arrayDataTag
                                            If cmbbox.Items.Count > 0 Then
                                                For k As Integer = 0 To cmbbox.Items.Count - 1
                                                    If cmbbox.Items.Item(k) Like nameMark Then
                                                        cmbbox.Value = nameMark
                                                        Exit For
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    Else
                                        'загружаем из файла xls балки
                                        Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
                                        Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
                                        'находим файлы xls в директории альбома
                                        Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
                                        If allFilesCSV.Length > 0 Then
                                            Dim arrayModel As String() = {""}
                                            Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
                                            If boolFindBeams = True Then
                                                If IsArray(arrayModel) = True Then
                                                    Dim cell2 As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(indRow).Cells(countCells)
                                                    cell2.DataSource = arrayModel
                                                    Dim arrayDataTag As String() = cmbbox.Tag
                                                    If IsArray(arrayDataTag) = True Then
                                                        arrayDataTag(0) = nameAlbum
                                                    Else
                                                        ReDim arrayDataTag(0)
                                                        arrayDataTag(0) = nameAlbum
                                                    End If
                                                    cmbbox.Tag = arrayDataTag
                                                End If
                                                If cmbbox.Items.Count > 0 Then
                                                    For k As Integer = 0 To cmbbox.Items.Count - 1
                                                        If cmbbox.Items.Item(k) Like nameMark Then
                                                            cmbbox.Value = nameMark
                                                            Exit For
                                                        End If
                                                    Next k
                                                End If
                                            End If
                                        End If
                                    End If
                                    countCells += 1
                                Next j
                            End If
                        End If
                        indRow += 1
                    End If
                Loop
                input.Close()
            Catch
                MsgBox("Не удалось загрузить файл. Возможно в строке " & count & " содержится ошибка!")
            End Try
        End If
        Me.Cursor = Cursors.Default
        boolWriteData = False
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShowDlg = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОК
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If CBox_ListPlacementBeams.Text Like "Фиксированная балка" Then
            metodPlacmentBeams = 0
        ElseIf CBox_ListPlacementBeams.Text Like "С расчетом длины балки" Then
            metodPlacmentBeams = 1
        ElseIf CBox_ListPlacementBeams.Text Like "С расчетом максимального зазора" Then
            metodPlacmentBeams = 2
        End If
        'If objectBridgeDictionary.Count = 0 Then
        '    Dim boolFindObject As Boolean = FuncBridge.FuncFindAllObjectBridge(ActivDocument, idBridge, objectBridgeDictionary)
        'End If
        If Not TryApplyPKMask() Then
            Exit Sub
        End If
        boolShowDlg = True
        Me.Hide()
    End Sub
    'выбор оси раскладки балок
    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_RowProperties.CellContentClick
        If e.RowIndex > -1 Then
            Dim tagRow As Integer = Val(DG_RowProperties.Rows(e.RowIndex).Tag)
            If e.ColumnIndex = 2 Then
                boolButtonRows = True
                numberSelectRows = e.RowIndex
                Me.Hide()
            End If
        End If
    End Sub

    Private Sub FormPlacementBeams_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SetupToolTips()
    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DG_ProletListBeams.DataError
        Dim indRow As Integer = e.RowIndex
        Dim indColl As Integer = e.RowIndex
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'задать пикет начала раскладки
    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            MaskTB_PK.Enabled = True
        Else
            MaskTB_PK.Enabled = False
        End If
    End Sub

    Private Sub MaskTB_PK_Enter(sender As Object, e As EventArgs) Handles MaskTB_PK.Enter
        MaskTB_PK.TextMaskFormat = MaskFormat.IncludeLiterals
        Dim currentText As String = MaskTB_PK.Text
        MaskTB_PK.Mask = String.Empty
        MaskTB_PK.Text = currentText.Replace("+", String.Empty)
        MaskTB_PK.SelectAll()
    End Sub

    Private Sub MaskTB_PK_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MaskTB_PK.Validating
        e.Cancel = Not TryApplyPKMask()
    End Sub

    Private Function TryApplyPKMask() As Boolean
        If Not CheckBox3.Checked Then
            lastInvalidPKText = Nothing
            Return True
        End If

        Dim inputText As String = MaskTB_PK.Text
        Dim formattedText As String = Nothing
        Dim dynamicMask As String = Nothing
        If Not FuncFormatZn.TryFormatPKText(inputText, formattedText, dynamicMask) Then
            If Not String.Equals(lastInvalidPKText, inputText, StringComparison.Ordinal) Then
                MessageBox.Show("Ошибка формата! Введите неотрицательный пикет, например: 1+23.123 или 123312.111")
                lastInvalidPKText = inputText
            End If
            Return False
        End If

        MaskTB_PK.Mask = String.Empty
        MaskTB_PK.Mask = dynamicMask
        MaskTB_PK.TextMaskFormat = MaskFormat.IncludeLiterals
        MaskTB_PK.Text = formattedText
        lastInvalidPKText = Nothing
        Return True
    End Function

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'заполнение таблицы пролета
    Private Sub ЗаполнитьВнизToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗаполнитьВнизToolStripMenuItem.Click
        Dim cellColl As DataGridViewSelectedCellCollection = DG_ProletListBeams.SelectedCells
        If cellColl.Count > 0 Then
            Dim cellSelect As DataGridViewComboBoxCell = cellColl.Item(0)
            Dim arrayData As String() = cellSelect.DataSource
            Dim textCell As String = cellSelect.EditedFormattedValue
            Dim tagCell As String() = cellSelect.Tag
            If cellSelect.RowIndex < DG_ProletListBeams.RowCount - 1 Then
                For i As Integer = cellSelect.RowIndex + 1 To DG_ProletListBeams.RowCount - 1
                    Dim newCell As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(i).Cells(cellSelect.ColumnIndex)
                    newCell.DataSource = arrayData
                    newCell.Value = textCell
                    newCell.Tag = tagCell
                Next i
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListModelStructures.SelectedIndexChanged

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'поменять альбом
    Private Sub СменитьАльбомToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СменитьАльбомToolStripMenuItem.Click
        Dim nameAlbum As String = CBox_AlbumsBeams.Text
        If IsNothing(nameAlbum) = True Then Exit Sub
        If nameAlbum.Trim.Length = 0 Then Exit Sub
        Dim selectCells As DataGridViewSelectedCellCollection = DG_ProletListBeams.SelectedCells
        Dim selectCell As DataGridViewCell = selectCells.Item(0)
        Dim cellSelect As DataGridViewComboBoxCell = selectCell
        Dim nameFolderBearm As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum
        Dim nameAlbumXml As String = New DirectoryInfo(nameFolderBearm).Name
        'находим файлы xls в директории альбома
        Dim allFilesCSV As String() = Directory.GetFiles(nameFolderBearm, "*.xml")
        If allFilesCSV.Length > 0 Then
            'загружаем из файла xls балки
            Dim arrayModel As String() = {""}
            Dim boolFindBeams As Boolean = FuncXML.readAlbumBeamsFromXMLFiles(allFilesCSV(0), dictionaryAlbumBeams, arrayModel, nameAlbumXml)
            If boolFindBeams = True Then
                If IsArray(arrayModel) = True Then
                    cellSelect.DataSource = arrayModel
                    Dim arrayTagCell As String() = {nameAlbum}
                    cellSelect.Tag = arrayTagCell
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DG_ProletListBeams.CellMouseEnter
        If e.RowIndex > -1 And e.ColumnIndex > -1 Then
            Dim cellSelect As DataGridViewComboBoxCell = DG_ProletListBeams.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim tagCell As Object = cellSelect.Tag
            If IsNothing(tagCell) = False Then
                If IsArray(tagCell) = True Then
                    Dim cellBounds As Rectangle = DG_ProletListBeams.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False)
                    Dim cellLocation As Point = DG_ProletListBeams.PointToScreen(New Point(cellBounds.Left, cellBounds.Bottom))
                    ' Показываем подсказку
                    Dim text As String = tagCell(0)
                    ToolTip2.Show(text, DG_ProletListBeams, cellBounds.Left + 5, cellBounds.Bottom + 5)
                End If
            End If
        End If
    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора схемы раскладки балок
    Private Sub ComboBox7_MouseHover(sender As Object, e As EventArgs) Handles CBox_ShemaPlacementBeams.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub ComboBox7_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ShemaPlacementBeams.MouseLeave
        ToolTip1.Hide(CBox_ShemaPlacementBeams)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора проекта
    Private Sub CBox_ListModelStructures_MouseHover(sender As Object, e As EventArgs) Handles CBox_ListModelStructures.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_ListModelStructures_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ListModelStructures.MouseLeave
        ToolTip1.Hide(CBox_ListModelStructures)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора моста
    Private Sub CBox_ListNameStructures_MouseHover(sender As Object, e As EventArgs) Handles CBox_ListNameStructures.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_ListNameStructures_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ListNameStructures.MouseLeave
        ToolTip1.Hide(CBox_ListNameStructures)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора поверхности
    Private Sub CBox_ListProjectSurfaces_MouseHover(sender As Object, e As EventArgs) Handles CBox_ListProjectSurfaces.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_ListProjectSurfaces_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ListProjectSurfaces.MouseLeave
        ToolTip1.Hide(CBox_ListProjectSurfaces)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора оси трассы
    Private Sub CBox_ListAxisRoads_MouseHover(sender As Object, e As EventArgs) Handles CBox_ListAxisRoads.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_ListAxisRoads_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ListAxisRoads.MouseLeave
        ToolTip1.Hide(CBox_ListAxisRoads)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора шаблона
    Private Sub CBox_ListTemplateXML_MouseHover(sender As Object, e As EventArgs) Handles CBox_ListTemplateXML.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip1.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_ListTemplateXML_MouseLeave(sender As Object, e As EventArgs) Handles CBox_ListTemplateXML.MouseLeave
        ToolTip1.Hide(CBox_ListTemplateXML)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'всплывающая подсказка для выбора альбома балок
    Private Sub CBox_AlbumsBeams_MouseHover(sender As Object, e As EventArgs) Handles CBox_AlbumsBeams.MouseHover
        Dim cb As ComboBox = CType(sender, ComboBox)
        If cb.SelectedItem IsNot Nothing Then
            ToolTip2.Show(cb.SelectedItem.ToString(), cb, cb.Width, 0, 3000)
        End If
    End Sub

    Private Sub CBox_AlbumsBeams_MouseLeave(sender As Object, e As EventArgs) Handles CBox_AlbumsBeams.MouseLeave
        ToolTip2.Hide(CBox_AlbumsBeams)
    End Sub

    Private Sub ЗаполнитьТолщинуПокрытияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗаполнитьТолщинуПокрытияToolStripMenuItem.Click
        Dim thickness As Integer = DG_RowProperties.Rows(0).Cells(1).Value
        If thickness > 0 Then
            For i As Integer = 0 To DG_RowProperties.RowCount - 1
                DG_RowProperties.Rows(i).Cells(1).Value = thickness
            Next i
        End If
    End Sub
    'выбор альбома с балками
    Private Sub CBox_AlbumsBeams_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_AlbumsBeams.SelectedIndexChanged
        Dim nameAlbum As String = CBox_AlbumsBeams.Text
        Dim nameAlbumXML As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum & ".xml"
        Dim nameAlbumXLS As String = generalDir & "\TopomaticRobur\DesignBridge\Beams\" & nameAlbum & ".xlsx"
        If File.Exists(nameAlbumXLS) = True Then
            If File.Exists(nameAlbumXML) = False Then
                Dim arrayAlbumXLS As String() = {nameAlbumXLS}
                Dim createXML As Boolean = FuncXML.createXMLFileBeamsByExcel(arrayAlbumXLS, nameAlbumXML)
            End If
        End If
    End Sub

    Private Sub MaskTB_PK_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskTB_PK.MaskInputRejected

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
End Class