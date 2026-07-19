Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents.DocumentStructures
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Xml
Imports BridgeSectionMenu.Messaging
Imports BridgeSectionMenu.Services
Imports CivilEnginStructures.StructureElement
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Topomatic
Imports Topomatic.Acax.Export
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.Alg.Layers
Imports Topomatic.Alg.Layers.BaseSectionLayer
Imports Topomatic.Alg.Prf
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Alg.Survey.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Design
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Pipes.Layers.Plan.Nodes
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Sfc
Imports Topomatic.Sites.Core
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Runtime
Imports SMB = BridgeSectionMenu.Messaging.SectionMessageBus
Public Class PanelProjectBridge
    Public CadViewPanel As CadView = Nothing
    Public ActivDocumentPanel As Topomatic.Dwg.Drawing = Nothing
    Public projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
    Public bridgeProject As ProjectBridge = Nothing
    Public arrProject As ArrangementModel = Nothing
    Public selectIDObject As UInteger = 0
    Public idBridge As String = ""
    Public dataBridge As StructureElement = New StructureElement
    Public userBridge As Bridges = Nothing
    'рабочие файлы и директории
    Public generalDir As String = ""
    Public templateXML As String = ""
    Public beamsDir As String = ""
    Public projectAlignment As Alignment = Nothing
    Public projectPolylineAlignment As Polyline3D = New Polyline3D
    Public projectSurface As Surface = Nothing
    Public egSurface As Surface = Nothing
    Public arrayNameAlbom As String() = {""}
    Public dictionaryBridge As Dictionary(Of String, StructureElement) = Nothing 'словарь с мостами активного чертежа
    Public dictNamesProjectBridge As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = New Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) 'словарь с элементами чертежа
    Public putchPile As String = ""
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public X1 As Integer = 0
    Public Y1 As Integer = 0
    Public X2 As Integer = 0
    Public Y2 As Integer = 0
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'балки
    Public axisBeamDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public elemetsBeamDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public axisPillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public brigeGeneralAxisDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public brigeBoundDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public monolitSitesDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'опоры
    Public styleNameTable As String = "Опоры мостовых сооружений"
    Public rigelPillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public racksPillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public icecutterPillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public grillagePillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))
    Public pilePillarsDictionary As Dictionary(Of String, String(,)) = New Dictionary(Of String, String(,))

    Public docRigel As ConstructionDocument = Nothing
    Public docRoks As ConstructionDocument = Nothing
    Public docIcecutter As ConstructionDocument = Nothing
    Public docGrillage As ConstructionDocument = Nothing
    Public docPile As ConstructionDocument = Nothing
    Private layerRigel As DwgLayer = Nothing
    Private colorRigel As CadColor = New CadColor(7)
    Private nameTypeLineRigel As DwgLinetype = Nothing
    Private ScaleTypeLineRigel As Integer = 1
    Private widthTypeLineRigel As Integer = 20
    'ось ригеля
    Private layerAxisRigel As DwgLayer = Nothing
    Private colorAxisRigel As CadColor = New CadColor(10)
    Private nameTypeLineAxisRigel As DwgLinetype = Nothing
    Private ScaleTypeLineAxisRigel As Integer = 1
    Private widthTypeLineAxisRigel As Integer = 20
    'стойка
    Private layerRack As DwgLayer = Nothing
    Private colorRack As CadColor = New CadColor(7)
    Private nameTypeLineRack As DwgLinetype = Nothing
    Private ScaleTypeLineRack As Integer = 1
    Private widthTypeLineRack As Integer = 20
    'ось стойки
    Private layerAxisRack As DwgLayer = Nothing
    Private colorAxisRack As CadColor = New CadColor(10)
    Private nameTypeLineAxisRack As DwgLinetype = Nothing
    Private ScaleTypeLineAxisRack As Integer = 1
    Private widthTypeLineAxisRack As Integer = 20
    'массивное тело'
    Dim layerIcecutter As DwgLayer = Nothing
    Dim colorIcecutter As CadColor = New CadColor(7)
    Dim nameTypeLineIcecutter As DwgLinetype = Nothing
    Dim ScaleTypeLineIcecutter As Integer = 1
    Dim widthTypeLineIcecutter As Integer = 20
    'ось массивного тела
    Dim boolAxisIcecutterStyle As Boolean = False
    Dim layerAxisIcecutter As DwgLayer = Nothing
    Dim colorAxisIcecutter As CadColor = New CadColor(10)
    Dim nameTypeLineAxisIcecutter As DwgLinetype = Nothing
    Dim ScaleTypeLineAxisIcecutter As Integer = 1
    Dim widthTypeLineAxisIcecutter As Integer = 20
    'ростверк
    Dim layerGrillage As DwgLayer = Nothing
    Dim colorGrillage As CadColor = New CadColor(7)
    Dim nameTypeLineGrillage As DwgLinetype = Nothing
    Dim ScaleTypeLineGrillage As Integer = 1
    Dim widthTypeLineGrillage As Integer = 20
    'ось ростверка
    Dim layerAxisGrillage As DwgLayer = Nothing
    Dim colorAxisGrillage As CadColor = New CadColor(10)
    Dim nameTypeLineAxisGrillage As DwgLinetype = Nothing
    Dim ScaleTypeLineAxisGrillage As Integer = 1
    Dim widthTypeLineAxisGrillage As Integer = 20
    'свая
    Dim layerPile As DwgLayer = Nothing
    Dim colorPile As CadColor = New CadColor(7)
    Dim nameTypeLinePile As DwgLinetype = Nothing
    Dim ScaleTypeLinePile As Integer = 1
    Dim widthTypeLinePile As Integer = 20
    'ось сваи
    Dim layerAxisPile As DwgLayer = Nothing
    Dim colorAxisPile As CadColor = New CadColor(10)
    Dim nameTypeLineAxisPile As DwgLinetype = Nothing
    Dim ScaleTypeLineAxisPile As Integer = 1
    Dim widthTypeLineAxisPile As Integer = 20

    Public keyParamStr As String() = Nothing
    Public nameModel As String = ""
    Public horizontalOffsetStructure As Double = 0
    Public horizontalTOffsetStructure As Double = 0
    Public verticalOffsetStructure As Double = 0
    Public prOffsetStructure As Double = 0

    Public nameModelEdit As String = ""
    Const WINDOW_UID As String = "Редактор сечения"
    Public selectEntModel As List(Of DwgModel3DElement) = New List(Of DwgModel3DElement)
    Public sectionService As SectionWindowManagementService = Nothing

    Private Sub Plan_view_SelectedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim project = ApplicationHost.Current.ActiveProject
        If project Is Nothing Then
            Return
        End If
        Dim window As IDocumentWindow
        If project.TryGetWindow(WINDOW_UID, window) Then
            Dim fwindow = TryCast(window, IFramableDocumentWindow)
            If fwindow IsNot Nothing Then
                Dim layer = TryCast(fwindow.CadView(MySectionLayer.GUID), MySectionLayer)
                If layer IsNot Nothing Then layer.InvalidateSelection(selectEntModel)
            End If
        End If
    End Sub

    Private Sub Frammable_window_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        Dim plan_view = CadViewDesignUtils.OnCadViewSelect(CadViewDesignUtils.PlanCadViewAlias)
        RemoveHandler plan_view.SelectedChanged, AddressOf Plan_view_SelectedChanged
    End Sub

    Private Function OpenWindow() As IFramableDocumentWindow
        Dim project = ApplicationHost.Current.ActiveProject

        If project Is Nothing Then
            Return Nothing
        End If
        Dim window As IDocumentWindow
        If project.TryGetWindow(WINDOW_UID, window) Then
            Return TryCast(window, IFramableDocumentWindow)
        End If
        Dim frammable_window = TryCast(project.AddDocumentWindow(WINDOW_UID, WINDOW_UID, False), IFramableDocumentWindow)
        If frammable_window Is Nothing Then
            Return Nothing
        End If
        frammable_window.CloseButton = True
        frammable_window.Text = "Редактор сечения"
        Dim frame = frammable_window.AddCadViewFrame(Consts.ModelFrame, "Пользовательское сечение")
        frame.Control.SuspendLayout()
        frame.CadView.ShowScreenRotationSetting = False
        frame.CadView.ShowScreenScaleRatio = True
        frame.CadView.MultiSelect = True
        frame.CadView.DraftingSettings.DrawGrid = True
        'Панель слева
        Dim vGrid = New VerticalGrid()
        vGrid.Dock = DockStyle.Left
        vGrid.BackColor = frame.CadView.BackColor
        frame.Control.Controls.Add(vGrid)
        AddHandler frame.CadView.Paint, AddressOf vGrid.DoPaint
        'Панель снизу
        Dim hGrid = New HorizontalGrid()
        hGrid.Dock = DockStyle.Bottom
        hGrid.BackColor = frame.CadView.BackColor
        frame.Control.Controls.Add(hGrid)
        AddHandler frame.CadView.Paint, AddressOf hGrid.DoPaint

        Dim colorChanged As EventHandler = Sub(ByVal s As Object, ByVal eargs As EventArgs)
                                               hGrid.BackColor = (CType(s, CadView)).BackColor
                                               vGrid.BackColor = (CType(s, CadView)).BackColor
                                           End Sub

        AddHandler frame.CadView.BackColorChanged, colorChanged
        Dim layer = New MySectionLayer()
        frame.CadView.AddLayer(layer)
        frame.CadView.SolveLimits(True)
        frame.CadView.Unlock()
        frame.CadView.Invalidate()
        frame.Control.ResumeLayout(True)
        vGrid.SectionLayer = layer
        hGrid.SectionLayer = layer
        Dim plan_view = CadViewDesignUtils.OnCadViewSelect(CadViewDesignUtils.PlanCadViewAlias)
        AddHandler plan_view.SelectedChanged, AddressOf Plan_view_SelectedChanged
        AddHandler frammable_window.FormClosing, AddressOf Frammable_window_FormClosing
        Return frammable_window
    End Function

    ' Рекурсивный поиск узла по Text
    'поиск дочернего нода
    Public Function FindChildNodeByName(parentNode As TreeNode, nodeName As String) As TreeNode
        For Each child As TreeNode In parentNode.Nodes
            If child.Name = nodeName Then
                Return child
            End If
        Next
        Return Nothing ' Не найден
    End Function
    Public Shared Function FindNodeByText(nodes As TreeNodeCollection, text As String) As TreeNode
        For Each node As TreeNode In nodes
            If node.Text = text Then
                Return node
            End If
            Dim childResult As TreeNode = FindNodeByText(node.Nodes, text)
            If childResult IsNot Nothing Then
                Return childResult
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function FindTreeNodeByTag(nodes As TreeNodeCollection, index As Integer, level As Integer) As TreeNode
        For Each node As TreeNode In nodes
            ' Если Tag совпадает, возвращаем узел
            If node.Index = index And node.Level = level Then
                Return node
            End If
            ' Рекурсивно ищем во вложенных узлах
            Dim foundNode As TreeNode = FindTreeNodeByTag(node.Nodes, index, level)
            If foundNode IsNot Nothing Then
                Return foundNode
            End If
        Next
        ' Если узел не найден, возвращаем Nothing
        Return Nothing
    End Function
    Private Function FindTreeNodeByLevelAndIndex(nodes As TreeNodeCollection, targetLevel As Integer, targetIndex As Integer, Optional currentLevel As Integer = 0) As TreeNode
        For Each node As TreeNode In nodes
            ' Если текущий уровень совпадает с целевым и индекс правильный
            If currentLevel = targetLevel Then
                If targetIndex = 0 Then
                    Return node ' Найден нужный узел
                Else
                    targetIndex -= 1 ' Уменьшаем счетчик индекса
                End If
            Else
                ' Рекурсивно ищем в дочерних узлах (увеличиваем уровень)
                Dim foundNode As TreeNode = FindTreeNodeByLevelAndIndex(node.Nodes, targetLevel, targetIndex, currentLevel + 1)
                If foundNode IsNot Nothing Then
                    Return foundNode
                End If
            End If
        Next
        Return Nothing ' Узел не найден
    End Function
    Public Shared Function FindNodeByTagBFS(treeView As System.Windows.Forms.TreeView, idElement As String) As System.Windows.Forms.TreeNode
        Dim queue As New Queue(Of TreeNode)()
        ' Добавляем все корневые узлы в очередь
        For Each node As TreeNode In treeView.Nodes
            queue.Enqueue(node)
        Next
        While queue.Count > 0
            Dim currentNode As TreeNode = queue.Dequeue()
            If currentNode.Tag IsNot Nothing Then
                Dim strTag As Object = currentNode.Tag
                If IsArray(strTag) = True Then
                    Dim arrayTag As String() = currentNode.Tag
                    If arrayTag.Length > 3 Then
                        If arrayTag(0) Like "Ригель" Then
                            Dim a = 0
                        End If
                        Dim idTagElement As String = arrayTag(3)
                        If idTagElement Like idElement Then
                            Return currentNode
                        End If
                    End If
                End If
            End If
            ' Добавляем все дочерние узлы в очередь
            For Each childNode As TreeNode In currentNode.Nodes
                queue.Enqueue(childNode)
            Next
        End While
        Return Nothing
    End Function
    Public Shared Function FindNodeByTag_ObjectID(treeView As System.Windows.Forms.TreeView, idObject As String) As System.Windows.Forms.TreeNode
        If treeView Is Nothing Then Return Nothing
        Dim queue As New Queue(Of TreeNode)()
        For Each node As TreeNode In treeView.Nodes
            queue.Enqueue(node)
        Next
        While queue.Count > 0
            Dim currentNode As TreeNode = queue.Dequeue()
            ' Проверка Tag более безопасным способом
            If currentNode.Tag IsNot Nothing Then
                ' Пытаемся привести к массиву строк
                Dim arrayTag As String() = TryCast(currentNode.Tag, String())
                If arrayTag IsNot Nothing AndAlso arrayTag.Length > 4 Then
                    ' Индекс 4 соответствует 5-му элементу (0-based)
                    Dim idTagElement As String = arrayTag(4)
                    ' Проверяем, не пустое ли значение
                    If Not String.IsNullOrEmpty(idTagElement) Then
                        ' Используем String.Compare или точное сравнение
                        ' в зависимости от требований
                        If idTagElement = idObject Then ' или idTagElement Like idObject
                            Return currentNode
                        End If
                    End If
                End If
            End If
            For Each childNode As TreeNode In currentNode.Nodes
                queue.Enqueue(childNode)
            Next
        End While
        Return Nothing
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'СООРУЖЕНИЕ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'обновить структуру проекта
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        UpdateStructureMenu.Show(Button1, New System.Drawing.Point(0, Button1.Height))
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ReportMenu.Show(Button3, New System.Drawing.Point(0, Button3.Height))
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MenuBeams.Show(Button4, New System.Drawing.Point(0, Button4.Height))
    End Sub
    'редактировать мост
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        BridgeMenu.Show(Button5, New System.Drawing.Point(0, Button5.Height))
    End Sub
    'редактировать опоры
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        PillarsMenu.Show(Button6, New System.Drawing.Point(0, Button6.Height))
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        UpdateStructureMenu.Show(Button1, New System.Drawing.Point(0, Button1.Height))
    End Sub
    'обновить сооружение
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        UpdateStructureBridgePanel(CBox_ListNamesArrProject.Text, True)
    End Sub
    '===================================================================================================================
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'работа палитры
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub PanelProjectBridge_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
    Public Sub New()
        InitializeComponent()
        Me.BackColor = Color.FromArgb(209, 223, 224)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        'шаблон оформления
        Dim directorySupport As String = generalDir & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If xmlFiles.Count > 0 Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    Dim tempXml As String = xmlFiles(i)
                    Dim boolTempFile As Boolean = IO.File.Exists(tempXml)
                    If boolTempFile = True Then
                        templateXML = tempXml
                        Exit For
                    End If
                Next
            End If
        End If
        ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'инициализируем новый проект
        bridgeProject = New ProjectBridge()
        'получаем для данного проекта список подъобъектов
        Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures 'модели
        Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
        Dim listEgSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
        Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
        'заполняем комбо с именами проектов
        Dim listProject As List(Of String) = projectCivil.listNameArrangementModels()
        CBox_ListNamesArrProject.DataSource = listProject
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем все альбомы
        Dim countArrayNameAlbom As Integer = 0
        If Directory.Exists(generalDir) = True Then
            If IO.Directory.Exists(generalDir & "\TopomaticRobur\DesignBridge\Beams\") = True Then
                beamsDir = generalDir & "\TopomaticRobur\DesignBridge\Beams\"
                'получаем директории с альбомами
                Dim allfolders As String() = Directory.GetDirectories(beamsDir)
                If allfolders.Length > 0 Then
                    For i As Integer = 0 To allfolders.Length - 1
                        Dim tempFolder As String = allfolders(i)
                        Dim nameAlbum As String = New DirectoryInfo(tempFolder).Name
                        '============================================================================================
                        '1 ищем файл xml
                        Dim fullPatchFiles As String = beamsDir & nameAlbum & "\"
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
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Сваи
        putchPile = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Сваи\"

        ToolTip1.SetToolTip(Me.Button1, "Обновить структуру проекта")
        ToolTip1.SetToolTip(Me.Button2, "Перестроить сооружение")
        'AddHandler Button3.Click, AddressOf Button3_Click
        'AddHandler Button4.Click, AddressOf Button4_Click
        'AddHandler Button6.Click, AddressOf Button6_Click
        'AddHandler Button5.Click, AddressOf Button5_Click
        'Dim nameModel As String = CBox_ListNamesArrProject.Text
        'If IsNothing(nameModel) = False Then
        '    If nameModel.Trim.Length > 0 Then
        '        UpdateStructureBridgePanel(CBox_ListNamesArrProject.Text)
        '    End If
        'End If
        'AddHandler SectionMessageBus.SectionSelected, AddressOf OnSectionSelected
        'sectionService = SectionWindowManagementService.Instance
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую пустую структуру
    Public Sub CreateBridgeStructure(ByVal IDBridge As String, ByVal userBridge As Bridges)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем дерево структуры (пустое дерево)
        Dim userDataDefault As StructureElement = New StructureElement("", classStructure.OtherObject, typeObject.OtherElement, "", "", IDBridge, "", "", "")
        Dim conditionalBridgeElementDefault As conditionalElement = New conditionalElement(0, userDataDefault)
        Dim userNode0 As TreeNode = New TreeNode(userBridge.NameBridge, 0, 0)
        userNode0.Name = StructureElement.typeObject.axisBridge
        userNode0.Tag = conditionalBridgeElementDefault
        TreeView1.Nodes.Add(userNode0)
        '==============================================================================================================
        'Мостовое полотно
        Dim userNode2 As TreeNode = New TreeNode("Мостовое полотно", 7, 7)
        userNode2.Tag = conditionalBridgeElementDefault
        userNode2.Name = "BRIDGE_DECK"
        userNode0.Nodes.Add(userNode2)
        userNode2.ForeColor = Color.Gray
        '================================================================================================================
        'Крайние опоры
        Dim userNode3 As TreeNode = New TreeNode("Крайние опоры", 8, 8)
        userNode3.Tag = conditionalBridgeElementDefault
        userNode3.Name = StructureElement.classStructure.LastPillar
        userNode0.Nodes.Add(userNode3)
        'второй уровень
        For i As Integer = 0 To 1
            Dim numberPillar As Integer = 0
            If i = 1 Then
                numberPillar = userBridge.ProletCount + 1
            Else
                numberPillar = 1
            End If
            Dim nodeLastPillar As TreeNode = New TreeNode("Опора №" & numberPillar, 8, 8)
            Dim conditionalFirstPillar As conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeLastPillar.Name = StructureElement.typeObject.axisPillar
            nodeLastPillar.Tag = conditionalFirstPillar
            nodeLastPillar.ForeColor = Color.Gray
            'заполняем остальные элементы
            Dim nodeElement As TreeNode = New TreeNode("Насадка", 9, 9)
            Dim conditionalElement As conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisNozzle
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Подферменники", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.classStructure.GroupSubFermenters
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Шкафная стенка", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisCabinetWall
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Откосное крыло левое", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisLeftHand
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Откосное крыло правое", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisRightHand
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Обратный открылок левый", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisLeftPostcard
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Обратный открылок правый", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisRightPostcard
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Стойки", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.classStructure.GroupRackPillar
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Ростверк", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisGrillage
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Подготовка", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = StructureElement.typeObject.axisPreparation
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            nodeElement = New TreeNode("Сваи", 9, 9)
            conditionalElement = New conditionalElement(numberPillar, userDataDefault)
            nodeElement.Tag = conditionalElement
            nodeElement.Name = classStructure.GroupPilePillar
            nodeLastPillar.Nodes.Add(nodeElement)
            nodeElement.ForeColor = Color.Gray

            userNode3.Nodes.Add(nodeLastPillar)
        Next i
        '================================================================================================================
        'Промежуточные опоры
        Dim userNode4 As TreeNode = New TreeNode("Промежуточные опоры", 9, 9)
        userNode4.Tag = conditionalBridgeElementDefault
        userNode4.Name = StructureElement.classStructure.MiddlePillar
        userNode0.Nodes.Add(userNode4)
        'второй уровень заполняем опоры
        Dim countPillars As Integer = userBridge.ProletCount + 1
        If countPillars > 2 Then
            For i As Integer = 2 To countPillars - 1
                Dim nodeMiddlePillar As TreeNode = New TreeNode("Опора №" & i, 8, 8)
                Dim conditionalMiddlePillar As conditionalElement = New conditionalElement(i, userDataDefault)
                nodeMiddlePillar.Name = StructureElement.typeObject.axisPillar
                nodeMiddlePillar.Tag = conditionalMiddlePillar

                Dim userNode4_1_2 As TreeNode = New TreeNode("Ось опирания балок №1", 10, 10)
                Dim conditionalAxisBeamsPillar1 As conditionalElement = New conditionalElement(i, userDataDefault)
                conditionalAxisBeamsPillar1.NumberRow = i - 1
                userNode4_1_2.Name = StructureElement.typeObject.axisPillarBeams
                userNode4_1_2.Tag = conditionalAxisBeamsPillar1
                nodeMiddlePillar.Nodes.Add(userNode4_1_2)
                userNode4_1_2.ForeColor = Color.Gray

                Dim userNode4_1_3 As TreeNode = New TreeNode("Ось опирания балок №2", 10, 10)
                Dim conditionalAxisBeamsPillar2 As conditionalElement = New conditionalElement(i, userDataDefault)
                conditionalAxisBeamsPillar2.NumberRow = i
                userNode4_1_3.Name = StructureElement.typeObject.axisPillarBeams
                userNode4_1_3.Tag = conditionalAxisBeamsPillar2
                nodeMiddlePillar.Nodes.Add(userNode4_1_3)
                userNode4_1_3.ForeColor = Color.Gray

                Dim nodeElement As TreeNode = New TreeNode("Ригель", 9, 9)
                Dim conditionalElement As conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = StructureElement.typeObject.axisRigel
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                nodeElement = New TreeNode("Подферменники", 9, 9)
                conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = StructureElement.classStructure.GroupSubFermenters
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                nodeElement = New TreeNode("Стойки", 9, 9)
                conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = StructureElement.classStructure.GroupRackPillar
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                nodeElement = New TreeNode("Ростверк", 9, 9)
                conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = StructureElement.typeObject.axisGrillage
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                nodeElement = New TreeNode("Подготовка", 9, 9)
                conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = StructureElement.typeObject.axisPreparation
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                nodeElement = New TreeNode("Сваи", 9, 9)
                conditionalElement = New conditionalElement(i, userDataDefault)
                nodeElement.Tag = conditionalElement
                nodeElement.Name = classStructure.GroupPilePillar
                nodeMiddlePillar.Nodes.Add(nodeElement)
                nodeElement.ForeColor = Color.Gray

                userNode4.Nodes.Add(nodeMiddlePillar)
            Next i
        End If
        'Пролетные строения
        Dim userNode5 As TreeNode = New TreeNode("Пролетные строения", 4, 4)
        userNode5.Tag = conditionalBridgeElementDefault
        userNode5.Name = "SPAN_STRUCTURES"
        userNode0.Nodes.Add(userNode5)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'балки пролетных строений
        Dim userNode5_1 As TreeNode = New TreeNode("Балки пролетных строений", 12, 12)
        userNode5_1.Tag = conditionalBridgeElementDefault
        userNode5_1.Name = "BEAMS"
        userNode5.Nodes.Add(userNode5_1)
        'Участки омоноличивания балок
        Dim userNode5_2 As TreeNode = New TreeNode("Участки омоноличивания балок", 13, 13)
        userNode5_2.Tag = conditionalBridgeElementDefault
        userNode5_2.Name = "MONOLIT_SITES"
        userNode5.Nodes.Add(userNode5_2)
        'Прочие элементы
        Dim userNode6 As TreeNode = New TreeNode("Прочие элементы", 6, 6)
        userNode6.Tag = conditionalBridgeElementDefault
        userNode6.Name = "OTHER_ELEMENTS"
        userNode0.Nodes.Add(userNode6)

        Dim deskBound As String = StructureElement.GetDescription(StructureElement.classStructure.BoundBridge)
        Dim userNode6_2 As TreeNode = New TreeNode(deskBound)
        Dim conditionalBound As conditionalElement = New conditionalElement(0, userDataDefault)
        userNode6_2.Name = StructureElement.classStructure.BoundBridge
        userNode6_2.Tag = conditionalBound
        userNode6.Nodes.Add(userNode6_2)
        userNode6_2.ForeColor = Color.Gray
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'обновление структуры 
    Public Sub UpdateStructureBridgePanel(ByVal nameModel As String, Optional ByVal boolUpdateProject As Boolean = False)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получаем все подобъекты
        Dim selectedNode As TreeNode = Nothing
        Dim nameSelectModelProject As String = CBox_ListNamesArrProject.Text
        Dim indexArrProject As Integer = CBox_ListNamesArrProject.SelectedIndex
        Dim arrProject = projectCivil.getArrangementModelByIndex(indexArrProject)
        If IsNothing(arrProject) = False Then
            ActivDocumentPanel = arrProject.Drawing
        Else
            MsgBox("Не удалось получить доступ к проекту")
            Exit Sub
        End If
        TreeView1.Nodes.Clear()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'Получаем все элементы сооружения
        If IsNothing(arrProject) = False Then
            Dim ListdataStructuresBridge As List(Of StructureElement) = ProjectBridge.getBridges(arrProject)
            If ListdataStructuresBridge.Count > 0 Then
                For i As Integer = 0 To ListdataStructuresBridge.Count - 1
                    Dim dataBridge As StructureElement = ListdataStructuresBridge.Item(i)
                    Dim idBridge As String = dataBridge.IdStructure
                    Dim userBridge As Bridges = dataBridge.getBridge()
                    If IsNothing(userBridge) = True Then Continue For
                    If IsNothing(dataBridge.DWGEntity) = True Then Continue For
                    If TypeOf dataBridge.DWGEntity Is DwgPolyline Then
                        Dim axisLineBridge As DwgPolyline = dataBridge.DWGEntity
                        If IsNothing(axisLineBridge) = True Then Continue For
                        dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
                        If dictionaryBridgeElements.Count > 0 Then
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'новый пустой мост
                            CreateBridgeStructure(idBridge, userBridge)
                            '==================================================================================================================================
                            'заполняем сооружение
                            '==================================================================================================================================
                            Dim userDataDefault As StructureElement = New StructureElement("", classStructure.OtherObject, typeObject.OtherElement, "", "", idBridge, "", "", "")
                            Dim nodeGBridge As TreeNode() = TreeView1.Nodes.Find(StructureElement.typeObject.axisBridge, True)
                            If IsArray(nodeGBridge) = True Then
                                For k As Integer = 0 To nodeGBridge.Length - 1
                                    Dim DataNode As conditionalElement = nodeGBridge(k).Tag
                                    If IsNothing(DataNode.DataElement) = False Then
                                        If DataNode.DataElement.IdStructure Like idBridge Then
                                            Dim node0 As TreeNode = nodeGBridge(k)
                                            Dim conditionalBridge As conditionalElement = New conditionalElement(0, dataBridge)
                                            node0.Tag = conditionalBridge
                                            node0.ForeColor = Color.Black
                                            Exit For
                                        End If
                                    End If
                                Next k
                            End If
                            '==================================================================================================================================
                            'граница сооружения
                            If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.boundaresBridge) = True Then
                                Dim listDataBound As List(Of StructureElement) = dictionaryBridgeElements.Item(StructureElement.typeObject.boundaresBridge)
                                If listDataBound.Count > 0 Then
                                    Dim nodeBounds As TreeNode() = TreeView1.Nodes.Find(StructureElement.classStructure.BoundBridge, True)
                                    If IsArray(nodeBounds) = True Then
                                        For k As Integer = 0 To nodeBounds.Length - 1
                                            Dim DataNode As conditionalElement = nodeBounds(k).Tag
                                            If IsNothing(DataNode.DataElement) = False Then
                                                If DataNode.DataElement.IdStructure Like idBridge Then
                                                    Dim node0 As TreeNode = nodeBounds(k)
                                                    For k1 As Integer = 0 To listDataBound.Count - 1
                                                        Dim dataBound As StructureElement = listDataBound.Item(k1)
                                                        Dim nameElement As String = "Граница сооружения"
                                                        Dim keyParam As String = dataBound.KeyParameter
                                                        If FuncGSON.IsValidJson(keyParam) = True Then
                                                            Dim tempNameElement As String = FuncGSON.getValue(keyParam, "name")
                                                            If tempNameElement.Trim.Length > 0 Then
                                                                nameElement = tempNameElement
                                                            End If
                                                        End If
                                                        Dim conditionalBound As conditionalElement = New conditionalElement(k1 + 1, dataBound)
                                                        Dim nodeBound As TreeNode = New TreeNode(nameElement, 9, 9)
                                                        nodeBound.Name = StructureElement.typeObject.boundaresBridge
                                                        nodeBound.Tag = conditionalBound
                                                        node0.Nodes.Add(nodeBound)
                                                        nodeBound.ForeColor = Color.Black
                                                    Next k1
                                                    Exit For
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                            End If
                            '==============================================================================================================================
                            'заполняем балки
                            '==============================================================================================================================
                            Dim dictBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
                            If dictBeams.Count > 0 Then
                                Dim arrayBeamsGNodes As TreeNode() = TreeView1.Nodes.Find("BEAMS", True)
                                If arrayBeamsGNodes.Length > 0 Then
                                    For j As Integer = 0 To arrayBeamsGNodes.Length - 1
                                        Dim nodeBeams As TreeNode = arrayBeamsGNodes(j)
                                        Dim condBeams As conditionalElement = nodeBeams.Tag
                                        If IsNothing(condBeams.DataElement) = False Then
                                            If condBeams.DataElement.IdStructure Like idBridge Then
                                                For k As Integer = 0 To dictBeams.Count - 1
                                                    Dim numberProlet As Integer = dictBeams.ElementAt(k).Key
                                                    Dim dataBeams As Dictionary(Of Integer, StructureElement) = dictBeams.ElementAt(k).Value
                                                    Dim userNodeProlet As TreeNode = New TreeNode("Пролет№ " & numberProlet, 12, 12)
                                                    Dim conditionalBridgeElementDefault As conditionalElement = New conditionalElement(numberProlet, userDataDefault)
                                                    Dim conditionalBeamsProlet As conditionalElement = conditionalBridgeElementDefault
                                                    userNodeProlet.Name = "PROLET_BEAMS"
                                                    userNodeProlet.Tag = conditionalBeamsProlet
                                                    nodeBeams.Nodes.Add(userNodeProlet)
                                                    If dataBeams.Count > 0 Then
                                                        For k2 As Integer = 0 To dataBeams.Count - 1
                                                            Dim dataBeam As StructureElement = dataBeams.ElementAt(k2).Value
                                                            If IsNothing(dataBeam) = False Then
                                                                If IsNothing(dataBeam.DWGEntity) = False Then
                                                                    Dim numberRow As Integer = dataBeams.ElementAt(k2).Key
                                                                    Dim strRow As String = FuncFormatZn.getConditionalRow(numberRow)
                                                                    Dim userNodeBeam As TreeNode = New TreeNode("Балка № " & strRow, 12, 12)
                                                                    Dim conditionalBeam As conditionalElement = New conditionalElement(numberRow, dataBeam)
                                                                    conditionalBeam.NumberColl = numberProlet
                                                                    conditionalBeam.NumberRow = numberRow
                                                                    userNodeBeam.Name = StructureElement.typeObject.axisBeam
                                                                    userNodeBeam.Tag = conditionalBeam
                                                                    userNodeProlet.Nodes.Add(userNodeBeam)
                                                                End If
                                                            End If
                                                        Next k2
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    Next j
                                    'заполняем ноды с элементами балок
                                    Dim arrayBeamsNodes As TreeNode() = TreeView1.Nodes.Find(StructureElement.typeObject.axisBeam, True)
                                    If arrayBeamsNodes.Length > 0 Then
                                        For j As Integer = 0 To dictionaryBridgeElements.Count - 1
                                            Dim keyTypeObject As StructureElement.typeObject = dictionaryBridgeElements.ElementAt(j).Key
                                            Dim listValue As List(Of StructureElement) = dictionaryBridgeElements.ElementAt(j).Value
                                            If keyTypeObject = typeObject.counterTopBeam Or keyTypeObject = typeObject.counterBottomBeam Or keyTypeObject = typeObject.modelBeam Then
                                                For k As Integer = 0 To listValue.Count - 1
                                                    Dim dataElement As StructureElement = listValue.Item(k)
                                                    Dim keyParamBeam As String = dataElement.KeyParameter
                                                    If FuncGSON.IsValidJson(keyParamBeam) = True Then
                                                        Dim numberProlet As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet"))
                                                        Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                        If numberProlet > 0 Then
                                                            For k2 As Integer = 0 To arrayBeamsNodes.Length - 1
                                                                Dim nodeAxisBeam As TreeNode = arrayBeamsNodes(k2)
                                                                Dim condElement As conditionalElement = nodeAxisBeam.Tag
                                                                If IsNothing(condElement.DataElement) = False Then
                                                                    If condElement.DataElement.IdStructure Like idBridge Then
                                                                        If condElement.NumberColl = numberProlet Then
                                                                            If condElement.NumberRow = numberRow Then
                                                                                Dim nameElement As String = StructureElement.GetDescription(keyTypeObject)
                                                                                Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                                nodeElement.Name = keyTypeObject
                                                                                Dim conditionalPillar As conditionalElement = New conditionalElement(numberProlet, dataElement)
                                                                                conditionalPillar.NumberColl = numberProlet
                                                                                conditionalPillar.NumberRow = numberRow
                                                                                nodeElement.Tag = conditionalPillar
                                                                                nodeElement.ForeColor = Color.Black
                                                                                nodeAxisBeam.Nodes.Add(nodeElement)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next k2
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                        Next j
                                    End If
                                End If
                            End If
                            '==============================================================================================================================
                            'заполняем Участки омоноличивания балок
                            '==============================================================================================================================
                            Dim dictSitesMonolitBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getSitesMonolitBeams(dictionaryBridgeElements)
                            If dictBeams.Count > 0 Then
                                Dim arrayBeamsGNodes As TreeNode() = TreeView1.Nodes.Find("MONOLIT_SITES", True)
                                If arrayBeamsGNodes.Length > 0 Then
                                    For j As Integer = 0 To arrayBeamsGNodes.Length - 1
                                        Dim nodeBeams As TreeNode = arrayBeamsGNodes(j)
                                        Dim condBeams As conditionalElement = nodeBeams.Tag
                                        If IsNothing(condBeams.DataElement) = False Then
                                            If condBeams.DataElement.IdStructure Like idBridge Then
                                                For k As Integer = 0 To dictBeams.Count - 1
                                                    Dim numberProlet As Integer = dictBeams.ElementAt(k).Key
                                                    Dim dataBeams As Dictionary(Of Integer, StructureElement) = dictBeams.ElementAt(k).Value
                                                    Dim conditionalBridgeElementDefault As conditionalElement = New conditionalElement(numberProlet, userDataDefault)
                                                    Dim userNodeProlet As TreeNode = New TreeNode("Пролет№ " & numberProlet, 12, 12)
                                                    userNodeProlet.Name = "PROLET_MonolitBEAMS"
                                                    userNodeProlet.Tag = conditionalBridgeElementDefault
                                                    nodeBeams.Nodes.Add(userNodeProlet)
                                                    If dataBeams.Count > 0 Then
                                                        For k2 As Integer = 0 To dataBeams.Count - 1
                                                            Dim dataBeam As StructureElement = dataBeams.ElementAt(k2).Value
                                                            If IsNothing(dataBeam) = False Then
                                                                If IsNothing(dataBeam.DWGEntity) = False Then
                                                                    Dim numberRow As Integer = dataBeams.ElementAt(k2).Key
                                                                    Dim strRow As String = FuncFormatZn.getConditionalRow(numberRow)
                                                                    Dim userNodeBeam As TreeNode = New TreeNode("Балка № " & strRow, 12, 12)
                                                                    Dim conditionalBeam As conditionalElement = New conditionalElement(numberRow, dataBeam)
                                                                    conditionalBeam.NumberColl = numberProlet
                                                                    conditionalBeam.NumberRow = numberRow
                                                                    userNodeBeam.Name = StructureElement.typeObject.axisBeam
                                                                    userNodeBeam.Tag = conditionalBeam
                                                                    userNodeProlet.Nodes.Add(userNodeBeam)
                                                                End If
                                                            End If
                                                        Next k2
                                                    End If
                                                Next k
                                            End If
                                        End If
                                    Next j
                                    'заполняем ноды с балками
                                    Dim arrayBeamsMonolitNodes As TreeNode() = TreeView1.Nodes.Find(StructureElement.typeObject.axisSiteMonolitBeams, True)
                                    If arrayBeamsMonolitNodes.Length > 0 Then
                                        For j As Integer = 0 To dictionaryBridgeElements.Count - 1
                                            Dim keyTypeObject As StructureElement.typeObject = dictionaryBridgeElements.ElementAt(j).Key
                                            Dim listValue As List(Of StructureElement) = dictionaryBridgeElements.ElementAt(j).Value
                                            If keyTypeObject = typeObject.counterSiteMonolitBeamsTop Or keyTypeObject = typeObject.counterSiteMonolitBeamsBottom Or keyTypeObject = typeObject.modelSiteMonolitBeams Then
                                                For k As Integer = 0 To listValue.Count - 1
                                                    Dim dataElement As StructureElement = listValue.Item(k)
                                                    Dim keyParamBeam As String = dataElement.KeyParameter
                                                    If FuncGSON.IsValidJson(keyParamBeam) = True Then
                                                        Dim numberProlet As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet"))
                                                        Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                        If numberProlet > 0 Then
                                                            For k2 As Integer = 0 To arrayBeamsMonolitNodes.Length - 1
                                                                Dim nodeAxisBeam As TreeNode = arrayBeamsMonolitNodes(k2)
                                                                Dim condElement As conditionalElement = nodeAxisBeam.Tag
                                                                If IsNothing(condElement.DataElement) = False Then
                                                                    If condElement.DataElement.IdStructure Like idBridge Then
                                                                        If condElement.NumberColl = numberProlet Then
                                                                            If condElement.NumberRow = numberRow Then
                                                                                Dim nameElement As String = StructureElement.GetDescription(keyTypeObject)
                                                                                Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                                nodeElement.Name = keyTypeObject
                                                                                Dim conditionalPillar As conditionalElement = New conditionalElement(numberProlet, dataElement)
                                                                                conditionalPillar.NumberColl = numberProlet
                                                                                conditionalPillar.NumberRow = numberRow
                                                                                nodeElement.Tag = conditionalPillar
                                                                                nodeElement.ForeColor = Color.Black
                                                                                nodeAxisBeam.Nodes.Add(nodeElement)
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next k2
                                                        End If
                                                    End If
                                                Next k
                                            End If
                                        Next j
                                    End If
                                End If
                            End If
                            '==============================================================================================================================
                            'заполняем опоры
                            '==============================================================================================================================
                            'получаем ноды для всх осей опор
                            Dim arrayPillarNodes As TreeNode() = TreeView1.Nodes.Find(StructureElement.typeObject.axisPillar, True)
                            Dim arraySubFermNodes As TreeNode() = {}
                            Dim arrayRackNodes As TreeNode() = {}
                            Dim arrayPilesNodes As TreeNode() = {}
                            If arrayPillarNodes.Length > 0 Then
                                For j As Integer = 0 To arrayPillarNodes.Length - 1
                                    Dim nodePillar As TreeNode = arrayPillarNodes(j)
                                    Dim nodeTag As conditionalElement = nodePillar.Tag
                                    If IsNothing(nodeTag.DataElement) = False Then
                                        If nodeTag.DataElement.IdStructure Like idBridge Then
                                            Dim nodeNumberPillar As Integer = nodeTag.NumberElement
                                            Dim nodeTypeObject As StructureElement.typeObject = DirectCast([Enum].Parse(GetType(StructureElement.typeObject), nodePillar.Name), StructureElement.typeObject)
                                            For k As Integer = 0 To dictionaryBridgeElements.Count - 1
                                                Dim keyTypeObject As StructureElement.typeObject = dictionaryBridgeElements.ElementAt(k).Key
                                                Dim listValue As List(Of StructureElement) = dictionaryBridgeElements.ElementAt(k).Value
                                                If keyTypeObject = typeObject.axisSubFermenters Then
                                                    Dim tempDict As Dictionary(Of Integer, StructureElement) = SubFermenters.sortListSubFermenter(listValue, nodeNumberPillar)
                                                    If tempDict.Count > 0 Then
                                                        listValue = tempDict.Values.ToList
                                                    End If
                                                ElseIf keyTypeObject = typeObject.axisRack Then
                                                    Dim tempDict As Dictionary(Of Integer, StructureElement) = RackPillar.sortListRack(listValue, nodeNumberPillar)
                                                    If tempDict.Count > 0 Then
                                                        listValue = tempDict.Values.ToList
                                                    End If
                                                ElseIf keyTypeObject = typeObject.axisPile Then
                                                    Dim tempDict As Dictionary(Of Integer, StructureElement) = PilePillar.sortListPiles(listValue, nodeNumberPillar)
                                                    If tempDict.Count > 0 Then
                                                        listValue = tempDict.Values.ToList
                                                    End If
                                                End If
                                                If listValue.Count > 0 Then
                                                    For k1 As Integer = 0 To listValue.Count - 1
                                                        Dim dataElement As StructureElement = listValue.Item(k1)
                                                        Dim classObject As StructureElement.classStructure = dataElement.ClassObject
                                                        Dim typeObject As StructureElement.typeObject = dataElement.Name
                                                        'обрабатываем оси опор
                                                        If typeObject = typeObject.axisPillar Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "Number"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                nodePillar.Tag = conditionalPillar
                                                                nodePillar.ForeColor = Color.Black
                                                            End If
                                                            '============================================================================================================
                                                            'ось подферменников
                                                        ElseIf typeObject = typeObject.axisSubFermenters Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim numberProlet As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet"))
                                                                Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                                Dim nodeGSubFerm As TreeNode = FindChildNodeByName(nodePillar, classStructure.GroupSubFermenters)
                                                                If IsNothing(nodeGSubFerm) = False Then
                                                                    Dim nodeElement As TreeNode = New TreeNode("Подферменник №" & FuncFormatZn.getConditionalRow(numberRow), 9, 9)
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    conditionalPillar.NumberRow = numberRow
                                                                    conditionalPillar.NumberColl = numberProlet
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeElement.Name = typeObject
                                                                    nodeGSubFerm.Nodes.Add(nodeElement)
                                                                    nodeGSubFerm.ForeColor = Color.Black
                                                                End If
                                                            End If
                                                            'оси стоек
                                                        ElseIf typeObject = typeObject.axisRack Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim numberRack As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "Number"))
                                                                Dim nodeGRack As TreeNode = FindChildNodeByName(nodePillar, classStructure.GroupRackPillar)
                                                                If IsNothing(nodeGRack) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode("Стойка №" & numberRack, 9, 9)
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    conditionalPillar.NumberRow = numberRack
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeElement.Name = typeObject
                                                                    nodeGRack.Nodes.Add(nodeElement)
                                                                    nodeGRack.ForeColor = Color.Black
                                                                End If
                                                            End If
                                                            'оси свай
                                                        ElseIf typeObject = typeObject.axisPile Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                                Dim numberColl As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberColumn"))
                                                                Dim nodeGPile As TreeNode = FindChildNodeByName(nodePillar, classStructure.GroupPilePillar)
                                                                If IsNothing(nodeGPile) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode("Свая " & "(Ряд №" & numberRow & ", столбец № " & numberColl & ")", 9, 9)
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    conditionalPillar.NumberRow = numberRow
                                                                    conditionalPillar.NumberColl = numberColl
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeGPile.ForeColor = Color.Black
                                                                    nodeElement.Name = typeObject
                                                                    nodeGPile.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.axisNozzle Or typeObject = typeObject.axisRigel Or typeObject = typeObject.axisCabinetWall Or typeObject = typeObject.axisGrillage Or typeObject = typeObject.axisPreparation Or typeObject = typeObject.axisLeftHand Or typeObject = typeObject.axisRightHand Or typeObject = typeObject.axisLeftPostcard Or typeObject = typeObject.axisRightPostcard Then
                                                            'остальные элементы
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodePillarElement As TreeNode = FindChildNodeByName(nodePillar, typeObject)
                                                                If IsNothing(nodePillarElement) = False Then
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodePillarElement.Tag = conditionalPillar
                                                                    nodePillarElement.ForeColor = Color.Black
                                                                End If
                                                            End If
                                                            'вспомогательные элементы
                                                        ElseIf typeObject = typeObject.contourNozzleTop Or typeObject = typeObject.contourNozzleBottom Or typeObject = typeObject.modelNozzle Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeNozzle As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisNozzle)
                                                                If IsNothing(nodeNozzle) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeNozzle.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterRigelTop Or typeObject = typeObject.counterRigelBottom Or typeObject = typeObject.modelRigel Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeNozzle As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisRigel)
                                                                If IsNothing(nodeNozzle) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeNozzle.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterGrillageTop Or typeObject = typeObject.counterGrillageBottom Or typeObject = typeObject.modelGrillage Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeNozzle As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisGrillage)
                                                                If IsNothing(nodeNozzle) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeNozzle.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterPreparationTop Or typeObject = typeObject.counterPreparationBottom Or typeObject = typeObject.modelPreparation Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeNozzle As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisPreparation)
                                                                If IsNothing(nodeNozzle) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeNozzle.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.contourCabinetWallBottom Or typeObject = typeObject.contourCabinetWallPlateBottom Or typeObject = typeObject.contourCabinetWallPlateTop Or typeObject = typeObject.contourCabinetWallTop Or typeObject = typeObject.modelCabinetWall Or typeObject = typeObject.modelCabinetWallPlate Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeNozzle As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisCabinetWall)
                                                                If IsNothing(nodeNozzle) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeNozzle.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterLeftHandTop Or typeObject = typeObject.counterLeftHandBottom Or typeObject = typeObject.counterLeftHandCorniceTop Or typeObject = typeObject.counterLeftHandCorniceBottom Or typeObject = typeObject.modelLeftHandCornice Or typeObject = typeObject.modelLeftHand Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeLeftHand As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisLeftHand)
                                                                If IsNothing(nodeLeftHand) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeLeftHand.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterRightHandTop Or typeObject = typeObject.counterRightHandBottom Or typeObject = typeObject.counterRightHandCorniceTop Or typeObject = typeObject.counterRightHandCorniceBottom Or typeObject = typeObject.modelRightHandCornice Or typeObject = typeObject.modelRightHand Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeRightHand As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisRightHand)
                                                                If IsNothing(nodeRightHand) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeRightHand.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterLeftPostcardTop Or typeObject = typeObject.counterLeftPostcardBottom Or typeObject = typeObject.modelLeftPostcard Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeLeftPostcard As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisLeftPostcard)
                                                                If IsNothing(nodeLeftPostcard) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeLeftPostcard.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterRightPostcardTop Or typeObject = typeObject.counterRightPostcardBottom Or typeObject = typeObject.modelRightPostcard Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            If numberPillar = nodeNumberPillar Then
                                                                Dim nodeRightPostcard As TreeNode = FindChildNodeByName(nodePillar, typeObject.axisRightPostcard)
                                                                If IsNothing(nodeRightPostcard) = False Then
                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                    nodeElement.Name = typeObject
                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                    nodeElement.Tag = conditionalPillar
                                                                    nodeElement.ForeColor = Color.Black
                                                                    nodeRightPostcard.Nodes.Add(nodeElement)
                                                                End If
                                                            End If
                                                        ElseIf typeObject = typeObject.counterSubFermentersTop Or typeObject = typeObject.counterSubFermentersBottom Or typeObject = typeObject.modelSubFermenters Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            Dim numberProlet As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet"))
                                                            If numberPillar = 1 And numberProlet = 0 Then
                                                                numberProlet = 1
                                                            End If
                                                            Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                            If arraySubFermNodes.Length = 0 Then
                                                                arraySubFermNodes = TreeView1.Nodes.Find(StructureElement.typeObject.axisSubFermenters, True)
                                                            End If
                                                            If arraySubFermNodes.Length > 0 Then
                                                                For k2 As Integer = 0 To arraySubFermNodes.Length - 1
                                                                    Dim nodeAxisSubFerm As TreeNode = arraySubFermNodes(k2)
                                                                    Dim condElement As conditionalElement = nodeAxisSubFerm.Tag
                                                                    If IsNothing(nodeTag.DataElement) = False Then
                                                                        If nodeTag.DataElement.IdStructure Like idBridge Then
                                                                            If condElement.NumberElement = nodeNumberPillar Then
                                                                                If condElement.NumberRow = numberRow Then
                                                                                    If condElement.NumberColl = numberProlet Then
                                                                                        Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                                        Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                                        nodeElement.Name = typeObject
                                                                                        Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                                        conditionalPillar.NumberColl = numberProlet
                                                                                        conditionalPillar.NumberRow = numberRow
                                                                                        nodeElement.Tag = conditionalPillar
                                                                                        nodeElement.ForeColor = Color.Black
                                                                                        nodeAxisSubFerm.Nodes.Add(nodeElement)
                                                                                        Exit For
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        ElseIf typeObject = typeObject.counterRackTop Or typeObject = typeObject.counterRackBottom Or typeObject = typeObject.modelRack Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            Dim numberRack As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRack"))
                                                            If arrayRackNodes.Length = 0 Then
                                                                arrayRackNodes = TreeView1.Nodes.Find(StructureElement.typeObject.axisRack, True)
                                                            End If
                                                            If arrayRackNodes.Length > 0 Then
                                                                For k2 As Integer = 0 To arrayRackNodes.Length - 1
                                                                    Dim nodeAxisRack As TreeNode = arrayRackNodes(k2)
                                                                    Dim condElement As conditionalElement = nodeAxisRack.Tag
                                                                    If nodeTag.DataElement.IdStructure Like idBridge Then
                                                                        If condElement.NumberElement = nodeNumberPillar Then
                                                                            If condElement.NumberRow = numberRack Then
                                                                                Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                                Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                                nodeElement.Name = typeObject
                                                                                Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                                conditionalPillar.NumberRow = numberRack
                                                                                nodeElement.Tag = conditionalPillar
                                                                                nodeElement.ForeColor = Color.Black
                                                                                nodeAxisRack.Nodes.Add(nodeElement)
                                                                                nodeAxisRack.Parent.ForeColor = Color.Black
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        ElseIf typeObject = typeObject.counterPileTop Or typeObject = typeObject.counterPileBottom Or typeObject = typeObject.modelPile Then
                                                            Dim numberPillar As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar"))
                                                            Dim numberColl As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberColumn"))
                                                            Dim numberRow As Integer = Val(FuncGSON.getValue(dataElement.KeyParameter, "NumberRow"))
                                                            If arrayPilesNodes.Length = 0 Then
                                                                arrayPilesNodes = TreeView1.Nodes.Find(StructureElement.typeObject.axisPile, True)
                                                            End If
                                                            If arrayPilesNodes.Length > 0 Then
                                                                For k2 As Integer = 0 To arrayPilesNodes.Length - 1
                                                                    Dim nodeAxisPile As TreeNode = arrayPilesNodes(k2)
                                                                    Dim condElement As conditionalElement = nodeAxisPile.Tag
                                                                    If nodeTag.DataElement.IdStructure Like idBridge Then
                                                                        If condElement.NumberElement = nodeNumberPillar Then
                                                                            If condElement.NumberRow = numberRow Then
                                                                                If condElement.NumberColl = numberColl Then
                                                                                    Dim nameElement As String = StructureElement.GetDescription(typeObject)
                                                                                    Dim nodeElement As TreeNode = New TreeNode(nameElement, 9, 9)
                                                                                    nodeElement.Name = typeObject
                                                                                    Dim conditionalPillar As conditionalElement = New conditionalElement(numberPillar, dataElement)
                                                                                    conditionalPillar.NumberColl = numberColl
                                                                                    conditionalPillar.NumberRow = numberRow
                                                                                    nodeElement.Tag = conditionalPillar
                                                                                    nodeElement.ForeColor = Color.Black
                                                                                    nodeAxisPile.Nodes.Add(nodeElement)
                                                                                    Exit For
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next k2
                                                            End If
                                                        End If
                                                    Next k1
                                                End If
                                            Next k
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор узла дерева
    Public Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Dim selectNode As TreeNode = e.Node
        If IsNothing(selectNode) = False Then
            'ищем корневой элент
            Dim generalNode As TreeNode = selectNode
            Dim nameNode As String = e.Node.Name
            'ищем корневой элемент
            If IsNothing(selectNode) = False Then
                Dim levelSelectNode As Integer = selectNode.Level
                If levelSelectNode > 0 Then
                    For i As Integer = levelSelectNode - 1 To 0 Step -1
                        generalNode = generalNode.Parent
                    Next
                End If
            End If
            If IsNothing(generalNode) = True Then
                Exit Sub
            End If

            Dim conditionalNodeTag As conditionalElement = e.Node.Tag
            If IsNothing(conditionalNodeTag) = True Then
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'свойства моста
            If IsNothing(conditionalNodeTag.DataElement) = True Then
                Exit Sub
            End If
            Dim idBridge As String = conditionalNodeTag.DataElement.IdStructure
            Dim dataBridge As StructureElement = conditionalNodeTag.DataElement
            Dim userBridge As Bridges = dataBridge.getBridge
            If IsNothing(userBridge) = True Then
                Exit Sub
            End If
            Dim typeSelectObject As StructureElement.typeObject = StructureElement.ConvertToTypeObject(nameNode)
            If typeSelectObject <> typeObject.OtherElement Then
                Dim nameAlign As String = userBridge.AlignmentName
                Dim nameSurface As String = userBridge.projectSurfaceName
                Dim nameEgSurface As String = userBridge.EarthSurfaceName
                If IsNothing(nameSurface) = True Then
                    nameSurface = ""
                End If
                If IsNothing(nameEgSurface) = True Then
                    nameEgSurface = ""
                End If
                If IsNothing(nameAlign) = True Then
                    nameAlign = ""
                End If
                If nameAlign.Trim.Length > 0 Then
                    Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlignment)
                    If boolFindAlign = True Then
                        projectAlignment.Plan.CompoundLine.ToPolyLine(projectPolylineAlignment)
                    End If
                End If
                If nameSurface.Trim.Length = 0 Then
                    Dim boolFindAlign = FuncAlignment.getAlignmentByName(nameAlign, projectAlignment)
                    projectSurface = FuncSurface.getSurfaceByName(nameSurface)
                Else
                    Dim boolFindAlign = FuncAlignment.getAlignmentByName(nameAlign, projectAlignment)
                    projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
                End If
                If nameEgSurface.Trim.Length > 0 Then
                    egSurface = FuncSurface.getSurfaceByName(nameEgSurface)
                End If
                'запоминаем выбор
                Dim arrayStr As String() = {selectNode.Name, selectNode.FullPath, selectNode.Index, idBridge}
                PropertyGrid1.Tag = arrayStr 'запоминаем имя нода и патчь
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбор объекта
                ObjectFactory.LoadToPropertyGrid(typeSelectObject, dataBridge.KeyParameter, PropertyGrid1)
                ' Или если нужно получить объект
                Dim myObject As Object = ObjectFactory.CreateObject(typeSelectObject, dataBridge.KeyParameter)
                If myObject IsNot Nothing Then
                    PropertyGrid1.SelectedObject = myObject
                    PropertyGrid1.Tag = selectNode.Tag
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ввод значения в datagrid
    Private Sub PropertyGrid1_PropertyValueChanged(s As Object, e As PropertyValueChangedEventArgs) Handles PropertyGrid1.PropertyValueChanged
        Dim nameField As String = e.ChangedItem.PropertyDescriptor.Name
        If e.ChangedItem.PropertyDescriptor.IsReadOnly = False Then
            Dim newValue As String = e.ChangedItem.Value
            Dim oldValue As String = e.OldValue
            Dim condElement As conditionalElement = PropertyGrid1.Tag
            If IsNothing(condElement) = False Then
                Dim dataElement As StructureElement = condElement.DataElement
                If IsNothing(dataElement) = False Then
                    Dim keyParameter As String = dataElement.KeyParameter
                    Dim idBridge As String = dataElement.IdStructure
                    If FuncGSON.IsValidJson(keyParameter) = True Then
                        Dim boolRecordData As Boolean = FuncGSON.setValue(keyParameter, nameField, newValue)
                        If boolRecordData = True Then
                            dataElement.KeyParameter = keyParameter
                            Dim entity As DwgEntity = dataElement.DWGEntity
                            Dim boolRecData As Boolean = FuncXRecords.setXRecords(entity, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataElement)
                            If e.ChangedItem.Parent.Label Like "CivilEnginStructures.NozzlePillar" Then
                                Dim userNozzle As NozzlePillar = dataElement.getNozzlePillar()
                                If IsNothing(userNozzle) = False Then
                                    Dim boolCalculateNozzle As Boolean = userNozzle.calculateNozzle(Nothing, projectPolylineAlignment)
                                    Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userNozzle._elementBridgePoint.ListPointModel
                                    If ListPointModel.Count > 3 Then
                                        'ось насадки
                                        Dim dataNozzle As StructureElement = userNozzle.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = NozzleContour.drawContoursNozzle(ActivDocumentPanel, userNozzle, idBridge, dictionaryBridgeElements, templateXML)
                                        'рисуем модель
                                        Dim topElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.contourNozzleTop)
                                        Dim BottomElement As DwgPolyline3D = dictionaryCounter.Item(typeObject.contourNozzleBottom)
                                        Dim middlePt1 As Vector3D = userNozzle.getPointByCode("middlePt1", True)
                                        Dim middlePt2 As Vector3D = userNozzle.getPointByCode("middlePt2", True)
                                        Dim tempLineCabinetWall As DwgLine = New DwgLine
                                        tempLineCabinetWall.StartPoint = middlePt1
                                        tempLineCabinetWall.EndPoint = middlePt2
                                        Dim model As DwgModel3DElement = NozzleModel.drawModelNozzle(ActivDocumentPanel, topElement, BottomElement, tempLineCabinetWall, dataNozzle, dictionaryBridgeElements, templateXML)
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.SubFermenters" Then
                                Dim userSubFerm As SubFermenters = dataElement.getSubFermenters()
                                If IsNothing(userSubFerm) = False Then
                                    Dim boolCalculateNozzle As Boolean = SubFermenters.calculateSubFermenter(userSubFerm, dictionaryBridgeElements)
                                    Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userSubFerm._elementBridgePoint.ListPointModel
                                    If ListPointModel.Count > 3 Then
                                        'ось насадки
                                        Dim dataSubFerm As StructureElement = userSubFerm.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = SubFermenterContour.drawContour(ActivDocumentPanel, userSubFerm, idBridge, dictionaryBridgeElements, templateXML)
                                        Dim boolDrawModel As Boolean = SubFermenterModel.drawModel(ActivDocumentPanel, userSubFerm, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.RigelPillar" Then
                                Dim userRigel As RigelPillar = dataElement.getRigelPillar()
                                If IsNothing(userRigel) = False Then
                                    Dim listBeamsPillar As List(Of Dictionary(Of Integer, StructureElement)) = Pillar.getBeamsPillarByNumber(userRigel.NumberPillar, dictionaryBridgeElements)
                                    If listBeamsPillar.Count > 1 Then
                                        Dim boolCalculateRigel As Boolean = userRigel.calculateRigel(Nothing, listBeamsPillar)
                                        Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userRigel._elementBridgePoint.ListPointModel
                                        If ListPointModel.Count > 3 Then
                                            'ось насадки
                                            Dim dataRigel As StructureElement = userRigel.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                            'рисуем контура
                                            Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = RigelContour.drawContour(ActivDocumentPanel, userRigel, idBridge, dictionaryBridgeElements, templateXML)
                                            Dim boolDrawModel As Boolean = RigelModel.drawModel(ActivDocumentPanel, userRigel, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                        End If
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.CabinetWallPillar" Then
                                Dim userCabinetWall As CabinetWallPillar = dataElement.getCabinetWallPillar()
                                If IsNothing(userCabinetWall) = False Then
                                    Dim listDataLeftHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, userCabinetWall.NumberPillar, Pillar.SidePillarElement.Left)
                                    Dim dataLeftHand As StructureElement = StructureElement.isValidateDataStructure(listDataLeftHand)
                                    Dim userLeftHand As HandPillar = Nothing
                                    If IsNothing(dataLeftHand) = False Then
                                        userLeftHand = dataLeftHand.getHandPillar
                                    End If
                                    Dim listDataRightHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, userCabinetWall.NumberPillar, Pillar.SidePillarElement.Right)
                                    Dim dataRightHand As StructureElement = StructureElement.isValidateDataStructure(listDataRightHand)
                                    Dim userRightHand As HandPillar = Nothing
                                    If IsNothing(dataRightHand) = False Then
                                        userRightHand = dataRightHand.getHandPillar
                                    End If
                                    Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, userCabinetWall.NumberPillar)
                                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                    Dim userNozzle As NozzlePillar = Nothing
                                    If IsNothing(dataNozzle) = False Then
                                        userNozzle = dataNozzle.getNozzlePillar
                                    End If
                                    'делаем расчет
                                    If IsNothing(projectSurface) = True Then
                                        projectSurface = FuncSurface.getSurfaceByName(userBridge.projectSurfaceName)
                                    End If
                                    If IsNothing(projectSurface) = True Then
                                        projectSurface = FuncAlignment.getSurfaceToAlignment(userBridge.AlignmentName)
                                    End If
                                    Dim boolCalculateCabinetWall As Boolean = userCabinetWall.calculateCabinetWall(userNozzle, userLeftHand, userRightHand, projectSurface, projectAlignment)
                                    Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userCabinetWall._elementBridgePoint.ListPointModel
                                    If ListPointModel.Count > 3 Then
                                        'ось насадки
                                        Dim dataRigel As StructureElement = userCabinetWall.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = CabinetWallContour.drawContours(ActivDocumentPanel, userCabinetWall, idBridge, dictionaryBridgeElements, templateXML)
                                        Dim boolDrawModel As Boolean = CabinetWallModel.drawModel(ActivDocumentPanel, userCabinetWall, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.HandPillar" Then
                                Dim userHand As HandPillar = dataElement.getHandPillar()
                                If IsNothing(userHand) = False Then
                                    Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, userHand.NumberPillar)
                                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                    Dim userNozzle As NozzlePillar = Nothing
                                    If IsNothing(dataNozzle) = False Then
                                        userNozzle = dataNozzle.getNozzlePillar
                                    End If
                                    Dim listDataCabinetWall As List(Of StructureElement) = CabinetWallPillar.getAxis(dictionaryBridgeElements, userHand.NumberPillar)
                                    Dim dataCabinetWall As StructureElement = StructureElement.isValidateDataStructure(listDataCabinetWall)
                                    Dim userCabinetWall As CabinetWallPillar = Nothing
                                    If IsNothing(dataCabinetWall) = False Then
                                        userCabinetWall = dataCabinetWall.getCabinetWallPillar
                                    End If
                                    Dim boolCalculateHand As Boolean = userHand.calculateHand(userNozzle, userCabinetWall, projectSurface)
                                    'ось 
                                    Dim dataHand As StructureElement = userHand.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                    'рисуем контура
                                    Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = HandContour.drawContours(ActivDocumentPanel, userHand, idBridge, dictionaryBridgeElements, templateXML)
                                    Dim boolDrawModel As Boolean = HandModel.drawModel(ActivDocumentPanel, userHand, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.PostcardPillar" Then
                                Dim userPostcard As PostcardPillar = dataElement.getPostcardPillar
                                If IsNothing(userPostcard) = False Then
                                    Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, userPostcard.NumberPillar)
                                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                    Dim userNozzle As NozzlePillar = Nothing
                                    If IsNothing(dataNozzle) = False Then
                                        userNozzle = dataNozzle.getNozzlePillar
                                    End If
                                    Dim boolCalculatePostcard As Boolean = userPostcard.calculatePostcard(userNozzle)
                                    'ось 
                                    Dim dataPostcard As StructureElement = userPostcard.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                    'рисуем контура
                                    Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PostcardContour.drawContour(ActivDocumentPanel, userPostcard, idBridge, dictionaryBridgeElements, templateXML)
                                    Dim boolDrawModel As Boolean = PostcardModel.drawModel(ActivDocumentPanel, userPostcard, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.GrillagePillar" Then
                                Dim userGrillage As GrillagePillar = dataElement.getGrillagePillar
                                If IsNothing(userGrillage) = False Then
                                    Dim dataPillar As StructureElement = Pillar.getAxisPillar(dictionaryBridgeElements, userGrillage.NumberPillar)
                                    Dim userPillar As Pillar = Nothing
                                    Dim axisPillar As DwgLine = New DwgLine
                                    Dim axisPlineAlign As Polyline3D = New Polyline3D
                                    If IsNothing(dataPillar) = False Then
                                        userPillar = dataPillar.getPillar
                                        axisPillar = dataPillar.DWGEntity
                                    End If
                                    If IsNothing(projectAlignment) = False Then
                                        projectAlignment.Plan.CompoundLine.ToPolyLine(axisPlineAlign)
                                    End If
                                    Dim arrayRack As RackPillar() = RackPillar.getRackPillarToArray(dictionaryBridgeElements, userGrillage.NumberPillar)
                                    If arrayRack.Length > 0 Then
                                        Dim boolCalculateGrillage As Boolean = userGrillage.calculateGrillage(axisPillar, arrayRack, axisPlineAlign)
                                        'ось 
                                        Dim dataPostcard As StructureElement = userGrillage.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = GrillageContour.drawContour(ActivDocumentPanel, userGrillage, idBridge, dictionaryBridgeElements, templateXML)
                                        Dim boolDrawModel As Boolean = GrillageModel.drawModel(ActivDocumentPanel, userGrillage, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.PreparationPillar" Then
                                Dim userPreparation As PreparationPillar = dataElement.getPreparationPillar
                                If IsNothing(userPreparation) = False Then
                                    Dim dataGrillage As StructureElement = GrillagePillar.getGrillagePillar(dictionaryBridgeElements, userPreparation.NumberPillar)
                                    Dim userGrillage As GrillagePillar = Nothing
                                    If IsNothing(dataGrillage) = False Then
                                        userGrillage = dataGrillage.getGrillagePillar
                                    End If
                                    Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, userPreparation.NumberPillar)
                                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                    Dim userNozzle As NozzlePillar = Nothing
                                    If IsNothing(dataNozzle) = False Then
                                        userNozzle = dataNozzle.getNozzlePillar
                                    End If
                                    Dim axisPlineAlign As Polyline3D = New Polyline3D
                                    If IsNothing(projectAlignment) = False Then
                                        projectAlignment.Plan.CompoundLine.ToPolyLine(axisPlineAlign)
                                    End If
                                    If IsNothing(userGrillage) = False Then
                                        Dim boolCalculatePreparation As Boolean = userPreparation.calculatePreparation(userGrillage, userNozzle, axisPlineAlign)
                                        'ось 
                                        Dim dataPreparation As StructureElement = userPreparation.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgPolyline3D) = PreparationContour.drawContour(ActivDocumentPanel, userPreparation, idBridge, dictionaryBridgeElements, templateXML)
                                        Dim boolDrawModel As Boolean = PreparationModel.drawModel(ActivDocumentPanel, userPreparation, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML)
                                    End If
                                End If
                            ElseIf e.ChangedItem.Parent.Label Like "CivilEnginStructures.RackPillar" Then
                                Dim userRack As RackPillar = dataElement.getRackPillar
                                If IsNothing(userRack) = False Then
                                    Dim dataGrillage As StructureElement = GrillagePillar.getGrillagePillar(dictionaryBridgeElements, userRack.NumberPillar)
                                    Dim userGrillage As GrillagePillar = Nothing
                                    If IsNothing(dataGrillage) = False Then
                                        userGrillage = dataGrillage.getGrillagePillar
                                    End If
                                    Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, userRack.NumberPillar)
                                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                    Dim userNozzle As NozzlePillar = Nothing
                                    If IsNothing(dataNozzle) = False Then
                                        userNozzle = dataNozzle.getNozzlePillar
                                    End If
                                    Dim listDataRigel As List(Of StructureElement) = RigelPillar.getAxis(dictionaryBridgeElements, userRack.NumberPillar)
                                    Dim dataRigel As StructureElement = StructureElement.isValidateDataStructure(listDataRigel)
                                    Dim userRigel As RigelPillar = Nothing
                                    If IsNothing(dataRigel) = False Then
                                        userRigel = dataRigel.getRigelPillar
                                    End If
                                    Dim axisPlineAlign As Polyline3D = New Polyline3D
                                    If IsNothing(projectAlignment) = False Then
                                        projectAlignment.Plan.CompoundLine.ToPolyLine(axisPlineAlign)
                                    End If
                                    If IsNothing(userGrillage) = False Then
                                        Dim boolCalculateRack As Boolean = userRack.calculateRack(userNozzle, userRigel, userGrillage)
                                        'ось 
                                        Dim dataRack As StructureElement = userRack.drawAxis(ActivDocumentPanel, idBridge, templateXML, dictionaryBridgeElements)
                                        'рисуем контура
                                        Dim dictionaryCounter As Dictionary(Of StructureElement.typeObject, DwgEntity) = RackContour.drawContour(ActivDocumentPanel, userRack, idBridge, dictionaryBridgeElements, templateXML)
                                        Dim boolDrawModel As Boolean = RackModel.drawModel(ActivDocumentPanel, userRack, dictionaryCounter, idBridge, dictionaryBridgeElements, templateXML, False)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ВЫБОР ОБЪЕКТОВ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'показать элемент с выбором
    Private Sub ПоказатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказатьToolStripMenuItem.Click
        Dim selectedNode As TreeNode = TreeView1.SelectedNode
        If IsNothing(selectedNode) = True Then
            MsgBox("Выберите элемент в списке!!!")
            Exit Sub
        End If
        Dim tagSelect As Object = selectedNode.Tag
        If TypeOf tagSelect Is conditionalElement Then
            Dim condElement As conditionalElement = tagSelect
            Dim dataElement As StructureElement = condElement.DataElement
            If IsNothing(dataElement) = False Then
                Dim selectedEntity As DwgEntity = dataElement.DWGEntity
                If IsNothing(selectedEntity) = False Then
                    Dim userSelectedBounds As BoundingBox2D = selectedEntity.Bounds
                    If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                        CadViewPanel.ZoomBound(userSelectedBounds, True)
                    ElseIf ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.VolumeWindow Then
                        Dim acModel As DwgModel3DElement = Nothing
                        If TypeOf selectedEntity Is DwgModel3DElement Then
                            acModel = selectedEntity
                        Else
                            MsgBox("Выберите модель!")
                            Exit Sub
                        End If
                        If IsNothing(acModel) = False Then
                            Dim wnd = TryCast(ApplicationHost.Current.ActiveProject(Consts.VolumeWindow), IFramableDocumentWindow)
                            If wnd IsNot Nothing Then
                                Dim panel = TryCast(wnd.ActiveFrame.Control.Controls(0), Panel3d)
                                If panel IsNot Nothing Then
                                    Try
                                        Dim bim_camera As Panel3d.BimCamera = CType(panel.Camera, Panel3d.BimCamera)
                                        bim_camera.SolveLimits(acModel.Bounds3d)
                                    Catch ex As System.ArgumentOutOfRangeException
                                    End Try
                                End If
                            End If
                        End If
                    ElseIf IsNumeric(ApplicationHost.Current.ActiveProject.ActiveDocument.UID) Then
                        Dim desk As String = ApplicationHost.Current.ActiveProject.ActiveDocument.Text
                        Dim number As Integer = Val(ApplicationHost.Current.ActiveProject.ActiveDocument.UID)
                        If desk.IndexOf("Редактор опоры") > -1 Then
                            If TypeOf selectedEntity Is DwgModel3DElement Then
                                SMB.RaiseHighlightSectionModel(number, selectedEntity.ObjectID)
                            Else
                                MsgBox("Выберите модель!")
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub ВыбратьЭлементToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьЭлементToolStripMenuItem.Click
        Dim selectedNode As TreeNode = TreeView1.SelectedNode
        If IsNothing(selectedNode) = True Then
            MsgBox("Выберите элемент в списке!!!")
            Exit Sub
        End If
        Dim tagSelect As Object = selectedNode.Tag
        If TypeOf tagSelect Is conditionalElement Then
            Dim condElement As conditionalElement = tagSelect
            Dim dataElement As StructureElement = condElement.DataElement
            If IsNothing(dataElement) = False Then
                Dim selectedEntity As DwgEntity = dataElement.DWGEntity
                selectIDObject = selectedEntity.ObjectID
                If IsNothing(selectedEntity) = False Then
                    Dim userSelectedBounds As BoundingBox2D = selectedEntity.Bounds
                    If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                        CadViewPanel.ZoomBound(userSelectedBounds, True)
                        Try
                            CadViewPanel.SelectionSet.Clear()
                            CadViewPanel.SelectionSet.Select(selectedEntity, True)
                        Catch ex As System.ArgumentOutOfRangeException
                        End Try
                    ElseIf ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.VolumeWindow Then
                        Dim acModel As DwgModel3DElement = Nothing
                        If TypeOf selectedEntity Is DwgModel3DElement Then
                            acModel = selectedEntity
                        Else
                            MsgBox("Выберите модель!")
                            Exit Sub
                        End If
                        If IsNothing(acModel) = False Then
                            Dim wnd = TryCast(ApplicationHost.Current.ActiveProject(Consts.VolumeWindow), IFramableDocumentWindow)
                            If wnd IsNot Nothing Then
                                Dim panel = TryCast(wnd.ActiveFrame.Control.Controls(0), Panel3d)
                                If panel IsNot Nothing Then
                                    Try
                                        CadViewPanel.SelectionSet.Clear()
                                        CadViewPanel.SelectionSet.Select(acModel, True)
                                        Dim bim_camera As Panel3d.BimCamera = CType(panel.Camera, Panel3d.BimCamera)
                                        bim_camera.SolveLimits(acModel.Bounds3d)
                                    Catch ex As System.ArgumentOutOfRangeException
                                    End Try
                                End If
                            End If
                        End If
                    ElseIf IsNumeric(ApplicationHost.Current.ActiveProject.ActiveDocument.UID) Then
                        Dim desk As String = ApplicationHost.Current.ActiveProject.ActiveDocument.Text
                        Dim number As Integer = Val(ApplicationHost.Current.ActiveProject.ActiveDocument.UID)
                        If desk.IndexOf("Редактор опоры") > -1 Then
                            If TypeOf selectedEntity Is DwgModel3DElement Then
                                SMB.RaiseHighlightSectionModel(number, selectedEntity.ObjectID)
                            Else
                                MsgBox("Выберите модель!")
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub ПоказатьВсеНаПоперечникеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказатьВсеНаПоперечникеToolStripMenuItem.Click
        Dim selectedNode As TreeNode = TreeView1.SelectedNode
        If IsNothing(selectedNode) = True Then
            MsgBox("Выберите элемент в списке!!!")
            Exit Sub
        End If
        Dim generalNode As TreeNode = selectedNode
        'ищем корневой элемент
        If IsNothing(selectedNode) = False Then
            Dim levelSelectNode As Integer = selectedNode.Level
            If levelSelectNode > 0 Then
                For i As Integer = levelSelectNode - 1 To 0 Step -1
                    generalNode = generalNode.Parent
                Next
            End If
        End If
        'выбираем данные
        Dim userBridge As Bridges = Nothing
        Dim axisLineBridge As DwgEntity = Nothing
        Dim condAxisBridge As conditionalElement = generalNode.Tag
        If IsNothing(condAxisBridge) = False = True Then
            Dim dataBridge As StructureElement = condAxisBridge.DataElement
            If IsNothing(dataBridge) = False Then
                userBridge = dataBridge.getBridge()
                axisLineBridge = dataBridge.DWGEntity
                If IsNothing(userBridge) = False Then
                    Dim nameAlign As String = userBridge.AlignmentName
                    Dim nameSurface As String = userBridge.projectSurfaceName
                    Dim nameEgSurface As String = userBridge.EarthSurfaceName
                    If IsNothing(nameAlign) = True Then
                        nameAlign = ""
                    End If
                    If IsNothing(nameSurface) = True Then
                        nameSurface = ""
                    End If
                    If IsNothing(nameEgSurface) = True Then
                        nameEgSurface = ""
                    End If
                    Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlignment)
                    If nameSurface.Trim.Length = 0 Then
                        projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
                    Else
                        projectSurface = FuncSurface.getSurfaceByName(nameSurface)
                    End If
                    If nameEgSurface.Trim.Length > 0 Then
                        egSurface = FuncSurface.getSurfaceByName(nameSurface)
                    End If
                End If
            End If
        End If

        Dim tagNodeSelect As conditionalElement = selectedNode.Tag
        Dim dataElement As StructureElement = tagNodeSelect.DataElement
        If IsNothing(dataElement) = False Then
            If IsNothing(dataElement.DWGEntity) = False Then
                Dim typeObjectNode As StructureElement.typeObject = StructureElement.ConvertToTypeObject(selectedNode.Name)
                Dim acLineAxis As DwgLine = New DwgLine
                Dim selectEntModel As List(Of DwgModel3DElement) = New List(Of DwgModel3DElement)
                Dim numberPillar As Integer = 0
                Dim numberProlet As Integer = 0
                If typeObjectNode = typeObject.axisPillar Then
                    numberPillar = FuncGSON.getValue(dataElement.KeyParameter, "Number")
                    If numberPillar > 0 Then
                        acLineAxis = dataElement.DWGEntity
                    End If
                ElseIf selectedNode.Name Like "PillarsAxisBeam" Then
                    'ось опирания балок
                    numberPillar = FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar")
                    numberProlet = FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet")
                    If numberPillar > 0 Then
                        acLineAxis = dataElement.DWGEntity
                    End If
                End If
                If numberPillar > 0 Then
                    dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
                    selectEntModel = Pillar.getModel3DElement(dictionaryBridgeElements, numberPillar, numberProlet)
                End If
                If selectEntModel.Count > 0 Then
                    Dim tempAxisLine As DwgLine = acLineAxis.Clone
                    Dim boolExt As Boolean = BridgeGeometry.extendBeam(tempAxisLine, 5, 5)
                    Dim sectionId As String = "Опора №" & numberPillar
                    SMB.RaiseShowOnWindowSectionModels(sectionId, tempAxisLine, selectEntModel, projectSurface, egSurface)
                End If
            End If
        End If
    End Sub
    'показ одиночного элемента на поперечнике
    Private Sub ПоказатьНаПоперечникеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоказатьНаПоперечникеToolStripMenuItem.Click
        Dim selectedNode As TreeNode = TreeView1.SelectedNode
        If IsNothing(selectedNode) = True Then
            MsgBox("Выберите элемент в списке!!!")
            Exit Sub
        End If
        Dim generalNode As TreeNode = selectedNode
        'ищем корневой элемент
        If IsNothing(selectedNode) = False Then
            Dim levelSelectNode As Integer = selectedNode.Level
            If levelSelectNode > 0 Then
                For i As Integer = levelSelectNode - 1 To 0 Step -1
                    generalNode = generalNode.Parent
                Next
            End If
        End If
        'выбираем данные
        Dim userBridge As Bridges = Nothing
        Dim axisLineBridge As DwgEntity = Nothing
        Dim condAxisBridge As conditionalElement = generalNode.Tag
        If IsNothing(condAxisBridge) = False = True Then
            Dim dataBridge As StructureElement = condAxisBridge.DataElement
            If IsNothing(dataBridge) = False Then
                userBridge = dataBridge.getBridge()
                axisLineBridge = dataBridge.DWGEntity
                If IsNothing(userBridge) = False Then
                    Dim nameAlign As String = userBridge.AlignmentName
                    Dim nameSurface As String = userBridge.projectSurfaceName
                    Dim nameEgSurface As String = userBridge.EarthSurfaceName
                    If IsNothing(nameAlign) = True Then
                        nameAlign = ""
                    End If
                    If IsNothing(nameSurface) = True Then
                        nameSurface = ""
                    End If
                    If IsNothing(nameEgSurface) = True Then
                        nameEgSurface = ""
                    End If
                    Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, projectAlignment)
                    If nameSurface.Trim.Length = 0 Then
                        projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
                    Else
                        projectSurface = FuncSurface.getSurfaceByName(nameSurface)
                    End If
                    If nameEgSurface.Trim.Length > 0 Then
                        egSurface = FuncSurface.getSurfaceByName(nameSurface)
                    End If
                End If
            End If
        End If

        Dim tagNodeSelect As conditionalElement = selectedNode.Tag
        Dim dataElement As StructureElement = tagNodeSelect.DataElement
        If IsNothing(dataElement) = False Then
            If IsNothing(dataElement.DWGEntity) = False Then
                Dim entity As DwgEntity = dataElement.DWGEntity
                If TypeOf entity Is DwgModel3DElement Then
                    Dim numberPillarElement As Integer = FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar")
                    Dim nodeParent As TreeNode = selectedNode.Parent
                    Dim parentDataElement As conditionalElement = nodeParent.Tag
                    Dim parentEntity As DwgEntity = parentDataElement.DataElement.DWGEntity
                    If IsNothing(parentEntity) = False Then
                        If TypeOf parentEntity Is DwgLine Then
                            Dim acLineAxis As DwgLine = parentEntity
                            Dim sectionId As String = "Опора №" & numberPillarElement
                            SMB.RaiseShowOnWindowSectionSelectedModel(sectionId, acLineAxis, entity, projectSurface, egSurface)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'РАБОТА С ОПОРАМИ СООРУЖЕНИЯ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'меню опоры
    Private Sub PillarsMenu_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles PillarsMenu.ItemClicked
        Dim selectNode As TreeNode = TreeView1.SelectedNode
        If IsNothing(selectNode) = True Then
            If IsArray(PropertyGrid1.Tag) = True Then
                Dim arrayTagPropertyGrid As String() = PropertyGrid1.Tag
                If arrayTagPropertyGrid.Length > 2 Then
                    Dim nameNode As String = arrayTagPropertyGrid(0)
                    Dim nodePatch As String = arrayTagPropertyGrid(1)
                    Dim indexNode As Integer = Val(arrayTagPropertyGrid(2))
                    Dim idTempBridge As String = arrayTagPropertyGrid(3)
                    Dim nodeUserFind As TreeNode() = TreeView1.Nodes.Find(nameNode, True)
                    If IsArray(nodeUserFind) = True Then
                        If nodeUserFind.Length > 0 Then
                            For i As Integer = 0 To nodeUserFind.Length - 1
                                Dim node As TreeNode = nodeUserFind(i)
                                If node.FullPath Like nodePatch Then
                                    If node.Index = indexNode Then
                                        If IsArray(node.Tag) = True Then
                                            Dim arrayNodeTag As String() = node.Tag
                                            If arrayNodeTag.Length > 1 Then
                                                If arrayNodeTag(2) Like idTempBridge Then
                                                    selectNode = node
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End If
            End If
        End If
        If IsNothing(selectNode) = True Then
            MsgBox("Выберите сооружение")
            Exit Sub
        Else
            TreeView1.SelectedNode = selectNode
        End If
        '=================================================================================================================
        'находим трассу
        Dim generalNode As TreeNode = selectNode
        Dim idBridge As String = ""
        'ищем корневой элемент
        If IsNothing(selectNode) = False Then
            Dim levelSelectNode As Integer = selectNode.Level
            If levelSelectNode > 0 Then
                For i As Integer = levelSelectNode - 1 To 0 Step -1
                    generalNode = generalNode.Parent
                Next
            End If
        End If
        Dim userBridge As Bridges = Nothing
        Dim axisLineBridge As DwgPolyline = Nothing
        Dim tagSelect As Object = generalNode.Tag
        If TypeOf tagSelect Is conditionalElement Then
            Dim condElement As conditionalElement = tagSelect
            Dim dataElement As StructureElement = condElement.DataElement
            If IsNothing(dataElement) = False Then
                Dim selectedEntity As DwgEntity = dataElement.DWGEntity
                If IsNothing(selectedEntity) = False Then
                    If TypeOf selectedEntity Is DwgPolyline Then
                        axisLineBridge = selectedEntity
                        userBridge = dataElement.getBridge()
                    End If
                End If
            End If
        End If
        If IsNothing(userBridge) = True Then
            MsgBox("Мост не найден!")
            Exit Sub
        End If
        If IsNothing(axisLineBridge) = True Then
            MsgBox("Мост не найден!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'пункты меню
        If IsNothing(bridgeProject) Then
            dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
            If e.ClickedItem.Text Like "Перерассчет элементов опоры" Then
                Dim typeSelectedNode As StructureElement.typeObject = StructureElement.ConvertToTypeObject(selectNode.Name)
                If typeSelectedNode = typeObject.axisPillar Then
                    Dim condElement As conditionalElement = selectNode.Tag
                    Dim dataAxisPillar As StructureElement = condElement.DataElement
                    If IsNothing(dataAxisPillar) = False Then
                        Dim userPillar As Pillar = dataAxisPillar.getPillar()
                        If IsNothing(userPillar) = False Then
                            Dim numberPillar As Integer = userPillar.Number
                            'опора крайняя ищем для нее все элементы
                            If dataAxisPillar.ClassObject = classStructure.LastPillar Then
                                'ищем насадку
                                Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, numberPillar)
                                Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzle)
                                Dim userNozzle As NozzlePillar = dataNozzle.getNozzlePillar()
                                'ищем подферменники
                                Dim arraySubFerm As SubFermenters() = SubFermenters.getSubFermentersPillar(dictionaryBridgeElements, numberPillar)
                                'ищем шкафную стенку
                                Dim listDataСabinetWall As List(Of StructureElement) = CabinetWallPillar.getAxis(dictionaryBridgeElements, numberPillar)
                                Dim dataСabinetWall As StructureElement = StructureElement.isValidateDataStructure(listDataСabinetWall)
                                Dim userСabinetWall As CabinetWallPillar = dataСabinetWall.getCabinetWallPillar()
                                'ищем левое откосное крыло
                                Dim listDataLeftHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Left)
                                Dim dataLeftHand As StructureElement = StructureElement.isValidateDataStructure(listDataLeftHand)
                                Dim userLeftHand As HandPillar = dataLeftHand.getHandPillar
                                'ищем правое откосное крыло
                                Dim listDataRightHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Right)
                                Dim dataRightHand As StructureElement = StructureElement.isValidateDataStructure(listDataRightHand)
                                Dim userRightHand As HandPillar = dataRightHand.getHandPillar
                                'ищем левый открылок
                                Dim listDataLeftPostcard As List(Of StructureElement) = PostcardPillar.getPoctcardPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Left)
                                Dim dataLeftPostcard As StructureElement = StructureElement.isValidateDataStructure(listDataLeftPostcard)
                                Dim userLeftPostcard As PostcardPillar = dataLeftPostcard.getPostcardPillar
                                'ищем правый открылок
                                Dim listDataRightPostcard As List(Of StructureElement) = PostcardPillar.getPoctcardPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Right)
                                Dim dataRightPostcard As StructureElement = StructureElement.isValidateDataStructure(listDataRightPostcard)
                                Dim userRightPostcard As PostcardPillar = dataRightPostcard.getPostcardPillar
                                'ищем стойки
                                Dim arrayRack As RackPillar() = RackPillar.getRackPillarToArray(dictionaryBridgeElements, numberPillar)
                                'ищем ростверк
                                Dim dataGrillage As StructureElement = GrillagePillar.getGrillagePillar(dictionaryBridgeElements, numberPillar)
                                Dim userGrillage As GrillagePillar = dataGrillage.getGrillagePillar
                                'ищем подготовку
                                Dim dataPreparation As StructureElement = PreparationPillar.getPreparationPillar(dictionaryBridgeElements, numberPillar)
                                Dim userPreparation As PreparationPillar = dataPreparation.getPreparationPillar
                                'ищем сваи
                                Dim arrayPile As PilePillar() = PilePillar.getPilePillarToArray(dictionaryBridgeElements, numberPillar)
                                Dim firstPile As String = arrayPile(0).NameModel
                                'делаем предварительный расчет
                                bridgeProject.PlacementLastPillar(dataAxisPillar, userNozzle, arraySubFerm, userСabinetWall, userLeftHand, userRightHand, userLeftPostcard, userRightPostcard, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, projectAlignment)
                                Dim docPileTLC As ConstructionDocument = PileModel.getConstructionElement(generalDir, "firstPile")
                                bridgeProject.DrawingLastPillar(dataAxisPillar, userNozzle, arraySubFerm, userСabinetWall, userLeftHand, userRightHand, userLeftPostcard, userRightPostcard, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, docPileTLC, templateXML)
                            ElseIf dataAxisPillar.ClassObject = classStructure.LastPillar Then
                                'ищем насадку
                                Dim listDataRigel As List(Of StructureElement) = RigelPillar.getAxis(dictionaryBridgeElements, numberPillar)
                                Dim dataRigel As StructureElement = StructureElement.isValidateDataStructure(listDataRigel)
                                Dim userRigel As RigelPillar = dataRigel.getRigelPillar()
                                'ищем подферменники
                                Dim arraySubFerm As SubFermenters() = SubFermenters.getSubFermentersPillar(dictionaryBridgeElements, numberPillar)
                                'ищем стойки
                                Dim arrayRack As RackPillar() = RackPillar.getRackPillarToArray(dictionaryBridgeElements, numberPillar)
                                'ищем ростверк
                                Dim dataGrillage As StructureElement = GrillagePillar.getGrillagePillar(dictionaryBridgeElements, numberPillar)
                                Dim userGrillage As GrillagePillar = dataGrillage.getGrillagePillar
                                'ищем подготовку
                                Dim dataPreparation As StructureElement = PreparationPillar.getPreparationPillar(dictionaryBridgeElements, numberPillar)
                                Dim userPreparation As PreparationPillar = dataPreparation.getPreparationPillar
                                'ищем сваи
                                Dim arrayPile As PilePillar() = PilePillar.getPilePillarToArray(dictionaryBridgeElements, numberPillar)
                                Dim firstPile As String = arrayPile(0).NameModel
                                'делаем предварительный расчет
                                bridgeProject.PlacementMiddlePillar(dataAxisPillar, userRigel, arraySubFerm, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, projectAlignment)
                                Dim docPileTLC As ConstructionDocument = PileModel.getConstructionElement(generalDir, "firstPile")
                                bridgeProject.DrawingMiddlePillar(dataAxisPillar, userRigel, arraySubFerm, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, docPileTLC, templateXML)
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox("Для запуска команды, выберите в опору для удаления!!!")
            End If
        ElseIf e.ClickedItem.Text Like "Удалить элементы всех опор" Then
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'получаем все промежуточные оси опор
            If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisPillar) Then
                Dim listAxisPillar As List(Of StructureElement) = dictionaryBridgeElements.Item(StructureElement.typeObject.axisPillar)
                bridgeProject.RemovePillars(listAxisPillar, dictionaryBridgeElements)
            End If
        ElseIf e.ClickedItem.Text Like "Удалить элементы опоры" Then
            Dim typeSelectedNode As StructureElement.typeObject = StructureElement.ConvertToTypeObject(selectNode.Name)
            If typeSelectedNode = typeObject.axisPillar Then
                Dim condElement As conditionalElement = selectNode.Tag
                Dim dataAxisPillar As StructureElement = condElement.DataElement
                If IsNothing(dataAxisPillar) = False Then
                    Dim listAxisPillar As List(Of StructureElement) = New List(Of StructureElement) From {dataAxisPillar}
                    bridgeProject.RemovePillars(listAxisPillar, dictionaryBridgeElements)
                Else
                    MsgBox("Для запуска команды, выберите в опору для удаления!!!")
                End If
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОТЧЕТЫ
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub ReportMenu_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ReportMenu.ItemClicked
        Dim selectNode As TreeNode = TreeView1.SelectedNode
        'ищем корневой элемент
        Dim generalNode As TreeNode = Nothing
        If IsNothing(selectNode) = False Then
            Dim levelSelectNode As Integer = selectNode.Level
            If levelSelectNode > 0 Then
                For i As Integer = levelSelectNode - 1 To 0 Step -1
                    generalNode = generalNode.Parent
                Next
            End If
        End If
        If IsNothing(generalNode) = True Then
            Exit Sub
        End If
        Dim conditionalNodeTag As conditionalElement = generalNode.Tag
        If IsNothing(conditionalNodeTag) = True Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

        Dim dataBridge As StructureElement = conditionalNodeTag.DataElement
        If IsNothing(dataBridge) = False Then
            Dim idBridge As String = conditionalNodeTag.DataElement.IdStructure
            If idBridge.Trim.Length > 1 Then
                Dim brReport As BridgeReports = New BridgeReports
                If e.ClickedItem.Text Like "Точки опирания балок" Then
                    brReport.ReportPointPrBeams(dataBridge)
                    Exit Sub
                ElseIf e.ClickedItem.Text Like "Верх плиты балок и толщины покрытия" Then
                    brReport.ReportUpPokr(dataBridge)
                    Exit Sub
                ElseIf e.ClickedItem.Text Like "Оси опор" Then
                    brReport.ReportAxisPillars(dataBridge)
                    Exit Sub
                ElseIf e.ClickedItem.Text Like "Деформационные зазоры" Then
                    brReport.ReportDefZazor(dataBridge)
                    Exit Sub
                ElseIf e.ClickedItem.Text Like "Экспорт элементов в dwg" Then
                    brReport.ExportToAutucad(dataBridge)
                    Exit Sub
                End If
            End If
        End If

    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменить размеры TreeView и datagrid
    Private Sub TreeView1_MouseDown(sender As Object, e As Forms.MouseEventArgs) Handles TreeView1.MouseDown
        X1 = e.X
        Me.Cursor = Forms.Cursors.HSplit
        Y1 = e.Y
    End Sub
    Private Sub TreeView1_MouseUp(sender As Object, e As Forms.MouseEventArgs) Handles TreeView1.MouseUp
        X2 = e.X
        Y2 = e.Y
        Me.Cursor = Forms.Cursors.Default
        If Y1 = 0 Then
            Exit Sub
        End If
        Dim delta As Integer = Y1 - Y2
        If delta > 0 Then
            TreeView1.Height = TreeView1.Height - delta
            Dim loc As System.Drawing.Point = New System.Drawing.Point(PropertyGrid1.Location.X, PropertyGrid1.Location.Y - delta)
            PropertyGrid1.Location = loc
            PropertyGrid1.Height = PropertyGrid1.Height + delta
        ElseIf delta < 0 Then
            TreeView1.Height = TreeView1.Height + Math.Abs(delta)
            Dim loc As System.Drawing.Point = New System.Drawing.Point(PropertyGrid1.Location.X, PropertyGrid1.Location.Y + Math.Abs(delta))
            PropertyGrid1.Location = loc
            PropertyGrid1.Height = PropertyGrid1.Height + delta
        End If
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
    'смена проекта err
    Private Sub CBox_ListNamesArrProject_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesArrProject.SelectedIndexChanged
        Dim indexSelectArrModel As Integer = CBox_ListNamesArrProject.SelectedIndex
        If indexSelectArrModel > -1 And IsNothing(projectCivil) = False Then
            'получаем ссылку на проект arr по индексу
            arrProject = projectCivil.getArrangementModelByIndex(indexSelectArrModel)
            If IsNothing(arrProject) = False Then
                bridgeProject = New ProjectBridge()
                bridgeProject.BridgeModel = arrProject
                bridgeProject.getBridges(False)
                ActivDocumentPanel = arrProject.Drawing
                dictNamesProjectBridge = bridgeProject.getDictionaryNamesBridge()
                If IsNothing(dictNamesProjectBridge) = False Then
                    If dictNamesProjectBridge.Count > 0 Then
                        idBridge = dictNamesProjectBridge.ElementAt(0).Key
                        dataBridge = bridgeProject.getStructureElementBridgeByIndex(0)
                        userBridge = bridgeProject.getBridgeByID(idBridge)
                    End If
                End If
            End If
        End If
    End Sub
End Class
Public Class conditionalElement
    Private _numberElement As Integer
    Private _numberSubPillar As Integer
    Private _numberRow As Integer
    Private _numberColl As Integer
    Private _sideElement As Pillar.SidePillarElement
    Private _dataElement As StructureElement
    Public Sub New()
        _numberElement = 0
        _numberSubPillar = 0
        _numberRow = 0
        _numberColl = 0
        _sideElement = Pillar.SidePillarElement.None
        _dataElement = New StructureElement
    End Sub
    Public Sub New(Number As Integer, DataElement As StructureElement)
        _numberElement = Number
        _dataElement = DataElement
    End Sub
    ' ==========================================
    ' Свойства (Properties) с Get/Set
    ' ==========================================
    Public Property NumberElement As Integer
        Get
            Return _numberElement
        End Get
        Set(value As Integer)
            _numberElement = value
        End Set
    End Property

    Public Property NumberSubPillar As Integer
        Get
            Return _numberSubPillar
        End Get
        Set(value As Integer)
            _numberSubPillar = value
        End Set
    End Property

    Public Property NumberRow As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property

    Public Property NumberColl As Integer
        Get
            Return _numberColl
        End Get
        Set(value As Integer)
            _numberColl = value
        End Set
    End Property

    Public Property SideElement As Pillar.SidePillarElement
        Get
            Return _sideElement
        End Get
        Set(value As Pillar.SidePillarElement)
            _sideElement = value
        End Set
    End Property

    Public Property DataElement As StructureElement
        Get
            Return _dataElement
        End Get
        Set(value As StructureElement)
            _dataElement = value
        End Set
    End Property
End Class
