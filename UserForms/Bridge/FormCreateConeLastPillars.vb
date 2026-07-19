Imports System.IO
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Sites.Core
Imports Topomatic.Visualization.Geometry

Public Class FormCreateConeLastPillars
    Public boolShow As Boolean = False
    Public boolSelectLine As Boolean = False
    Public generalDir As String = ""
    Public templateXML As String = ""
    Public idBridge As String = ""
    Public activProjectDocument As Topomatic.Dwg.Drawing = Nothing
    Public activDocumentSite As Topomatic.Dwg.Drawing = Nothing
    Public civilStructuresProject As ProjectCivilStructures = Nothing 'проекты arr
    Public bridgeProject As ProjectBridge = Nothing
    Public arrProject As ArrangementModel = Nothing
    Public dataStructuresBridge As StructureElement = Nothing
    Public dictionaryBridge As Dictionary(Of String, StructureElement) = Nothing 'словарь с мостами активного чертежа
    Public dictNamesProjectBridge As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = New Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) 'словарь с элементами чертежа
    Public dataPillar As StructureElement = Nothing
    Public axisLinePillar As DwgLine = New DwgLine
    Public userAxisPillar As Pillar = Nothing
    Public userBridge As Bridges = Nothing
    Public listPolyline3d As List(Of DwgPolyline3D) = New List(Of DwgPolyline3D)

    Public brigeGeneralAxisDictionary As Dictionary(Of String, String()) = New Dictionary(Of String, String()) 'словарь с осями сооружений подъобъекта


    Public axisLineBridge As DwgPolyline = Nothing
    Public arrayNameBridges As String() = {""} 'массив с именами сооружений
    Public numberPillar As Integer = 0

    'трасса
    Public projectAlignment As Alignment = Nothing
    'поверхность
    Public egSurface As Surface = Nothing
    Public projectSurface As Surface = Nothing
    Public startConeLine As DwgLine = Nothing
    Public k As Integer = 1
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        Dim dirTemplate As String() = {""}
        Dim directorySupport As String = generalDir & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    Dim tempXml As String = xmlFiles(i)
                    ReDim Preserve dirTemplate(i)
                    dirTemplate(i) = IO.Path.GetFileNameWithoutExtension(tempXml)
                    If templateXML.Trim.Length = 0 Then
                        templateXML = tempXml
                    End If
                Next
            End If
        End If
        If IsArray(dirTemplate) = True Then
            CBox_ListNamesTemplateXML.DataSource = dirTemplate
            CBox_ListNamesTemplateXML.Tag = directorySupport
        End If
        If IsNothing(civilStructuresProject) = False Then
            CB_NameSites.DataSource = civilStructuresProject.ListModelSites
        End If

        DGV_PropertiesCone.Rows.Clear()
        DGV_PropertiesCone.Rows.Add(15)
        DGV_PropertiesCone.Rows(0).Cells(0).Value = "Длина конуса слева до края относного крыла, м"
        DGV_PropertiesCone.Rows(0).Cells(1).Value = 5
        DGV_PropertiesCone.Rows(1).Cells(0).Value = "Длина конуса справа до края относного крыла, м"
        DGV_PropertiesCone.Rows(1).Cells(1).Value = 5
        DGV_PropertiesCone.Rows(2).Cells(0).Value = "Длина вдоль левого откосного крыла до начала закругления, м"
        DGV_PropertiesCone.Rows(2).Cells(1).Value = 1
        DGV_PropertiesCone.Rows(3).Cells(0).Value = "Длина вдоль правого откосного крыла до начала закругления, м"
        DGV_PropertiesCone.Rows(3).Cells(1).Value = 1
        DGV_PropertiesCone.Rows(4).Cells(0).Value = "Уширение конуса слева от откосного крыла, м"
        DGV_PropertiesCone.Rows(4).Cells(1).Value = 0.75
        DGV_PropertiesCone.Rows(5).Cells(0).Value = "Уширение конуса справа от откосного крыла, м"
        DGV_PropertiesCone.Rows(5).Cells(1).Value = 0.75
        DGV_PropertiesCone.Rows(6).Cells(0).Value = "Уширение конуса вдоль насадки, м"
        DGV_PropertiesCone.Rows(6).Cells(1).Value = 0.15
        DGV_PropertiesCone.Rows(7).Cells(0).Value = "Коэффициет заложения откоса конуса, 1:"
        DGV_PropertiesCone.Rows(7).Cells(1).Value = 1
        DGV_PropertiesCone.Rows(8).Cells(0).Value = "Радиус нижнего основания, м"
        DGV_PropertiesCone.Rows(8).Cells(1).Value = 2
        DGV_PropertiesCone.Rows(9).Cells(0).Value = "Смещение вниз от верха насадки, м"
        DGV_PropertiesCone.Rows(9).Cells(1).Value = 0.3
        DGV_PropertiesCone.Rows(10).Cells(0).Value = "Отметка верха обочины слева в начале конуса, м"
        DGV_PropertiesCone.Rows(10).Cells(1).Value = 0
        DGV_PropertiesCone.Rows(11).Cells(0).Value = "Отметка верха обочины справа в начале конуса, м"
        DGV_PropertiesCone.Rows(11).Cells(1).Value = 0
        DGV_PropertiesCone.Rows(12).Cells(0).Value = "Отметка верха обочины слева в районе закругления конуса, м"
        DGV_PropertiesCone.Rows(12).Cells(1).Value = 0
        DGV_PropertiesCone.Rows(13).Cells(0).Value = "Отметка верха обочины справа в районе закругления конуса, м"
        DGV_PropertiesCone.Rows(13).Cells(1).Value = 0
        DGV_PropertiesCone.Rows(14).Cells(0).Value = "Отметка основаия конуса, м"
        DGV_PropertiesCone.Rows(14).Cells(1).Value = 0
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор проекта arr
    Private Sub CBox_ListNamesArrProject_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesArrProject.SelectedIndexChanged
        Dim indexSelectArrModel As Integer = CBox_ListNamesArrProject.SelectedIndex
        If indexSelectArrModel > -1 And IsNothing(civilStructuresProject) = False Then
            'получаем ссылку на проект arr по индексу
            arrProject = civilStructuresProject.getArrangementModelByIndex(indexSelectArrModel)
            If IsNothing(arrProject) = False Then
                bridgeProject = New ProjectBridge()
                bridgeProject.BridgeModel = arrProject
                bridgeProject.getBridges(False)
                activProjectDocument = arrProject.Drawing
                dictNamesProjectBridge = bridgeProject.getDictionaryNamesBridge()
                If IsNothing(dictNamesProjectBridge) = False Then
                    If dictNamesProjectBridge.Count > 0 Then
                        idBridge = dictNamesProjectBridge.ElementAt(0).Key
                        userBridge = bridgeProject.getBridgeByID(idBridge)
                        dataStructuresBridge = bridgeProject.getStructureElementBridgeByIndex(0)
                        CBox_ListNamesBridge.DataSource = dictNamesProjectBridge.Values.ToList
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор сооружения
    Private Sub CBox_ListNamesBridge_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesBridge.SelectedIndexChanged
        '=========================================================================================
        Dim indexSelectBridge As Integer = CBox_ListNamesBridge.SelectedIndex
        If IsNothing(bridgeProject) = False Then
            If indexSelectBridge < bridgeProject.ListBridges.Count Then
                If IsNothing(arrProject) = False Then
                    'bridgeProject.BridgeModel = arrProject
                    'bridgeProject.getBridgeByIndex(indexSelectBridge)
                    dataStructuresBridge = bridgeProject.getStructureElementBridgeByIndex(indexSelectBridge)
                    If IsNothing(dataStructuresBridge) = False Then
                        idBridge = dataStructuresBridge.IdStructure
                        userBridge = dataStructuresBridge.getBridge()
                        If IsNothing(userBridge) = False Then
                            activProjectDocument = arrProject.Drawing
                            Dim axisEnt As DwgEntity = dataStructuresBridge.DWGEntity
                            If IsNothing(axisEnt) = False Then
                                If TypeOf axisEnt Is DwgPolyline Then
                                    axisLineBridge = axisEnt
                                End If
                                dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
                                If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
                                    Dim dictPillar As Dictionary(Of Integer, StructureElement) = userBridge.getAxisPillars(dictionaryBridgeElements, False)
                                    Dim listNumberPillar As List(Of String) = New List(Of String) From {""}
                                    If dictPillar.Count > 0 Then
                                        listNumberPillar.Add(dictPillar.First.Key)
                                        listNumberPillar.Add(dictPillar.Last.Key)
                                    End If
                                    CB_NumberPillar.DataSource = listNumberPillar
                                    CB_NumberPillar.Text = ""
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Не удалось получить элементы мостового сооружения.")
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор крайней опоры
    Private Sub CB_NumberPillar_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_NumberPillar.SelectedIndexChanged
        Dim strNumberPillar As String = CB_NumberPillar.Text
        If IsNothing(strNumberPillar) = True Then Exit Sub
        If strNumberPillar.Trim.Length = 0 Then Exit Sub
        numberPillar = Val(CB_NumberPillar.Text)
        If IsNothing(dataStructuresBridge) = False Then
            axisLineBridge = dataStructuresBridge.DWGEntity
        Else
            Exit Sub
        End If
        If IsNothing(axisLineBridge) = False Then
            dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
        End If
        '=======================================================================================================================
        'находим ось опоры
        If IsNothing(dictionaryBridgeElements) = False Then
            If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
                dataPillar = Pillar.getAxisPillar(dictionaryBridgeElements, numberPillar)
                userAxisPillar = dataPillar.getPillar()
                axisLinePillar = dataPillar.DWGEntity
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'РАСЧЕТ ПРЕДВАРИТЕЛЬНЫЙ
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim nameBridge As String = CBox_ListNamesBridge.Text
        Dim nameEgSurface As String = CB_EgSurface.Text
        Dim nameProjectSurface As String = CB_ProjectSurface.Text
        Dim nameAlignment As String = CB_NameAlignment.Text
        'имя сооружения
        Dim nameModel As String = CBox_ListNamesArrProject.Text
        If IsNothing(projectAlignment) = True Then
            Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlignment, projectAlignment)
            If IsNothing(projectAlignment) = True Then
                MsgBox("Проектная ось трассы автомобильной дороги, не найдена!!!")
                Exit Sub
            End If
        End If
        'фактическая поверхность
        egSurface = FuncSurface.getSurfaceByName(nameEgSurface)
        If ChB_ProjectSurfaceInAlignment.Checked = False Then
            projectSurface = FuncSurface.getSurfaceByName(nameProjectSurface)
        Else
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlignment)
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена или выключена!!!")
            Exit Sub
        End If
        Dim userNozzle As NozzlePillar = Nothing
        Dim userLeftHand As HandPillar = Nothing
        Dim userRightHand As HandPillar = Nothing
        'находим точки насадки и левого и правого Hand
        If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisNozzle) = True Then
            Dim listDataNozzles As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, numberPillar)
            If IsNothing(listDataNozzles) = False Then
                If listDataNozzles.Count > 0 Then
                    Dim dataNozzle As StructureElement = StructureElement.isValidateDataStructure(listDataNozzles)
                    userNozzle = dataNozzle.getNozzlePillar
                End If
            End If
        End If
        If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisLeftHand) = True Then
            Dim dataLeftHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Left)
            If IsNothing(dataLeftHand) = False Then
                If dataLeftHand.Count > 0 Then
                    Dim dataHand As StructureElement = StructureElement.isValidateDataStructure(dataLeftHand)
                    userLeftHand = dataHand.getHandPillar
                End If
            End If
        End If
        If dictionaryBridgeElements.ContainsKey(StructureElement.typeObject.axisRightHand) = True Then
            Dim dataRightHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, Pillar.SidePillarElement.Right)
            If IsNothing(dataRightHand) = False Then
                If dataRightHand.Count > 0 Then
                    Dim dataHand As StructureElement = StructureElement.isValidateDataStructure(dataRightHand)
                    userRightHand = dataHand.getHandPillar
                End If
            End If
        End If
        Dim userCone As ConesPillar = New ConesPillar
        userCone.NumberPillar = Val(CB_NumberPillar.Text)
        userCone.LenghtLeftTop = DGV_PropertiesCone.Rows(0).Cells(1).Value
        userCone.LenghtRightTop = DGV_PropertiesCone.Rows(1).Cells(1).Value
        userCone.LenghtRightHand = DGV_PropertiesCone.Rows(2).Cells(1).Value
        userCone.LenghtLeftHand = DGV_PropertiesCone.Rows(3).Cells(1).Value
        userCone.OffsetLeftHand = DGV_PropertiesCone.Rows(4).Cells(1).Value
        userCone.OffsetRightHand = DGV_PropertiesCone.Rows(5).Cells(1).Value
        userCone.HorizontalOffsetPlate = DGV_PropertiesCone.Rows(6).Cells(1).Value
        userCone.Slope = DGV_PropertiesCone.Rows(7).Cells(1).Value
        userCone.BottomRadius = DGV_PropertiesCone.Rows(8).Cells(1).Value
        userCone.VerticalOffsetNozzle = DGV_PropertiesCone.Rows(9).Cells(1).Value
        userCone.ElevationLeftEgeStart = DGV_PropertiesCone.Rows(10).Cells(1).Value
        userCone.ElevationRightEgeStart = DGV_PropertiesCone.Rows(11).Cells(1).Value
        userCone.ElevationLeftMiddleHand = DGV_PropertiesCone.Rows(12).Cells(1).Value
        userCone.ElevationRightMiddleHand = DGV_PropertiesCone.Rows(13).Cells(1).Value
        userCone.ElevationGround = DGV_PropertiesCone.Rows(14).Cells(1).Value
        Dim countSegments As Integer = NUpD_CountSegmets.Value
        listPolyline3d = userCone.calulateFirstLine(userNozzle, userLeftHand, userRightHand, countSegments, projectSurface)
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОТМЕНА
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = False
        boolSelectLine = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОК создаем конус
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        boolShow = True
        boolSelectLine = False
        Me.Hide()
        Dim modelProject As ModelProject = ApplicationHost.Current.ActiveProject
        Dim modelProjectChilds As IProjectModel() = ModelProject.Model.GetChilds()
        If IsNothing(activProjectDocument) = True Then
            Dim nameBridgeProject As String = CBox_ListNamesArrProject.Text
            For Each modelProjectChild As IProjectModel In modelProjectChilds
                Dim modelProjectUri As Topomatic.FoundationClasses.URI = modelProjectChild.Uri
                Dim modelFile As String = modelProjectUri.AsFilePath
                Dim nameFilePrj As String = IO.Path.GetFileNameWithoutExtension(modelFile)
                If nameFilePrj Like nameBridgeProject Then
                    If modelProjectChild.ModelType Like "arr" Then
                        Dim model As ArrangementModel = modelProjectChild.Model
                        activProjectDocument = model.Drawing
                        Exit For
                    End If
                End If
            Next
        End If
        Dim coneDrawSurface As Surface = Nothing
        If IsNothing(activDocumentSite) = True Then
            Dim nameSiteProject As String = CB_NameSites.Text
            For Each modelProjectChild As IProjectModel In modelProjectChilds
                Dim modelProjectUri As Topomatic.FoundationClasses.URI = modelProjectChild.Uri
                Dim modelFile As String = modelProjectUri.AsFilePath
                Dim nameFilePrj As String = IO.Path.GetFileNameWithoutExtension(modelFile)
                If nameFilePrj Like nameSiteProject Then
                    If modelProjectChild.ModelType Like "site" Then
                        Dim model As SiteModel = modelProjectChild.Model
                        activDocumentSite = model.Drawing
                        coneDrawSurface = model.Surface
                        Exit For
                    End If
                End If
            Next
        End If
        'стоим структурные линии
        If IsNothing(coneDrawSurface) = False Then
            If listPolyline3d.Count > 1 Then
                For i As Integer = 0 To listPolyline3d.Count - 1
                    Dim poly3d As DwgPolyline3D = listPolyline3d.Item(i)
                    Dim structLine As StructureLine = StructuresLines.CreateStructureLineByPolyline3d(coneDrawSurface, poly3d)
                Next
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать площадку
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim userProject As ModelProject = ApplicationHost.Current.ActiveProject
        Try
            If IsNothing(userProject) = False Then
                Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "ИССО", "Путепроводы", "Конусы"}))
                Try
                    Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "site"})
                    ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                    If IsNothing(iUserProject) = False Then
                        Dim siteProject As SiteModel = iUserProject.Model
                        Dim nameSite As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {siteProject})
                        Dim listSites As List(Of SiteModel) = civilStructuresProject.ListModelSites
                        If IsNothing(listSites) = True Then
                            listSites = New List(Of SiteModel)
                        End If
                        listSites.Add(siteProject)
                        activDocumentSite = siteProject.Drawing
                        CB_NameSites.DataSource = civilStructuresProject.listNameSitesModels
                    End If
                Catch ex As System.OperationCanceledException
                End Try
            End If
        Catch ex As Topomatic.ApplicationPlatform.MessageException
            MsgBox(ex.Message)
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор опорной линии
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        boolSelectLine = True
        Me.Hide()
    End Sub

    Private Sub FormCreateConeLastPillars_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ChB_ProjectSurfaceInAlignment_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_ProjectSurfaceInAlignment.CheckedChanged
        If ChB_ProjectSurfaceInAlignment.Checked = True Then
            CB_ProjectSurface.Enabled = False
        Else
            CB_ProjectSurface.Enabled = True
        End If
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class