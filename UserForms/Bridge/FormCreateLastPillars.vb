Imports System.ComponentModel
Imports System.Diagnostics.Contracts
Imports System.Drawing
Imports System.IO
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media
Imports System.Xml
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Offsets
Imports Topomatic.Alg.Road
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Runtime
Imports Color = System.Drawing.Color
Imports Control = System.Windows.Forms.Control
Public Class FormCreateLastPillars
    Public boolShow As Boolean = False
    Public activProjectDocument As Topomatic.Dwg.Drawing = Nothing
    Public civilStructuresProject As ProjectCivilStructures = Nothing 'проекты arr
    Public bridgeProject As ProjectBridge = Nothing
    Public arrProject As ArrangementModel = Nothing

    Public dictionaryBridge As Dictionary(Of String, StructureElement) = Nothing 'словарь с мостами активного чертежа
    Public dictNamesProjectBridge As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Public dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = New Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) 'словарь с элементами чертежа

    Public dataStructuresBridge As StructureElement = Nothing
    Public userBridge As Bridges = Nothing
    Public formUserTempl As FormUserTemptate = New FormUserTemptate
    Public userMatrixTransform As TransformCivilStructure = New TransformCivilStructure()

    Public generalDir As String = ""
    Public schemaDir As String = ""
    Public templateXML As String = ""
    'сооружение
    Public arrayNameBridges As String() = {""} 'массив с именами сооружений
    Public axisLineBridge As DwgPolyline = Nothing
    Public idBridge As String = ""
    'характеристики опоры
    Public numberPillar As Integer = 0
    Public numberProlet As Integer = 0
    Public numberSubPillar As Integer = 0
    Public arrayBeams As String(,) = {}
    Public minElevationBeam As Double = 0 'отметка самой нижней балки в опоре
    Public dataPillar As StructureElement = Nothing
    Public axisLinePillar As DwgLine = New DwgLine
    Public userAxisPillar As Pillar = Nothing
    Public clearanceBeam As Double = 0
    Public lineCabinetWall As DwgLine = New DwgLine()
    Public directionAngleNozzle As Double = 0
    Public reverseDirectionAngleNozzle As Double = 0
    Public maxElevation As Double = -9999
    'крайние балки
    Public leftBeam As DwgLine = New DwgLine()
    Public leftPlateBeam As DwgLine = New DwgLine()
    Public userLeftBeam As BeamI = Nothing
    Public rightBeam As DwgLine = New DwgLine()
    Public rightPlateBeam As DwgLine = New DwgLine()
    Public userRightBeam As BeamI = Nothing
    'трасса
    Public projectAlignment As Alignment = Nothing
    'поверхность
    Public egSurface As Surface = Nothing
    Public projectSurface As Surface = Nothing
    Public elevationEgSurface As Double = -1
    Public elevationStartProjectSurface As Double = -1
    Public elevationEndProjectSurface As Double = -1
    'насадка
    Public dataNozzle As StructureElement = Nothing
    Public userNozzle As NozzlePillar = New NozzlePillar
    Public axisLineNozzle As DwgLine = New DwgLine()
    Public listFullPointNozzle As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointNozzle As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public lineLeftConsoleNozzle As DwgLine = New DwgLine()
    Public lineRightConsoleNozzle As DwgLine = New DwgLine()
    Public polyTopNozzle As DwgPolyline3D = New DwgPolyline3D
    Public polyBottomNozzle As DwgPolyline3D = New DwgPolyline3D
    Public putchNozzle As String = ""
    Public docNozzleTLC As ConstructionDocument = New ConstructionDocument()
    'подферменник
    Public userSubFermenter As SubFermenters() = {}
    Public dictProjectPointPrevSubFermenters As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint))) = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
    Public dictProjectPointNextSubFermenters As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint))) = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
    Public putchSubFermenter As String = ""
    Public docSubFermenterTLC As ConstructionDocument = New ConstructionDocument()
    'шкафная стенка
    Public dataCabinetWall As StructureElement = Nothing
    Public userCabinetWall As CabinetWallPillar = New CabinetWallPillar()
    Public axisLineCabinetWall As DwgLine = New DwgLine()
    Public listFullPointCabinetWall As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointCabinetWall As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointPlateCabinetWall As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public polyTopCabinetWall As DwgPolyline3D = New DwgPolyline3D
    Public polyBottomCabinetWall As DwgPolyline3D = New DwgPolyline3D
    Public polyTopPlateCabinetWall As DwgPolyline3D = New DwgPolyline3D
    Public polyBottomPlateCabinetWall As DwgPolyline3D = New DwgPolyline3D
    Public putchCabinetWall As String = ""
    Public docCabinetWallTLC As ConstructionDocument = New ConstructionDocument()
    'Откосные крылья
    Public dataLeftHand As StructureElement = Nothing
    Public dataRightHand As StructureElement = Nothing
    Public userLeftHand As HandPillar = New HandPillar()
    Public userRightHand As HandPillar = New HandPillar()
    Public axisLineLeftHand As DwgLine = New DwgLine()
    Public axisLineRightHand As DwgLine = New DwgLine()
    Public listFullPointLeftHand As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointLeftHand As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointLeftHandCornice As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listFullPointRightHand As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointRightHand As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointRightHandCornice As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))

    Public polyLeftTopHand As DwgPolyline3D = New DwgPolyline3D
    Public polyLeftBottomHand As DwgPolyline3D = New DwgPolyline3D
    Public polyRightTopHand As DwgPolyline3D = New DwgPolyline3D
    Public polyRightBottomHand As DwgPolyline3D = New DwgPolyline3D
    Public polyLeftTopCornice As DwgPolyline3D = New DwgPolyline3D
    Public polyLeftBottomCornice As DwgPolyline3D = New DwgPolyline3D
    Public polyRightTopCornice As DwgPolyline3D = New DwgPolyline3D
    Public polyRightBottomCornice As DwgPolyline3D = New DwgPolyline3D
    Public putchHand As String = ""
    Public docLeftHandTLC As ConstructionDocument = New ConstructionDocument()
    Public docRightHandTLC As ConstructionDocument = New ConstructionDocument()
    'Обратные открылки
    Public dataLeftPostcard As StructureElement = Nothing
    Public dataRightPostcard As StructureElement = Nothing
    Public userLeftPostcard As PostcardPillar = New PostcardPillar()
    Public userRightPostcard As PostcardPillar = New PostcardPillar()
    Public axisLineLeftPostcard As DwgLine = New DwgLine()
    Public axisLineRightPostcard As DwgLine = New DwgLine()
    Public listFullPointLeftPostcard As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointLeftPostcard As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listFullPointRightPostcard As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointRightPostcard As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))

    Public polyLeftTopPostcard As DwgPolyline3D = New DwgPolyline3D
    Public polyLeftBottomPostcard As DwgPolyline3D = New DwgPolyline3D
    Public polyRightTopPostcard As DwgPolyline3D = New DwgPolyline3D
    Public polyRightBottomPostcard As DwgPolyline3D = New DwgPolyline3D
    Public putchPostcard As String = ""
    Public docLeftPostcardTLC As ConstructionDocument = New ConstructionDocument()
    Public docRightPostcardTLC As ConstructionDocument = New ConstructionDocument()
    'Ростверк
    Public dataGrillage As StructureElement = Nothing
    Public userGrillage As GrillagePillar = New GrillagePillar()
    Public axisLineGrillage As DwgLine = New DwgLine()
    Public listFullPointGrillage As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointGrillage As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public polyTopGrillage As DwgPolyline3D = New DwgPolyline3D
    Public polyBottomGrillage As DwgPolyline3D = New DwgPolyline3D
    Public putchGrillage As String = ""
    Public docGrillageTLC As ConstructionDocument = New ConstructionDocument()
    'Подготовка
    Public dataPreparation As StructureElement = Nothing
    Public userPreparation As PreparationPillar = New PreparationPillar
    Public axisLinePreparation As DwgLine = New DwgLine()
    Public listFullPointPreparation As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public listPointPreparation As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
    Public polyTopPreparation As DwgPolyline3D = New DwgPolyline3D
    Public polyBottomPreparation As DwgPolyline3D = New DwgPolyline3D
    Public putchPreparation As String = ""
    Public docPreparationTLC As ConstructionDocument = New ConstructionDocument()
    'Стойки
    Public userRack As RackPillar = New RackPillar() 'по умолчанию (общие настройки)
    Public arrayRack As RackPillar() = {} 'по умолчанию (общие настройки)
    Public listPointRack As Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint))) = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
    Public dictionaryDataRack As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
    Public dictionaryUserRack As Dictionary(Of Integer, RackPillar) = New Dictionary(Of Integer, RackPillar)
    Public dictionaryAxisLineRack As Dictionary(Of Integer, DwgLine) = New Dictionary(Of Integer, DwgLine)
    Public arrayPolyTopRack As DwgPolyline3D() = {}
    Public arrayPolyBottomRack As DwgPolyline3D() = {}
    Public putchRack As String = ""
    Public docRackTLC As ConstructionDocument = New ConstructionDocument()
    'Сваи
    Public userPile As PilePillar = New PilePillar()
    Public arrayPile As PilePillar() = {}
    Public dictProjectionPointPile As New Dictionary(Of Integer, Dictionary(Of Integer, ProjectionPoint()))
    Public dictDataPile As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = New Dictionary(Of Integer, Dictionary(Of Integer, StructureElement))
    Public dictUserPile As Dictionary(Of Integer, Dictionary(Of Integer, PilePillar)) = New Dictionary(Of Integer, Dictionary(Of Integer, PilePillar))
    Public arrayUserPile As PilePillar() = {}
    Public putchPile As String = ""
    Public docPileTLC As ConstructionDocument = New ConstructionDocument()
    Public typePile As String = ""
    'ригель
    Public userRigel As RigelPillar = New RigelPillar
    '===============================================================================================
    Public guidPillar As String = ""
    Public categoryTables As String = "Искусственные сооружения"
    Public nameTableBridge As String = "Мостовое сооружение"
    Public nameTablePillars As String = "Опоры мостовых сооружений"
    Public tablePS As String = "PROJECT_BRIDGE"
    Public nameElement As String = ""
    Public tagSelect As String = ""
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'инициализация формы
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
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

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'библиотека к схемам парамеров элементов
        Dim templateDir As String = generalDir & "\TopomaticRobur\DesignBridge\UserProperties\LastPillarsSchema\"
        If Directory.Exists(templateDir) = True Then
            schemaDir = templateDir
        Else
            'создаем новый каталог
            Try
                Directory.CreateDirectory(generalDir & "\TopomaticRobur\DesignBridge\UserProperties\LastPillarsSchema\")
            Catch ex As Exception
                MsgBox("Не удалось создать каталог: " & generalDir & "\TopomaticRobur\DesignBridge\UserProperties\LastPillarsSchema\")
            End Try
        End If
        If Directory.Exists(templateDir) = True Then
            ComboBox11.Tag = schemaDir
            Dim arrayShemaFiles As String() = Directory.GetFiles(schemaDir, "*.xml")
            If IsArray(arrayShemaFiles) = True Then
                Dim arrayFileName As String() = {""}
                Dim count As Integer = 1
                For i As Integer = 0 To arrayShemaFiles.Length - 1
                    ReDim Preserve arrayFileName(count)
                    Dim fInfo As FileInfo = New FileInfo(arrayShemaFiles(i))
                    arrayFileName(count) = fInfo.Name
                    count += 1
                Next
                ComboBox11.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для насадки
        putchNozzle = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Насадка\"
        If Directory.Exists(putchNozzle) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchNozzle, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_NozzleTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для шкафной стенки
        putchCabinetWall = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Шкафная стенка\"
        If Directory.Exists(putchCabinetWall) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchCabinetWall, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_CabinetWallTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Откосных крыльев
        putchHand = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Откосные крылья\"
        If Directory.Exists(putchHand) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchHand, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileLeftName As String() = {}
                Dim arrayFileRightName As String() = {}
                Dim countLeft As Integer = 0
                Dim countRight As Integer = 0
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    If fInfo.Name.IndexOf("left") > -1 Or fInfo.Name.IndexOf("Left") > -1 Then
                        ReDim Preserve arrayFileLeftName(countLeft)
                        arrayFileLeftName(countLeft) = fInfo.Name
                        countLeft += 1
                    ElseIf fInfo.Name.IndexOf("right") > -1 Or fInfo.Name.IndexOf("Right") > -1 Then
                        ReDim Preserve arrayFileRightName(countRight)
                        arrayFileRightName(countRight) = fInfo.Name
                        countRight += 1
                    Else
                        ReDim Preserve arrayFileLeftName(countLeft)
                        arrayFileLeftName(countLeft) = fInfo.Name
                        countLeft += 1
                        ReDim Preserve arrayFileRightName(countRight)
                        arrayFileRightName(countRight) = fInfo.Name
                        countRight += 1
                    End If
                Next
                CB_LeftHandTLC.DataSource = arrayFileLeftName
                CB_RightHandTLC.DataSource = arrayFileRightName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Обратные открылки
        putchPostcard = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Обратные открылки\"
        If Directory.Exists(putchPostcard) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchPostcard, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileLeftName As String() = {}
                Dim arrayFileRightName As String() = {}
                Dim countLeft As Integer = 0
                Dim countRight As Integer = 0
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    If fInfo.Name.IndexOf("left") > -1 Or fInfo.Name.IndexOf("Left") > -1 Then
                        ReDim Preserve arrayFileLeftName(countLeft)
                        arrayFileLeftName(countLeft) = fInfo.Name
                        countLeft += 1
                    ElseIf fInfo.Name.IndexOf("right") > -1 Or fInfo.Name.IndexOf("Right") > -1 Then
                        ReDim Preserve arrayFileRightName(countRight)
                        arrayFileRightName(countRight) = fInfo.Name
                        countRight += 1
                    Else
                        ReDim Preserve arrayFileLeftName(countLeft)
                        arrayFileLeftName(countLeft) = fInfo.Name
                        countLeft += 1
                        ReDim Preserve arrayFileRightName(countRight)
                        arrayFileRightName(countRight) = fInfo.Name
                        countRight += 1
                    End If
                Next
                CB_LeftPostcardTLC.DataSource = arrayFileLeftName
                CB_RightPostcardTLC.DataSource = arrayFileRightName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Подготовка
        putchPreparation = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Подготовка\"
        If Directory.Exists(putchPreparation) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchPreparation, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_PreparationTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Ростверк
        putchGrillage = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Ростверк\"
        If Directory.Exists(putchGrillage) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchGrillage, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_GrillageTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Стойки
        putchRack = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Стойки\"
        If Directory.Exists(putchRack) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchRack, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_RackTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем tlc для Сваи
        putchPile = generalDir & "\TopomaticRobur\DesignBridge\Pillars\Сваи\"
        If Directory.Exists(putchPile) = True Then
            Dim arrayTLCFiles As String() = Directory.GetFiles(putchPile, "*.tlc")
            If IsArray(arrayTLCFiles) = True Then
                Dim arrayFileName As String() = {}
                For i As Integer = 0 To arrayTLCFiles.Length - 1
                    ReDim Preserve arrayFileName(i)
                    Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                    arrayFileName(i) = fInfo.Name
                Next
                CB_PileTLC.DataSource = arrayFileName
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем датагрид с подферменником
        DGV_SubFermenter.Rows.Add(4)
        DGV_SubFermenter.Rows(0).Cells(0).Value = "Номер ряда"
        DGV_SubFermenter.Rows(0).Cells(1).Value = 1
        DGV_SubFermenter.Rows(0).Tag = "numberRow"

        DGV_SubFermenter.Rows(1).Cells(0).Value = "Длина (вдоль насадки), м"
        DGV_SubFermenter.Rows(1).Cells(1).Value = 0.5
        DGV_SubFermenter.Rows(1).Tag = "lenght"

        DGV_SubFermenter.Rows(2).Cells(0).Value = "Ширина (поперек насадки), м"
        DGV_SubFermenter.Rows(2).Cells(1).Value = 0.5
        DGV_SubFermenter.Rows(2).Tag = "width"

        DGV_SubFermenter.Rows(3).Cells(0).Value = "Просвет с балкой, м"
        DGV_SubFermenter.Rows(3).Cells(1).Value = 0.1
        DGV_SubFermenter.Rows(3).Tag = "deltaHeightBeam"

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim listNumberPillar As List(Of String) = New List(Of String) From {""}
        CB_NumberPillar.DataSource = listNumberPillar
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем насадку
    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_NozzleTLC.SelectedIndexChanged
        Dim nameFileNozzle As String = putchNozzle & CB_NozzleTLC.Text.Trim
        If File.Exists(nameFileNozzle) = True Then
            docNozzleTLC.LoadFromFile(nameFileNozzle)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docNozzleTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_Nozzle.RowCount > 1 Then
                    DGV_Nozzle.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_Nozzle.Rows.Add(1)
                    DGV_Nozzle.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_Nozzle.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_Nozzle.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем шкафную стенку
    Private Sub ComboBox21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_CabinetWallTLC.SelectedIndexChanged
        Dim nameFileCabinetWall As String = putchCabinetWall & CB_CabinetWallTLC.Text
        If File.Exists(nameFileCabinetWall) = True Then
            docCabinetWallTLC.LoadFromFile(nameFileCabinetWall)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docCabinetWallTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_CabinetWall.RowCount > 1 Then
                    DGV_CabinetWall.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_CabinetWall.Rows.Add(1)
                    DGV_CabinetWall.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_CabinetWall.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_CabinetWall.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Откосные крылья
    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_LeftHandTLC.SelectedIndexChanged
        Dim nameFileHand As String = putchHand & CB_LeftHandTLC.Text
        If File.Exists(nameFileHand) = True Then
            docLeftHandTLC.LoadFromFile(nameFileHand)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docLeftHandTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_LeftHand.RowCount > 1 Then
                    DGV_LeftHand.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_LeftHand.Rows.Add(1)
                    DGV_LeftHand.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_LeftHand.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_LeftHand.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_RightHandTLC.SelectedIndexChanged
        Dim nameFileHand As String = putchHand & CB_RightHandTLC.Text
        If File.Exists(nameFileHand) = True Then
            docRightHandTLC.LoadFromFile(nameFileHand)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docRightHandTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_RightHand.RowCount > 1 Then
                    DGV_RightHand.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_RightHand.Rows.Add(1)
                    DGV_RightHand.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_RightHand.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_RightHand.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Открылки
    Private Sub ComboBox29_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_LeftPostcardTLC.SelectedIndexChanged
        Dim nameFilePostcard As String = putchPostcard & CB_LeftPostcardTLC.Text
        If File.Exists(nameFilePostcard) = True Then
            docLeftPostcardTLC.LoadFromFile(nameFilePostcard)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docLeftPostcardTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_LeftPostcard.RowCount > 1 Then
                    DGV_LeftPostcard.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_LeftPostcard.Rows.Add(1)
                    DGV_LeftPostcard.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_LeftPostcard.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_LeftPostcard.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_RightPostcardTLC.SelectedIndexChanged
        Dim nameFilePostcard As String = putchPostcard & CB_RightPostcardTLC.Text
        If File.Exists(nameFilePostcard) = True Then
            docRightPostcardTLC.LoadFromFile(nameFilePostcard)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docRightPostcardTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_RightPostcard.RowCount > 1 Then
                    DGV_RightPostcard.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_RightPostcard.Rows.Add(1)
                    DGV_RightPostcard.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_RightPostcard.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_RightPostcard.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Стойки
    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_RackTLC.SelectedIndexChanged
        Dim nameFileRack As String = putchRack & CB_RackTLC.Text
        If File.Exists(nameFileRack) = True Then
            docRackTLC.LoadFromFile(nameFileRack)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docRackTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_Rack.RowCount > 1 Then
                    DGV_Rack.Rows.Clear()
                End If
                Dim countRack As Integer = NUpD_CountRack.Value
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_Rack.Rows.Add(1)
                    DGV_Rack.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_Rack.Rows(countrows).Tag = generalPropLevel.Tag
                    For j As Integer = 1 To countRack
                        DGV_Rack.Rows(countrows).Cells(j).Value = generalPropLevel.Value
                    Next j
                    countrows += 1
                Next i
                For i As Integer = 0 To DGV_Rack.RowCount - 1
                    Dim tagRow As String = DGV_Rack.Rows(i).Tag
                    If tagRow Like "number" Then
                        For j As Integer = 1 To DGV_Rack.ColumnCount - 1
                            DGV_Rack.Rows(i).Cells(j).Value = j
                        Next j
                    End If
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Ростверк
    Private Sub ComboBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_GrillageTLC.SelectedIndexChanged
        Dim nameFileGrillage As String = putchGrillage & CB_GrillageTLC.Text
        If File.Exists(nameFileGrillage) = True Then
            docGrillageTLC.LoadFromFile(nameFileGrillage)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docGrillageTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_Grillage.RowCount > 1 Then
                    DGV_Grillage.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_Grillage.Rows.Add(1)
                    DGV_Grillage.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_Grillage.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_Grillage.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Подготовка
    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_PreparationTLC.SelectedIndexChanged
        Dim nameFilePreparation As String = putchPreparation & CB_PreparationTLC.Text
        If File.Exists(nameFilePreparation) = True Then
            docPreparationTLC.LoadFromFile(nameFilePreparation)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docPreparationTLC, el)
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            If generalPropertiesObject.Count > 0 Then
                If DGV_Preparation.RowCount > 1 Then
                    DGV_Preparation.Rows.Clear()
                End If
                Dim countrows As Integer = 0
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    Dim nameProperties As String = generalPropLevel.Name
                    If nameProperties.IndexOf("|") > -1 Then
                        Continue For
                    End If
                    DGV_Preparation.Rows.Add(1)
                    DGV_Preparation.Rows(countrows).Cells(0).Value = generalPropLevel.Name
                    DGV_Preparation.Rows(countrows).Cells(1).Value = generalPropLevel.Value
                    DGV_Preparation.Rows(countrows).Tag = generalPropLevel.Tag
                    countrows += 1
                Next i
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читаем Сваи
    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_PileTLC.SelectedIndexChanged
        Dim nameFilePile As String = putchPile & CB_PileTLC.Text
        If File.Exists(nameFilePile) = True Then
            docPileTLC.LoadFromFile(nameFilePile)
            Dim el As ImProperties = New ImProperties()
            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(docPileTLC, el)
            typePile = element.Name
            Dim generalPropertiesObject As ImProperties = element.GetProperties()
            Dim count As Integer = 0
            If generalPropertiesObject.Count > 0 Then
                If DGV_Piles.RowCount > 1 Then
                    DGV_Piles.Rows.Clear()
                End If
                DGV_Piles.Rows.Add(generalPropertiesObject.Count)
                For i As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel As ImProperty = generalPropertiesObject.Item(i)
                    If generalPropLevel.Tag Like "type" Then
                        Dim iw As EnumerationPropertyInfo = generalPropLevel.Info
                        Dim dict As Dictionary(Of String, String) = iw.Values
                        Dim arrayType As String() = {}
                        ReDim arrayType(dict.Count - 1)
                        For j As Integer = 0 To dict.Count - 1
                            arrayType(j) = dict.ElementAt(j).Value
                        Next j
                        Continue For
                    End If
                    DGV_Piles.Rows(count).Cells(0).Value = generalPropLevel.Name
                    DGV_Piles.Rows(count).Cells(1).Value = generalPropLevel.Value
                    DGV_Piles.Rows(count).Tag = generalPropLevel.Tag
                    For j As Integer = 1 To DGV_Piles.ColumnCount - 1
                        DGV_Piles.Rows(count).Cells(j).Value = generalPropLevel.Value
                    Next j
                    count += 1
                Next i
                For i As Integer = 0 To DGV_Piles.RowCount - 1
                    Dim tagRow As String = DGV_Piles.Rows(i).Tag
                    If tagRow Like "numberColl" Then
                        For j As Integer = 1 To DGV_Piles.ColumnCount - 1
                            DGV_Piles.Rows(i).Cells(j).Value = j
                        Next j
                    End If
                Next i
                For i As Integer = 0 To DGV_Piles.Rows.Count - 1
                    Dim tagRows As String = DGV_Piles.Rows(i).Tag
                    If tagRows Like "bridge_piles_ushir" Then
                        If DGV_Piles.Rows(i).Cells(1).Value = "Да" Then
                            ChB_PileExpand.Checked = True
                        Else
                            ChB_PileExpand.Checked = False
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'сменить проект arr
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesArrProject.SelectedIndexChanged
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
                        CB_NumberPillar.Text = ""
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор сооружения
    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBox_ListNamesBridge.SelectedIndexChanged
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор опоры
    Private Sub ComboBox15_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_NumberPillar.SelectedIndexChanged
        Dim strNumberPillar As String = CB_NumberPillar.Text
        If IsNothing(strNumberPillar) = True Then Exit Sub
        If strNumberPillar.Trim.Length = 0 Then Exit Sub
        If IsNothing(idBridge) = True Then
            Exit Sub
        End If
        If idBridge.Trim.Length < 2 Then
            Exit Sub
        End If

        If strNumberPillar.Trim.Length > 0 Then
            If IsNumeric(strNumberPillar) = True Then
                numberPillar = Val(strNumberPillar)
                If numberPillar <= 0 Then Exit Sub
                If IsNothing(userBridge) = False Then
                    If numberPillar = 1 Then
                        numberProlet = 1
                    Else
                        numberProlet = userBridge.ProletCount
                    End If
                End If

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
                    If dictionaryBridgeElements.Count > 0 Then
                        'ищем ось опоры
                        dataPillar = Pillar.getAxisPillar(dictionaryBridgeElements, numberPillar)
                        If IsNothing(dataPillar) = True Then
                            Exit Sub
                        End If
                        userAxisPillar = dataPillar.getPillar()
                        axisLinePillar = dataPillar.DWGEntity
                        If IsNothing(userAxisPillar) = True Then Exit Sub
                        If IsNothing(axisLinePillar) = True Then Exit Sub
                        If axisLinePillar.Length = 0 Then Exit Sub
                        'фиксированная высота шкафной стенки
                        If userAxisPillar.FixedHeightCabinetWall = True Then
                            ChB_FixedHeightCabinetWall.Checked = True
                        End If
                        'наличие стоек
                        If userAxisPillar.PresenceRacks = True Then
                            ChB_CreateRack.Checked = False
                        End If
                        'наличие ростверка
                        If userAxisPillar.PresencGrillage = True Then
                            ChB_CreateGrillage.Checked = False
                        End If
                        'наличие подготовки
                        If userAxisPillar.PresencPreparation = True Then
                            ChB_CreatePreparation.Checked = False
                        End If
                        'вставить сваю в стойку
                        If userAxisPillar.PileInRack = True Then
                            ChB_InsertPileInRack.Checked = True
                        End If
                        'отметка земли
                        If userAxisPillar.ElevationLand <> 0 Then
                            ChB_ElevationLand.Checked = True
                            NUpD_ElevationLand.Value = userAxisPillar.ElevationLand
                        End If
                        '====================================================================================================
                        ' ищем уже существующую насадку

                        Dim listDataNozzle As List(Of StructureElement) = NozzlePillar.getAxis(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        dataNozzle = StructureElement.isValidateDataStructure(listDataNozzle)
                        If IsNothing(dataNozzle) = False Then
                            userNozzle = dataNozzle.getNozzlePillar()
                            axisLineNozzle = dataNozzle.DWGEntity
                            Dim nameTLC As String = userNozzle.NameModel
                            userNozzle.writePropertiesNozzle(DGV_Nozzle)
                            If CB_NozzleTLC.Items.Contains(userNozzle.NameModel & ".tlc") Then
                                CB_NozzleTLC.Text = userNozzle.NameModel & ".tlc"
                            End If
                            If userNozzle.PileRowsDiagram.Trim.Length > 0 Then
                                TxtB_PileRowDiagram.Text = userNozzle.PileRowsDiagram
                            End If
                            If userNozzle.PileColumnDiagram.Trim.Length > 0 Then
                                TxtB_PileCollDiagram.Text = userNozzle.PileColumnDiagram
                            End If
                        Else
                            'создаем новую насадку
                            'dataNozzle = NozzlePillar.createAxisNozzlePillar(idBridge)
                        End If
                        '===================================================================================================================
                        'заполняем подферменники
                        'находим балки для выбранной опоры (0-предыдущий пролет, 1-следующий пролет)
                        Dim listBeamsPillar As List(Of Dictionary(Of Integer, StructureElement)) = Pillar.getBeamsPillarByNumber(numberPillar, dictionaryBridgeElements)
                        If listBeamsPillar.Count = 0 Then
                            MsgBox("Не удалось найти балки для выбранного пролета мостового сооружения!!!")
                        Else
                            Dim indexColumn As Integer = 1
                            For i As Integer = 0 To listBeamsPillar.Count - 1
                                Dim dictSubFerm As Dictionary(Of Integer, StructureElement) = listBeamsPillar.Item(i)
                                If IsNothing(dictSubFerm) = True Then Continue For
                                If dictSubFerm.Count = 0 Then Continue For
                                For j As Integer = 0 To dictSubFerm.Count - 1
                                    Dim dataBeam As StructureElement = dictSubFerm.ElementAt(j).Value
                                    If IsNothing(dataBeam) = True Then Continue For
                                    Dim userBeam As BeamI = dataBeam.getBeamI()
                                    If IsNothing(userBeam) = True Then Continue For
                                    Dim numberProlet As Integer = userBeam.numberProlet
                                    Dim numberRow As Integer = userBeam.numberRow
                                    Dim listdataSubFerm As List(Of StructureElement) = SubFermenters.getAxis(dictionaryBridgeElements, numberPillar, numberProlet, numberRow, numberSubPillar)
                                    Dim dataSubFerm As StructureElement = StructureElement.isValidateDataStructure(listdataSubFerm)
                                    Dim userSubFerm As SubFermenters = Nothing
                                    If IsNothing(dataSubFerm) = False Then
                                        userSubFerm = dataSubFerm.getSubFermenters()
                                    Else
                                        userSubFerm = New SubFermenters()
                                        userSubFerm.NumberPillar = numberPillar
                                        userSubFerm.NumberProlet = numberProlet
                                        userSubFerm.NumberRow = numberRow
                                    End If
                                    If DGV_SubFermenter.Columns.Count <= indexColumn Then
                                        Dim index = DGV_SubFermenter.Columns.Add("SubFermenters", "Значение")
                                        DGV_SubFermenter.Columns.Item(index).Width = 70
                                    End If
                                    userSubFerm.writePropertiesSubFermenters(DGV_SubFermenter, indexColumn, True)
                                    indexColumn += 1
                                Next j
                            Next i
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим шкафную стенку
                        Dim listdataCabinetWall As List(Of StructureElement) = CabinetWallPillar.getAxis(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        dataCabinetWall = StructureElement.isValidateDataStructure(listdataCabinetWall)
                        If IsNothing(dataCabinetWall) = False Then
                            userCabinetWall = dataCabinetWall.getCabinetWallPillar()
                            axisLineCabinetWall = dataCabinetWall.DWGEntity
                            userCabinetWall.writePropertiesCabinetWall(DGV_CabinetWall)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим левое и правое откосное крыло
                        Dim listDataLeftHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, numberSubPillar, Pillar.SidePillarElement.Left)
                        dataLeftHand = StructureElement.isValidateDataStructure(listDataLeftHand)
                        If IsNothing(dataLeftHand) = False Then
                            userLeftHand = dataLeftHand.getHandPillar
                            axisLineLeftHand = dataLeftHand.DWGEntity
                            userLeftHand.writePropertiesHand(DGV_LeftHand)
                        End If
                        Dim listDataRightHand As List(Of StructureElement) = HandPillar.getHandPillar(dictionaryBridgeElements, numberPillar, numberSubPillar, Pillar.SidePillarElement.Right)
                        dataRightHand = StructureElement.isValidateDataStructure(listDataLeftHand)
                        If IsNothing(dataRightHand) = False Then
                            userRightHand = dataRightHand.getHandPillar
                            axisLineRightHand = dataRightHand.DWGEntity
                            userRightHand.writePropertiesHand(DGV_RightHand)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим левый и правый обратный открылок
                        Dim listDataLeftPostcard As List(Of StructureElement) = PostcardPillar.getPoctcardPillar(dictionaryBridgeElements, numberPillar, numberSubPillar, Pillar.SidePillarElement.Left)
                        dataLeftPostcard = StructureElement.isValidateDataStructure(listDataLeftPostcard)
                        If IsNothing(dataLeftPostcard) = False Then
                            userLeftPostcard = dataLeftPostcard.getPostcardPillar
                            axisLineLeftPostcard = dataLeftPostcard.DWGEntity
                            userLeftPostcard.writePropertiesPostcard(DGV_LeftPostcard)
                        End If
                        Dim listDataRightPostcard As List(Of StructureElement) = PostcardPillar.getPoctcardPillar(dictionaryBridgeElements, numberPillar, numberSubPillar, Pillar.SidePillarElement.Right)
                        dataRightPostcard = StructureElement.isValidateDataStructure(listDataRightPostcard)
                        If IsNothing(dataRightPostcard) = False Then
                            userRightPostcard = dataRightPostcard.getPostcardPillar
                            axisLineRightPostcard = dataRightPostcard.DWGEntity
                            userRightPostcard.writePropertiesPostcard(DGV_RightPostcard)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим стойки
                        dictionaryDataRack = RackPillar.getRackPillar(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        If dictionaryUserRack.Count > 0 Then
                            Dim indexColumn As Integer = 1
                            For i As Integer = 0 To dictionaryUserRack.Count - 1
                                Dim dataRack As StructureElement = dictionaryDataRack.ElementAt(i).Value
                                Dim userRack As RackPillar = dataRack.getRackPillar
                                Dim numberRack As Integer = userRack.Number
                                If i = 0 Then
                                    Dim nameModel As String = userRack.Model
                                    CB_RackTLC.Text = nameModel
                                End If
                                If numberRack > DGV_Rack.ColumnCount + 1 Then
                                    NUpD_CountRack.Value = numberRack
                                End If
                                userRack.writePropertiesRack(DGV_Rack)
                            Next i
                            'уменьшаем количество стоек
                            If NUpD_CountRack.Value > dictionaryUserRack.Count Then
                                For i = NUpD_CountRack.Value To dictionaryUserRack.Count Step -1
                                    NUpD_CountRack.Value = i
                                Next i
                            End If
                            ChB_CreateRack.Checked = True
                            ChB_CreateGrillage.Checked = True
                        Else
                            ChB_CreateRack.Checked = False
                            ChB_CreateGrillage.Checked = False
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим ростверк
                        dataGrillage = GrillagePillar.getGrillagePillar(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        If IsNothing(dataGrillage) = False Then
                            userGrillage = dataGrillage.getGrillagePillar()
                            axisLineGrillage = dataGrillage.DWGEntity
                            userGrillage.writePropertiesGrillage(DGV_Grillage)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим подготовку
                        dataPreparation = PreparationPillar.getPreparationPillar(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        If IsNothing(dataPreparation) = False Then
                            userPreparation = dataPreparation.getPreparationPillar()
                            axisLinePreparation = dataPreparation.DWGEntity
                            userPreparation.writePropertiesPreparation(DGV_Preparation)
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'находим сваи
                        dictDataPile = PilePillar.getPilePillar(dictionaryBridgeElements, numberPillar, numberSubPillar)
                        If dictUserPile.Count > 0 Then
                            Dim indexPile As Integer = 1
                            For i As Integer = 0 To dictUserPile.Count - 1
                                Dim columnDataPile As Dictionary(Of Integer, StructureElement) = dictDataPile.ElementAt(i).Value
                                If columnDataPile.Count > 0 Then
                                    For j As Integer = 0 To columnDataPile.Count - 1
                                        Dim dataPile As StructureElement = columnDataPile.ElementAt(j).Value
                                        Dim userPile As PilePillar = dataPile.getPilePillar()
                                        If i = 0 And j = 0 Then
                                            Dim modelPile As String = userPile.NameModel
                                            CB_PileTLC.Text = modelPile
                                            NUpD_CountColumnsPile.Maximum = columnDataPile.Count
                                        End If
                                        If DGV_Piles.Columns.Count < indexPile Then
                                            DGV_Piles.Columns.Add("Pile", "")
                                        End If
                                        userPile.writePropertiesPile(DGV_Piles, indexPile)
                                        indexPile += 1
                                    Next j
                                End If
                            Next i
                            NUpD_CountRowsPile.Minimum = 1
                            NUpD_CountRowsPile.Maximum = dictUserPile.Count
                            NUpD_CountColumnsPile.Minimum = 1
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'библиотека крайних опор (записать опору в библиотеку)
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        formUserTempl.templateDir = ComboBox11.Tag
        If formUserTempl.ListBox1.Items.Count > 0 Then
            formUserTempl.ListBox1.Items.Clear()
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim oldArray As String() = ComboBox11.DataSource
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
            formUserTempl.Label3.Text = ComboBox11.Tag
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем новую форму
        formUserTempl.ShowDialog()
        If formUserTempl.boolWriteFile = True Then 'запись разрешена
            Dim nameFile As String = formUserTempl.TextBox1.Text
            Dim FileNameGeo As String = ComboBox11.Tag & nameFile & ".xml"
            If FileNameGeo.Trim.Length > 0 Then
                Try
                    Using xw As XmlWriter = XmlWriter.Create(FileNameGeo)
                        xw.WriteStartDocument()
                        xw.WriteStartElement("LastPillarsBridge")
                        xw.WriteAttributeString("Ver", "Версия 1")

                        xw.WriteStartElement("GeneralProperties")
                        xw.WriteAttributeString("nameAlignment", CB_NameAlignment.Text)
                        xw.WriteAttributeString("nameProjectSurface", CB_ProjectSurface.Text)
                        xw.WriteAttributeString("ProjectSurfaceInAlign", ChB_ProjectSurfaceInAlignment.Checked)
                        xw.WriteAttributeString("nameEgSurface", CB_EgSurface.Text)
                        xw.WriteAttributeString("numberPillar", CB_NumberPillar.Text)
                        xw.WriteAttributeString("elevationLand", NUpD_ElevationLand.Value)
                        xw.WriteEndElement()
                        xw.WriteStartElement("Properties")
                        For Each control As Control In Me.Controls
                            If TypeOf control Is System.Windows.Forms.TabControl Then
                                Dim tabUserControl As System.Windows.Forms.TabControl = control
                                For Each tabSelectControl As System.Windows.Forms.Control In tabUserControl.Controls
                                    If TypeOf tabSelectControl Is System.Windows.Forms.TabPage Then
                                        Dim userTabPage As System.Windows.Forms.TabPage = tabSelectControl
                                        For Each userControl As System.Windows.Forms.Control In userTabPage.Controls
                                            If TypeOf userControl Is System.Windows.Forms.DataGridView Then
                                                xw.WriteStartElement(userControl.Name)
                                                Dim userDataGrid As DataGridView = userControl
                                                If userDataGrid.RowCount > 1 Then
                                                    If userDataGrid.Name Like "DGV_Rack" Or userDataGrid.Name Like "DGV_Piles" Or userDataGrid.Name Like "DGV_SubFermenter" Then
                                                        For i As Integer = 0 To userDataGrid.RowCount - 1
                                                            Dim nameField As String = userDataGrid.Rows(i).Tag
                                                            Dim value As String = userDataGrid.Rows(i).Cells(1).Value
                                                            If userDataGrid.ColumnCount > 2 Then
                                                                If nameField Like "numberRow" And userDataGrid.Name Like "DGV_SubFermenter" Then
                                                                    For j As Integer = 1 To userDataGrid.ColumnCount - 1
                                                                        If j = 1 Then
                                                                            value = userDataGrid.Columns.Item(j).Tag
                                                                        Else
                                                                            value = value & ";" & userDataGrid.Columns.Item(j).Tag
                                                                        End If
                                                                    Next j
                                                                Else
                                                                    For j As Integer = 2 To userDataGrid.ColumnCount - 1
                                                                        value = value & ";" & userDataGrid.Rows(i).Cells(j).Value
                                                                    Next j
                                                                End If
                                                            End If
                                                            If IsNothing(nameField) = True Then Continue For
                                                            If nameField.Trim.Length = 0 Then Continue For
                                                            If IsNothing(value) = True Then value = ""
                                                            xw.WriteAttributeString(nameField, value)
                                                        Next
                                                    Else
                                                        For i As Integer = 0 To userDataGrid.RowCount - 1
                                                            Dim nameField As String = userDataGrid.Rows(i).Tag
                                                            Dim value As String = userDataGrid.Rows(i).Cells(1).Value
                                                            If IsNothing(nameField) = True Then Continue For
                                                            If nameField.Trim.Length = 0 Then Continue For
                                                            If IsNothing(value) = True Then value = ""
                                                            xw.WriteAttributeString(nameField, value)
                                                        Next
                                                    End If
                                                End If
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is NumericUpDown Then
                                                xw.WriteStartElement(userControl.Name)
                                                Dim num As NumericUpDown = userControl
                                                xw.WriteAttributeString("Field", num.Value)
                                                xw.WriteAttributeString("minZn", num.Minimum)
                                                xw.WriteAttributeString("maxZn", num.Maximum)
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is System.Windows.Forms.ComboBox Then
                                                If userControl.Name Like "ComboBox11" Then Continue For
                                                xw.WriteStartElement(userControl.Name)
                                                Dim combo As System.Windows.Forms.ComboBox = userControl
                                                xw.WriteAttributeString("Field", combo.Text)
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is System.Windows.Forms.CheckBox Then
                                                xw.WriteStartElement(userControl.Name)
                                                Dim cb As System.Windows.Forms.CheckBox = userControl
                                                xw.WriteAttributeString("Field", cb.Checked)
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is System.Windows.Forms.RadioButton Then
                                                xw.WriteStartElement(userControl.Name)
                                                Dim rb As System.Windows.Forms.RadioButton = userControl
                                                xw.WriteAttributeString("Field", rb.Checked)
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is System.Windows.Forms.TextBox Then
                                                xw.WriteStartElement(userControl.Name)
                                                Dim tb As System.Windows.Forms.TextBox = userControl
                                                xw.WriteAttributeString("Field", tb.Text)
                                                xw.WriteEndElement()
                                            ElseIf TypeOf userControl Is System.Windows.Forms.TabControl Then
                                                Dim userSubTabControl As System.Windows.Forms.TabControl = userControl
                                                For Each tabSubSelectControl As System.Windows.Forms.Control In userSubTabControl.Controls
                                                    If TypeOf tabSubSelectControl Is System.Windows.Forms.TabPage Then
                                                        Dim userSubTabPage As System.Windows.Forms.TabPage = tabSubSelectControl
                                                        For Each userSubControl As System.Windows.Forms.Control In userSubTabPage.Controls
                                                            If TypeOf userSubControl Is System.Windows.Forms.DataGridView Then
                                                                Dim userSubDataGrid As DataGridView = userSubControl
                                                                xw.WriteStartElement(userSubDataGrid.Name)
                                                                If userSubDataGrid.RowCount > 1 Then
                                                                    For i As Integer = 0 To userSubDataGrid.RowCount - 1
                                                                        Dim nameField As String = userSubDataGrid.Rows(i).Tag
                                                                        Dim value As String = userSubDataGrid.Rows(i).Cells(1).Value
                                                                        If IsNothing(nameField) = True Then Continue For
                                                                        If nameField.Trim.Length = 0 Then Continue For
                                                                        If IsNothing(value) = True Then value = ""
                                                                        xw.WriteAttributeString(nameField, value)
                                                                    Next
                                                                End If
                                                                xw.WriteEndElement()
                                                            ElseIf TypeOf userSubControl Is NumericUpDown Then
                                                                xw.WriteStartElement(userSubControl.Name)
                                                                Dim num As NumericUpDown = userSubControl
                                                                xw.WriteAttributeString("Field", num.Value)
                                                                xw.WriteEndElement()
                                                            ElseIf TypeOf userSubControl Is System.Windows.Forms.ComboBox Then
                                                                xw.WriteStartElement(userSubControl.Name)
                                                                Dim combo As System.Windows.Forms.ComboBox = userSubControl
                                                                xw.WriteAttributeString("Field", combo.Text)
                                                                xw.WriteEndElement()
                                                            ElseIf TypeOf userSubControl Is System.Windows.Forms.CheckBox Then
                                                                xw.WriteStartElement(userSubControl.Name)
                                                                Dim cb As System.Windows.Forms.CheckBox = userSubControl
                                                                xw.WriteAttributeString("Field", cb.Checked)
                                                                xw.WriteEndElement()
                                                            ElseIf TypeOf userSubControl Is System.Windows.Forms.RadioButton Then
                                                                xw.WriteStartElement(userSubControl.Name)
                                                                Dim rb As System.Windows.Forms.RadioButton = userSubControl
                                                                xw.WriteAttributeString("Field", rb.Checked)
                                                                xw.WriteEndElement()
                                                            ElseIf TypeOf userSubControl Is System.Windows.Forms.TextBox Then
                                                                xw.WriteStartElement(userSubControl.Name)
                                                                Dim tb As System.Windows.Forms.TextBox = userSubControl
                                                                xw.WriteAttributeString("Field", tb.Text)
                                                                xw.WriteEndElement()
                                                            End If
                                                        Next
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                        xw.WriteEndElement()
                        xw.WriteEndDocument()
                    End Using
                    MsgBox("Текущая схема успешно сохранена!")
                    Dim boolFindSheme As Integer = -1
                    Dim arrayPillarSheme As String() = ComboBox11.DataSource
                    If IsArray(arrayPillarSheme) = True Then
                        boolFindSheme = MathFunction.FuncFindValueToFArray(nameFile, arrayPillarSheme)
                    End If
                    If boolFindSheme = -1 Then
                        ReDim Preserve arrayPillarSheme(arrayPillarSheme.Length)
                        arrayPillarSheme(arrayPillarSheme.Length - 1) = nameFile
                        ComboBox11.DataSource = arrayPillarSheme
                        ComboBox11.Text = nameFile
                    End If
                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try
            Else
                'создаем новый каталог
                Try
                    Directory.CreateDirectory(generalDir & "\TopomaticRobur\DesignBridge\UserProperties\MiddlePillarsSchema\Rigels\")
                Catch ex As Exception
                    MsgBox("Не удалось создать каталог: " & generalDir & "\TopomaticRobur\DesignBridge\UserProperties\MiddlePillarsSchema\Rigels\")
                End Try
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'загрузка шаблона опоры из файла xml
    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        Dim nameFileXml As String = ComboBox11.Text
        Dim directorySheme As String = ComboBox11.Tag
        Dim putchShemePillar As String = directorySheme & nameFileXml
        If File.Exists(putchShemePillar) = True Then
            Using reader As XmlReader = XmlReader.Create(putchShemePillar)
                While reader.Read()
                    If reader.NodeType = XmlNodeType.Element Then
                        Dim nameElement As String = reader.Name
                        If reader.HasAttributes = True Then
                            While (reader.MoveToNextAttribute())
                                Dim name As String = reader.Name
                                Dim valN As String = reader.Value
                                If nameElement Like "GeneralProperties" And name Like "nameAlignment" Then
                                    Dim listElements As List(Of String) = CB_NameAlignment.DataSource
                                    If IsNothing(listElements) = False Then
                                        If listElements.Count > 0 Then
                                            Dim index As Integer = listElements.IndexOf(valN)
                                            If index > -1 Then
                                                CB_NameAlignment.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "GeneralProperties" And name Like "nameProjectSurface" Then
                                    Dim listElements As List(Of String) = CB_ProjectSurface.DataSource
                                    If IsNothing(listElements) = False Then
                                        If listElements.Count > 0 Then
                                            Dim ind As Integer = listElements.IndexOf(valN)
                                            If ind > -1 Then
                                                CB_ProjectSurface.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "GeneralProperties" And name Like "nameEgSurface" Then
                                    Dim listElements As List(Of String) = CB_EgSurface.DataSource
                                    If IsNothing(listElements) = False Then
                                        If listElements.Count > 0 Then
                                            Dim ind As Integer = listElements.IndexOf(valN)
                                            If ind > -1 Then
                                                CB_EgSurface.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "GeneralProperties" And name Like "numberPillar" Then
                                    Dim listElements As List(Of String) = CB_NumberPillar.DataSource
                                    If IsNothing(listElements) = False Then
                                        If listElements.Count > 0 Then
                                            Dim ind As Integer = listElements.IndexOf(valN)
                                            If ind > -1 Then
                                                CB_NumberPillar.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "GeneralProperties" And name Like "ProjectSurfaceInAlign" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_ProjectSurfaceInAlignment.Checked = boolZn
                                ElseIf nameElement Like "GeneralProperties" And name Like "elevationLand" Then
                                    Dim elevLand As Double = CDbl(valN)
                                    If elevLand <> 0 Then
                                        ChB_ElevationLand.Checked = True
                                        NUpD_ElevationLand.Value = elevLand
                                    End If
                                ElseIf nameElement Like "CB_NozzleTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_NozzleTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_NozzleTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_CabinetWallTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_CabinetWallTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_CabinetWallTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_LeftHandTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_LeftHandTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_LeftHandTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_RightHandTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_RightHandTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_RightHandTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_LeftPostcardTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_LeftPostcardTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_LeftPostcardTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_RightPostcardTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_RightPostcardTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_RightPostcardTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_RackTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_RackTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_RackTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_GrillageTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_GrillageTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_GrillageTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_PreparationTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_PreparationTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_PreparationTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "CB_PileTLC" And name Like "Field" Then
                                    Dim arrayElements As String() = CB_PileTLC.DataSource
                                    If IsNothing(arrayElements) = False Then
                                        If IsArray(arrayElements) = True Then
                                            Dim ind As Integer = MathFunction.FuncFindValueToFArray(valN, arrayElements)
                                            If ind > -1 Then
                                                CB_PileTLC.Text = valN
                                            End If
                                        End If
                                    End If
                                ElseIf nameElement Like "ChB_FixedHeightCabinetWall" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_FixedHeightCabinetWall.Checked = boolZn
                                ElseIf nameElement Like "ChB_CreateRack" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_CreateRack.Checked = boolZn
                                ElseIf nameElement Like "ChB_EgeParallel" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_EgeParallel.Checked = boolZn
                                ElseIf nameElement Like "ChB_fixedHeightRack" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_fixedHeightRack.Checked = boolZn
                                ElseIf nameElement Like "ChB_CreateGrillage" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_CreateGrillage.Checked = boolZn
                                ElseIf nameElement Like "ChB_GrillageEge" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_GrillageEge.Checked = boolZn
                                ElseIf nameElement Like "ChB_CreatePreparation" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_CreatePreparation.Checked = boolZn
                                ElseIf nameElement Like "ChB_PileExpand" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_PileExpand.Checked = boolZn
                                ElseIf nameElement Like "ChB_InsertPileInRack" And name Like "Field" Then
                                    Dim boolZn As Boolean = CBool(valN)
                                    ChB_InsertPileInRack.Checked = boolZn
                                ElseIf nameElement Like "TxtB_PileRowDiagram" And name Like "Field" Then
                                    TxtB_PileRowDiagram.Text = valN
                                ElseIf nameElement Like "TxtB_PileCollDiagram" And name Like "Field" Then
                                    TxtB_PileCollDiagram.Text = valN
                                ElseIf nameElement Like "NUpD_CountRack" And name Like "Field" Then
                                    Dim countRack As Integer = CInt(valN)
                                    If countRack > 1 Then
                                        For i As Integer = 2 To countRack
                                            NUpD_CountRack.Value = i
                                        Next
                                    End If
                                ElseIf nameElement Like "NUpD_CountRowsPile" And name Like "Field" Then
                                    Dim countRows As Integer = CInt(valN)
                                    If countRows > 0 Then
                                        NUpD_CountRowsPile.Maximum = countRows
                                    End If
                                ElseIf nameElement Like "NUpD_CountColumnsPile" And name Like "Field" Then
                                    Dim countColls As Integer = CInt(valN)
                                    If countColls > 0 Then
                                        NUpD_CountColumnsPile.Maximum = countColls
                                    End If
                                End If
                            End While
                        End If
                    End If
                End While
            End Using
            Using reader As XmlReader = XmlReader.Create(putchShemePillar)
                While reader.Read()
                    If reader.NodeType = XmlNodeType.Element Then
                        Dim nameElement As String = reader.Name
                        If reader.HasAttributes = True Then
                            While (reader.MoveToNextAttribute())
                                Dim name As String = reader.Name
                                Dim valN As String = reader.Value
                                If nameElement Like "DGV_Nozzle" Then
                                    For i As Integer = 0 To DGV_Nozzle.Rows.Count - 1
                                        Dim tagRow As String = DGV_Nozzle.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_Nozzle.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_CabinetWall" Then
                                    For i As Integer = 0 To DGV_CabinetWall.Rows.Count - 1
                                        Dim tagRow As String = DGV_CabinetWall.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_CabinetWall.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_SubFermenter" Then
                                    For i As Integer = 0 To DGV_SubFermenter.Rows.Count - 1
                                        Dim tagRow As String = DGV_SubFermenter.Rows(i).Tag
                                        If tagRow Like name Then
                                            Dim arrayValue As String() = valN.Split(";")
                                            If arrayValue.Length > 1 Then
                                                For j As Integer = 0 To arrayValue.Length - 1
                                                    Dim countColumn As Integer = j + 1
                                                    If countColumn > DGV_SubFermenter.ColumnCount - 1 Then
                                                        DGV_SubFermenter.Columns.Add(countColumn, "Значение")
                                                        DGV_SubFermenter.Columns.Item(countColumn).Width = 70
                                                    End If
                                                    If countColumn <= DGV_SubFermenter.ColumnCount - 1 Then
                                                        If name Like "numberRow" Then
                                                            DGV_SubFermenter.Columns.Item(countColumn).Tag = Val(arrayValue(j))
                                                            DGV_SubFermenter.Rows(i).Cells(countColumn).Value = FuncFormatZn.getConditionalRow(Val(arrayValue(j)))
                                                        Else
                                                            DGV_SubFermenter.Rows(i).Cells(countColumn).Value = arrayValue(j)
                                                        End If
                                                    End If
                                                Next j
                                            End If
                                            Exit For
                                        End If
                                    Next i
                                ElseIf nameElement Like "DGV_LeftHand" Then
                                    For i As Integer = 0 To DGV_LeftHand.Rows.Count - 1
                                        Dim tagRow As String = DGV_LeftHand.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_LeftHand.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_RightHand" Then
                                    For i As Integer = 0 To DGV_RightHand.Rows.Count - 1
                                        Dim tagRow As String = DGV_RightHand.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_RightHand.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_Rack" Then
                                    For i As Integer = 0 To DGV_Rack.Rows.Count - 1
                                        Dim tagRow As String = DGV_Rack.Rows(i).Tag
                                        If tagRow Like name Then
                                            Dim arrayValue As String() = valN.Split(";")
                                            If arrayValue.Length > 1 Then
                                                For j As Integer = 0 To arrayValue.Length - 1
                                                    Dim countColumn As Integer = j + 1
                                                    If countColumn <= DGV_Rack.ColumnCount - 1 Then
                                                        DGV_Rack.Rows(i).Cells(countColumn).Value = arrayValue(j)
                                                    End If
                                                Next j
                                            End If
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_Grillage" Then
                                    For i As Integer = 0 To DGV_Grillage.Rows.Count - 1
                                        Dim tagRow As String = DGV_Grillage.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_Grillage.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_Preparation" Then
                                    For i As Integer = 0 To DGV_Preparation.Rows.Count - 1
                                        Dim tagRow As String = DGV_Preparation.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_Preparation.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_Piles" Then
                                    Dim arrayData As String() = valN.Split(";")
                                    If IsArray(arrayData) = True Then
                                        If arrayData.Length > DGV_Piles.Columns.Count - 1 Then
                                            Dim delta As Integer = arrayData.Length - DGV_Piles.Columns.Count
                                            For i As Integer = 0 To delta
                                                Dim dataColl As DataGridViewColumn = DGV_Piles.Columns.Item(1)
                                                Dim newDataColl As DataGridViewColumn = New DataGridViewColumn()
                                                newDataColl = dataColl.Clone()
                                                Dim posRow As Integer = DGV_Piles.Columns.Add(newDataColl)
                                            Next i
                                        End If
                                        For i As Integer = 0 To DGV_Piles.Rows.Count - 1
                                            Dim tagRow As String = DGV_Piles.Rows(i).Tag
                                            If tagRow Like name Then
                                                For j = 0 To arrayData.Length - 1
                                                    DGV_Piles.Rows(i).Cells(j + 1).Value = arrayData(j)
                                                Next
                                            End If
                                        Next
                                    End If
                                ElseIf nameElement Like "DGV_LeftPostcard" Then
                                    For i As Integer = 0 To DGV_LeftPostcard.Rows.Count - 1
                                        Dim tagRow As String = DGV_LeftPostcard.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_LeftPostcard.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                ElseIf nameElement Like "DGV_RightPostcard" Then
                                    For i As Integer = 0 To DGV_RightPostcard.Rows.Count - 1
                                        Dim tagRow As String = DGV_RightPostcard.Rows(i).Tag
                                        If tagRow Like name Then
                                            DGV_RightPostcard.Rows(i).Cells(1).Value = valN
                                            Exit For
                                        End If
                                    Next
                                End If
                            End While
                        End If
                    End If
                End While
            End Using
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'галочка взять поверхность из трассы
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_ProjectSurfaceInAlignment.CheckedChanged
        If ChB_ProjectSurfaceInAlignment.Checked = True Then
            CB_ProjectSurface.Enabled = False
        Else
            CB_ProjectSurface.Enabled = True
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'наличие уширения у сваи
    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_PileExpand.CheckedChanged
        If ChB_PileExpand.Checked = False Then
            Dim boolReadOnly As Boolean = False
            For i As Integer = 0 To DGV_Piles.Rows.Count - 1
                Dim tagRows As String = DGV_Piles.Rows(i).Tag
                If tagRows Like "bridge_piles_ushir" Then
                    DGV_Piles.Rows(i).Cells(1).Value = "Нет"
                    boolReadOnly = True
                End If
                If boolReadOnly = True Then
                    FreezeRow(DGV_Piles.Rows(i))
                End If
            Next
        Else
            Dim boolReadOnly As Boolean = False
            For i As Integer = 0 To DGV_Piles.Rows.Count - 1
                Dim tagRows As String = DGV_Piles.Rows(i).Tag
                If tagRows Like "bridge_piles_ushir" Then
                    DGV_Piles.Rows(i).Cells(1).Value = "Да"
                    boolReadOnly = True
                End If
                If boolReadOnly = True Then
                    UnfreezeRow(DGV_Piles.Rows(i))
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Разморозить строку
    Private Shared Sub FreezeRow(row As DataGridViewRow)
        row.ReadOnly = True
        For Each cell As DataGridViewCell In row.Cells
            cell.Style.BackColor = Color.LightGray
            cell.Style.ForeColor = Color.Gray
            cell.ToolTipText = "Заблокировано для редактирования"
        Next
    End Sub
    'Заморозить строку
    Private Shared Sub UnfreezeRow(row As DataGridViewRow)
        row.ReadOnly = False
        For Each cell As DataGridViewCell In row.Cells
            cell.Style.BackColor = Color.White
            cell.Style.ForeColor = Color.Black
            cell.ToolTipText = ""
        Next
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'расчеты
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка ОТМЕНА
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = False
        Me.Close()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'Предварительный расчет
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'активный проект
        Dim nameBridge As String = CBox_ListNamesBridge.Text
        If IsNothing(activProjectDocument) = True Then
            arrProject = ProjectBridge.getIArrangementModel(nameBridge)
            If IsNothing(arrProject) = False Then
                activProjectDocument = arrProject.Drawing
            End If
        End If
        If IsNothing(activProjectDocument) = True Then
            MsgBox("Выберите подобъект к котором содержится Ваше сооружение!!!")
            Exit Sub
        End If
        'имя сооружения
        Dim nameModel As String = CBox_ListNamesArrProject.Text
        'имя трассы
        Dim nameAlignment As String = CB_NameAlignment.Text
        If IsNothing(projectAlignment) = True Then
            Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlignment, projectAlignment)
            If IsNothing(projectAlignment) = True Then
                MsgBox("Проектная ось трассы автомобильной дороги, не найдена!!!")
                Exit Sub
            End If
        End If
        'фактическая поверхность
        Dim nameEgSurface As String = CB_EgSurface.Text
        egSurface = FuncSurface.getSurfaceByName(nameEgSurface)
        If IsNothing(egSurface) = True And ChB_fixedHeightRack.Checked = False Then
            MsgBox("Поверхность земли не найдена или выключена!!!")
            Exit Sub
        End If
        'проектная поверхность
        Dim nameProjectSurface As String = CB_ProjectSurface.Text
        If ChB_ProjectSurfaceInAlignment.Checked = False Then
            projectSurface = FuncSurface.getSurfaceByName(nameProjectSurface)
        Else
            projectSurface = FuncAlignment.getSurfaceToAlignment(nameAlignment)
        End If
        If IsNothing(projectSurface) = True Then
            MsgBox("Проектная поверхность не найдена или выключена!!!")
            Exit Sub
        End If
        'проектное сооружение
        If IsNothing(userBridge) = True Then
            MsgBox("Сооружение не найдено!!!")
            Exit Sub
        End If
        'номер опоры
        If IsNumeric(CB_NumberPillar.Text) = False Then
            MsgBox("Выберите номер опоры!")
            Exit Sub
        End If
        numberPillar = Val(CB_NumberPillar.Text)
        If numberPillar <= 0 Then
            MsgBox("Не корректно выбран номер опоры.")
            Exit Sub
        End If
        If IsNothing(dataPillar) = True Then
            'получаем заново все элементы сооружения
            dictionaryBridgeElements = userBridge.getBridgeObjects(axisLineBridge)
            userAxisPillar = dataPillar.getPillar()
            axisLinePillar = dataPillar.DWGEntity
        End If
        Dim collDiagram As String = TxtB_PileCollDiagram.Text
        Dim rowDiagram As String = TxtB_PileRowDiagram.Text
        '=====================================================================================================
        'записываем новые данные в опору
        userAxisPillar.SingleSubFarmer = True
        If ChB_ElevationLand.Checked = True Then
            userAxisPillar.ElevationLand = NUpD_ElevationLand.Value
        End If
        userAxisPillar.PresenceRacks = ChB_CreateRack.Checked
        userAxisPillar.PresencGrillage = ChB_CreateGrillage.Checked
        userAxisPillar.PresencPreparation = ChB_CreatePreparation.Checked
        userAxisPillar.PileInRack = ChB_InsertPileInRack.Checked
        '=====================================================================================================
        'читаем насадку
        '======================================================================================================
        'читаем характеристики насадки
        userNozzle = New NozzlePillar
        If ChB_CreateRack.Checked = True Then
            'без опор (из насадки идут сваи)
            userNozzle = NozzlePillar.readPropertiesNozzle(numberPillar, 1, collDiagram, rowDiagram, DGV_Nozzle)
        Else
            userNozzle = NozzlePillar.readPropertiesNozzle(numberPillar, 1, "", "", DGV_Nozzle)
        End If
        'единый подферменник
        userAxisPillar.SingleSubFarmer = True
        '=======================================================================================================
        'читаем характеристики подферменника
        Erase userSubFermenter
        userSubFermenter = SubFermenters.readPropertiesSubFermenter(numberPillar, 1, DGV_SubFermenter, True)
        '=======================================================================================================
        'читаем характеристики откосного крыла
        userLeftHand = New HandPillar
        userRightHand = New HandPillar
        userLeftHand = HandPillar.readPropertiesHand(numberPillar, 1, DGV_LeftHand, True)
        userRightHand = HandPillar.readPropertiesHand(numberPillar, 1, DGV_RightHand, False)
        '======================================================================================================
        'читаем характеристики обратного открылка
        userLeftPostcard = New PostcardPillar
        userRightPostcard = New PostcardPillar
        userLeftPostcard = PostcardPillar.readPropertiesPostcard(numberPillar, 1, DGV_LeftPostcard, True)
        userRightPostcard = PostcardPillar.readPropertiesPostcard(numberPillar, 1, DGV_RightPostcard, False)
        If ChB_LeftPostcsrdFixedLenght.Checked = True Then
            userLeftPostcard.FixedLenght = True
            userRightPostcard.FixedLenght = True
        End If
        '======================================================================================================
        'читаем характеристики шкафной стенки
        userCabinetWall = New CabinetWallPillar
        userCabinetWall = CabinetWallPillar.readPropertiesCabinetWall(numberPillar, 1, DGV_CabinetWall)
        If ChB_FixedHeightCabinetWall.Checked = True Then
            userCabinetWall.FixedHeight = True
            userAxisPillar.FixedHeightCabinetWall = True
        End If
        '=====================================================================================================
        'читаеи характеристики ростверка
        userGrillage = New GrillagePillar
        userGrillage = GrillagePillar.readPropertiesGrillage(numberPillar, 1, collDiagram, rowDiagram, DGV_Grillage)
        'наличие ростверка
        If ChB_CreateGrillage.Checked = True Then
            userAxisPillar.PresencGrillage = False
        Else
            userAxisPillar.PresencGrillage = True
        End If
        '========================================================================================================
        'читаем характеристики подготовки
        userPreparation = New PreparationPillar
        userPreparation = PreparationPillar.readPropertiesPreparation(numberPillar, 1, DGV_Preparation)
        'наличие подготовки
        If ChB_CreatePreparation.Checked = True Then
            userAxisPillar.PresencPreparation = False
        Else
            userAxisPillar.PresencPreparation = True
        End If
        '=========================================================================================================
        'читаем характеристики стоек
        Erase arrayRack
        Dim userTypeRack As RackPillar.TypeRack = RackPillar.TypeRack.None
        If CB_RackTLC.Text Like "Circle_rack.tlc" Then
            userTypeRack = RackPillar.TypeRack.Circle
        ElseIf CB_RackTLC.Text Like "Oktagon_rack.tlc" Then
            userTypeRack = RackPillar.TypeRack.Octagonal
        ElseIf CB_RackTLC.Text Like "Trapezoidal_rack.tlc" Then
            userTypeRack = RackPillar.TypeRack.Trapezoidal
        End If
        Dim fixedHeight As Boolean = ChB_fixedHeightRack.Checked
        Dim egeParalel As Boolean = ChB_EgeParallel.Checked
        arrayRack = RackPillar.readPropertiesRack(numberPillar, 1, DGV_Rack, fixedHeight, egeParalel, userTypeRack)
        'вставить свю в стойку
        If ChB_InsertPileInRack.Checked = True Then
            userAxisPillar.PileInRack = True
            userAxisPillar.PresencGrillage = False
            userAxisPillar.PresencPreparation = False
        Else
            userAxisPillar.PileInRack = False
        End If
        '=========================================================================================================
        'читаем характеристики сваи
        Erase arrayPile
        Dim userTypePile As PilePillar.TypePile = PilePillar.TypePile.None
        If CB_PileTLC.Text Like "Circle_pile.tlc" Then
            userTypePile = PilePillar.TypePile.Drilling
        ElseIf CB_PileTLC.Text Like "Prismatic_pile.tlc" Then
            userTypePile = PilePillar.TypePile.Prismatic
        End If
        arrayPile = PilePillar.readPropertiesPile(numberPillar, 1, DGV_Piles, ChB_InsertPileInRack.Checked, ChB_PileExpand.Checked, userTypePile)
        'наличие стоек
        If ChB_CreateRack.Checked = True Then
            userAxisPillar.PresenceRacks = False
        Else
            userAxisPillar.PresenceRacks = True
        End If
        '===========================================================================================================================
        Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar)
        dataPillar.KeyParameter = keyParam
        bridgeProject.PlacementLastPillar(dataPillar, userNozzle, userSubFermenter, userCabinetWall, userLeftHand, userRightHand, userLeftPostcard, userRightPostcard, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, projectAlignment)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'находим максимальную отметку земли по оси шкафной стенки
        If IsNothing(projectSurface) = False Then
            Dim listPoly As List(Of Vector2D) = New List(Of Vector2D)
            listPoly.Add(userCabinetWall._elementBridgePoint.StartAxisPoint.Pos)
            listPoly.Add(userCabinetWall._elementBridgePoint.EndAxisPoint.Pos)
            Dim sectionSurface As Sfc.Sections.Section = projectSurface.CreateSection(listPoly, Sfc.Sections.SectionFlags.FilterRibs)
            If sectionSurface.Count > 0 Then
                For i As Integer = 0 To sectionSurface.Count - 1
                    Dim section As Sfc.Sections.SectionNode = sectionSurface.ElementAt(i)
                    If section.Vertex.Y > maxElevation Then
                        maxElevation = section.Vertex.Y
                    End If
                Next
            End If
        End If
        'вычисляем параметры трансформации относительно оси насадки
        If IsNothing(userNozzle) = False Then
            axisLineNozzle.StartPoint = userNozzle._elementBridgePoint.StartAxisPoint
            axisLineNozzle.EndPoint = userNozzle._elementBridgePoint.EndAxisPoint
            If axisLineNozzle.Length <= 0 Then
                Exit Sub
            End If
        End If
        userMatrixTransform.CalculateTransformation(axisLineNozzle.StartPoint, New Vector3D(0, 0, 0), axisLineNozzle.EndPoint, New Vector3D(axisLineNozzle.Length, 0, 0), maxElevation)
        '=========================================================================================================================
        'Получаем точки Насадки с учетом их трансформирования
        If IsNothing(userNozzle) = False Then
            If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
                Dim boolSetProp As Boolean = userNozzle.writeProjectData(DGV_Nozzle)
                listFullPointNozzle = userNozzle.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointNozzle = Pillar.removePointInList(listFullPointNozzle)
            End If
        End If
        '=========================================================================================================================
        'работаем с подферменниками (получаем все точки подферменников
        If IsArray(userSubFermenter) = True Then
            If userSubFermenter.Length > 0 Then
                For i As Integer = 0 To userSubFermenter.Length - 1
                    Dim tempUserSubFerm As SubFermenters = userSubFermenter(i)
                    Dim boolProp As Boolean = tempUserSubFerm.writeProjectData(DGV_SubFermenter)
                Next i
                Dim boolListointSubFerm As Boolean = SubFermenters.getProjectionPoint(userMatrixTransform, numberPillar, userSubFermenter, dictProjectPointPrevSubFermenters, dictProjectPointNextSubFermenters)
            End If
        End If
        '=========================================================================================================================
        'работаем со Шкафной стенкой
        If IsNothing(userCabinetWall) = False Then
            If userCabinetWall._elementBridgePoint.ListPointModel.Count > 0 Then
                Dim boolWriteProjectData As Boolean = userCabinetWall.writeProjectData(DGV_CabinetWall)
                listFullPointCabinetWall = userCabinetWall.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointCabinetWall = Pillar.removePointInList(listFullPointCabinetWall)
                'зуб упора
                listPointPlateCabinetWall = userCabinetWall.getProjectionPlatePoint(userMatrixTransform)
            End If
        End If
        '=========================================================================================================================
        'обратный открылок
        '=========================================================================================================================
        'Получаем точки Обратного левого открылка с учетом их трансформирования
        If IsNothing(userLeftHand) = False Then
            If userLeftHand._elementBridgePoint.ListPointModel.Count > 3 Then
                Dim boolWriteData As Boolean = userLeftHand.writeProjectData(DGV_LeftHand)
                listFullPointLeftHand = userLeftHand.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointLeftHand = Pillar.removePointInList(listFullPointLeftHand)
                'получаем точки карниза левого
                listPointLeftHandCornice = userLeftHand.getProjectionPointCornice(userMatrixTransform)
            End If
        End If
        If IsNothing(userRightHand) = False Then
            If userRightHand._elementBridgePoint.ListPointModel.Count > 3 Then
                Dim boolWriteData As Boolean = userRightHand.writeProjectData(DGV_RightHand)
                'Получаем точки Обратного правого открылка с учетом их трансформирования
                listFullPointRightHand = userRightHand.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointRightHand = Pillar.removePointInList(listFullPointRightHand)
                'получаем точки карниза правого
                listPointRightHandCornice = userRightHand.getProjectionPointCornice(userMatrixTransform)
            End If
        End If
        '==========================================================================================================================
        'откосное крыло
        '==========================================================================================================================
        'Получаем точки Откосного левого крыла с учетом их трансформирования
        If IsNothing(userLeftPostcard) = False Then
            If userLeftPostcard._elementBridgePoint.ListPointModel.Count > 3 Then
                Dim boolWriteData As Boolean = userLeftPostcard.writeProjectData(DGV_LeftPostcard)
                listFullPointLeftPostcard = userLeftPostcard.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointLeftPostcard = Pillar.removePointInList(listFullPointLeftPostcard)
            End If
        End If
        If IsNothing(userRightPostcard) = False Then
            If userRightPostcard._elementBridgePoint.ListPointModel.Count > 3 Then
                Dim boolWriteData As Boolean = userLeftPostcard.writeProjectData(DGV_RightPostcard)
                'Получаем точки Откосного правого крыла с учетом их трансформирования
                listFullPointRightPostcard = userRightPostcard.getProjectionPoint(userMatrixTransform)
                'удаляем точки лежащие на одной прямой
                listPointRightPostcard = Pillar.removePointInList(listFullPointRightPostcard)
            End If
        End If
        '==========================================================================================================================
        'стойки
        '==========================================================================================================================
        If IsArray(arrayRack) = True Then
            If arrayRack.Length > 0 Then
                For i As Integer = 0 To arrayRack.Length - 1
                    Dim tempUserRack As RackPillar = arrayRack(i)
                    Dim doorwritedata As Boolean = tempUserRack.writeProjectData(DGV_Rack)
                Next
                listPointRack = New Dictionary(Of Integer, List(Of Dictionary(Of String, ProjectionPoint)))
                If userAxisPillar.PresenceRacks = True Then
                    Dim boolFindRack As Boolean = RackPillar.getProjectionPoint(userMatrixTransform, numberPillar, arrayRack, listPointRack)
                End If
            End If
        End If
        '==========================================================================================================================
        'ростверк
        '==========================================================================================================================
        'Получаем точки Ростверка с учетом их трансформирования
        listFullPointGrillage = New List(Of Dictionary(Of String, ProjectionPoint))
        If userAxisPillar.PresencGrillage = True Then
            Dim doorwritedata As Boolean = userGrillage.writeProjectData(DGV_Grillage)
            listFullPointGrillage = userGrillage.getProjectionPoint(userMatrixTransform)
            'удаляем точки лежащие на одной прямой
            listPointGrillage = Pillar.removePointInList(listFullPointGrillage)
        End If
        '==========================================================================================================================
        'Подготовка
        '==========================================================================================================================
        'Получаем точки Подготовки с учетом их трансформирования
        listFullPointPreparation = New List(Of Dictionary(Of String, ProjectionPoint))
        If userAxisPillar.PresencPreparation = True Then
            Dim doorwritedata As Boolean = userPreparation.writeProjectData(DGV_Preparation)
            listFullPointPreparation = userPreparation.getProjectionPoint(userMatrixTransform)
            'удаляем точки лежащие на одной прямой
            listPointPreparation = Pillar.removePointInList(listFullPointPreparation)
        End If
        '==========================================================================================================================
        'Свая
        '==========================================================================================================================
        If IsArray(arrayPile) = True Then
            If arrayPile.Length > 0 Then
                For i As Integer = 0 To arrayPile.Length - 1
                    Dim tempUserPile As PilePillar = arrayPile(i)
                    Dim doorwritedata As Boolean = tempUserPile.writeProjectData(DGV_Piles)
                Next
                Dim boolFindPile As Boolean = PilePillar.getProjectionPoint(userMatrixTransform, arrayPile, dictProjectionPointPile)
            End If
        End If
        'отрисовываем виды
        PaintViewFrontPillar()
        PaintViewLeftPillar()
        PaintViewRightPillar()
        'PaintViewTopPillar()
        Button8.Enabled = True
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка OK ПОСТРОИТЬ СООРУЖЕНИЕ
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        bridgeProject = New ProjectBridge
        bridgeProject.BridgeModel = arrProject
        bridgeProject.EgSurface = egSurface
        bridgeProject.ProjectSurface = projectSurface
        bridgeProject.ProjectAlignment = projectAlignment
        bridgeProject.ActivDocument = activProjectDocument
        bridgeProject.DrawingLastPillar(dataPillar, userNozzle, userSubFermenter, userCabinetWall, userLeftHand, userRightHand, userLeftPostcard, userRightPostcard, arrayRack, userGrillage, userPreparation, arrayPile, dictionaryBridgeElements, projectSurface, egSurface, docPileTLC, templateXML)
        Try
            ApplicationHost.Current.Plugins.Execute("redrawall")
        Catch ex As System.Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование опоры спереди
    Private Sub PaintViewFrontPillar()
        'фронтальный вид PictureBox1
        Dim heightPB As Integer = PictureBox1.Height 'высота окна
        Dim widthPB As Integer = PictureBox1.Width 'ширина окна
        'определяем ось относительно которой будем вести построения (это ось насадки)
        Dim startPointAxis As Vector3D = userNozzle._elementBridgePoint.StartAxisPoint
        Dim endPointAxis As Vector3D = userNozzle._elementBridgePoint.EndAxisPoint
        Dim deltaL As Double = 5 'удлинняем ось чтобы рисование элемента не шло в притык к краям pictureBox
        Dim projectLenghtNozzle As Integer = axisLineNozzle.Length + deltaL
        If projectLenghtNozzle = 0 Then Exit Sub
        'вычисляем коэффициент по горизонтали по оси Х
        Dim k As Double = MathFunction.FuncTrimDigitFloo(widthPB / projectLenghtNozzle, 3)
        'определяем масштаб по вертикали (по оси Y)'двойка - это запас)
        '1. определяем максимальную отметку земли
        If maxElevation = -9999 Then
            If IsNothing(userCabinetWall) = False Then
                maxElevation = startPointAxis.Z + userCabinetWall.CenterHeight + userCabinetWall.ElevationOffsetProjectSurface
            End If
        End If
        If maxElevation = -9999 Then
            maxElevation = startPointAxis.Z
        End If
        'прикидываем среднюю высоту опоры
        Dim heightPillar As Integer = 4 'дополнительный запас
        If IsNothing(userCabinetWall) = False Then
            heightPillar += userCabinetWall.CenterHeight + userCabinetWall.ElevationOffsetProjectSurface
        End If
        If IsNothing(userNozzle) = False Then
            heightPillar += userNozzle.SecondHeight
        End If
        If IsArray(arrayRack) = True Then
            Dim userRack As RackPillar = arrayRack(0)
            heightPillar += userRack.Height
        End If
        If IsNothing(userGrillage) = False Then
            heightPillar += userGrillage.Height
        End If
        If IsNothing(userPreparation) = False Then
            heightPillar += userPreparation.Height
        End If
        If IsArray(arrayUserPile) = True Then
            Dim userPile As PilePillar = arrayPile(0)
            heightPillar += userPile.Height
        End If
        'вычисляем вертикальный коэффициент
        Dim k2 As Double = MathFunction.FuncTrimDigitFloo(heightPB / heightPillar, 3)
        If k2 < k Then k = k2
        'умножаем полученный результат на пользовательский масштаб
        k = k * NUpD_ScaleFront.Value
        'начальные координаты для рисования
        Dim lenghtAxisPillar As Integer = axisLineNozzle.Length * k
        'крайняя левая точка
        Dim x0 As Integer = (widthPB - lenghtAxisPillar) / 2 + NUpD_dxFront.Value * k
        'рассчитываем вертикаль
        heightPillar *= k
        Dim y0 As Integer = (heightPB - heightPillar) / 2 + NUpD_dyFront.Value * k
        'настройки рисования
        Dim borderRigel As Integer = Integer.Parse(1)
        Dim bmp As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        Dim colorPen As Color = Color.Black
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(Color.White)
            '============================================================================================================================================
            'рисуем насадку
            'забираем точки для проекции front
            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Nothing
            Dim positionPointNozzleX As Double = 0
            If listPointNozzle.Count > 0 Then
                dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Front)
                If dictProjectionPoint.ContainsKey("leftPt1-top") = True Then
                    positionPointNozzleX = dictProjectionPoint.Item("leftPt1-top").projectPoint.X
                End If
                Dim offsetNozzle As Double = -1 * k * positionPointNozzleX
                'рисуем элемент
                Dim drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetNozzle, y0, ProjectionPoint.projectView.Front)
            End If
            '============================================================================================================================================
            'рисуем подферменник
            Dim pointLeft1CabinetWall As Vector3D = userCabinetWall.getPointByCode("leftPt1")
            Dim pointRight1CabinetWall As Vector3D = userCabinetWall.getPointByCode("rightPt1")
            Dim projectLeft1CabinetWall = userMatrixTransform.transformPoint(pointLeft1CabinetWall.Pos)
            Dim projectright1CabinetWall = userMatrixTransform.transformPoint(pointRight1CabinetWall.Pos)
            Dim d As Double = (projectright1CabinetWall.X - projectLeft1CabinetWall.X)
            Dim offsetX As Double = 0
            Dim DeltaX As Double = 0
            If dictProjectPointPrevSubFermenters.Count > 0 Then
                dictProjectionPoint = dictProjectPointPrevSubFermenters.First.Value.Item(0)
                If dictProjectionPoint.ContainsKey("leftPt1-top") = True Then
                    DeltaX = dictProjectionPoint.Item("leftPt1-top").projectPoint.X
                End If
                offsetX = -1 * k * positionPointNozzleX - d * k
                For i As Integer = 0 To dictProjectPointPrevSubFermenters.Count - 1
                    'забираем точки для проекции back
                    Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointPrevSubFermenters.ElementAt(i).Value, ProjectionPoint.projectView.Front)
                    'рисуем элемент
                    Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
                Next i
            ElseIf dictProjectPointNextSubFermenters.Count > 0 Then
                dictProjectionPoint = dictProjectPointNextSubFermenters.First.Value.Item(0)
                If dictProjectionPoint.ContainsKey("leftPt1-top") = True Then
                    DeltaX = dictProjectionPoint.Item("leftPt1-top").projectPoint.X
                End If
                offsetX = -1 * k * positionPointNozzleX - d * k
                For i As Integer = 0 To dictProjectPointNextSubFermenters.Count - 1
                    'забираем точки для проекции back
                    Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointNextSubFermenters.ElementAt(i).Value, ProjectionPoint.projectView.Front)
                    'рисуем элемент
                    Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
                Next i
            End If
            '============================================================================================================================================
            'рисуем шкафную стенку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointCabinetWall, ProjectionPoint.projectView.Front)
            'рисуем элемент
            Dim deltaPosNozzle As Double = positionPointNozzleX
            offsetX = -1 * k * deltaPosNozzle
            Dim drawCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            'рисуем проектную поверхность
            Dim axisLineCabinetWall As DwgLine = New DwgLine
            axisLineCabinetWall.StartPoint = userCabinetWall._elementBridgePoint.StartAxisPoint
            axisLineCabinetWall.EndPoint = userCabinetWall._elementBridgePoint.EndAxisPoint
            d = (projectright1CabinetWall.X - projectLeft1CabinetWall.X)
            offsetX = -1 * k * positionPointNozzleX - d * k
            Dim boolDrawSurface = Pillar.drawSurfaceLine(bmp, projectSurface, axisLineCabinetWall, userMatrixTransform.DeltaZ, k, x0 + offsetX, y0, False, False)
            'рисуем поверхность земли
            Dim boolDrawEgSurface = Pillar.drawSurfaceLine(bmp, egSurface, axisLineCabinetWall, userMatrixTransform.DeltaZ, k, x0 + offsetX, y0, False, True)
            '============================================================================================================================================
            'рисуем обратные открылки
            'забираем точки для проекции back
            Dim ptNozzleMiddle1 As Vector3D = userNozzle.getPointByCode("middlePt1")
            Dim st As Vector2D = userMatrixTransform.transformPoint(ptNozzleMiddle1)
            offsetX = -1 * k * st.X
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftHand, ProjectionPoint.projectView.Left)
            Dim drawLeftHand = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightHand, ProjectionPoint.projectView.Left)
            Dim drawRightHand = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftHandCornice, ProjectionPoint.projectView.Left)
            Dim drawLeftHandCornice = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightHandCornice, ProjectionPoint.projectView.Left)
            Dim drawRightHandCornice = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            '============================================================================================================================================
            'рисуем откосные крылья
            'забираем точки для проекции back
            Dim ptNozzleRight1 As Vector3D = userNozzle.getPointByCode("rightPt1")
            st = userMatrixTransform.transformPoint(ptNozzleMiddle1)
            offsetX = -1 * st.X * k
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftPostcard, ProjectionPoint.projectView.Left)
            Dim drawLeftPostcard = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightPostcard, ProjectionPoint.projectView.Left)
            Dim drawRightPostcard = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            '============================================================================================================================================
            'рисуем стойки
            'забираем точки для проекции back
            If listPointRack.Count > 0 Then
                Dim ptRackRight1 As Vector3D = New Vector3D
                Dim boolFindPoint As Boolean = arrayRack(0).getPointByCode("rightPt1", ptRackRight1)
                If boolFindPoint = True Then
                    If arrayRack(0).RackType = RackPillar.TypeRack.Trapezoidal Then
                        st = userMatrixTransform.transformPoint(ptNozzleMiddle1)
                        offsetX = -1 * k * st.X
                    Else
                        offsetX = 0
                    End If
                    For i As Integer = 0 To listPointRack.Count - 1
                        'забираем точки для проекции back
                        Dim dictProjectionPointRack As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointRack.ElementAt(i).Value, ProjectionPoint.projectView.Front)
                        'рисуем элемент
                        Dim drawRack = Pillar.drawCounterPictureBox(bmp, dictProjectionPointRack, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
                    Next i
                Else
                    For i As Integer = 0 To listPointRack.Count - 1
                        'забираем точки для проекции back
                        Dim dictProjectionPointRack As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointCircleRackForDraw(listPointRack.ElementAt(i).Value, ProjectionPoint.projectView.Front)
                        'рисуем элемент
                        Dim drawRack = Pillar.drawCounterPictureBox(bmp, dictProjectionPointRack, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
                    Next i
                End If

            End If
            '============================================================================================================================================
            'рисуем ростверк
            Dim ptGrillageRight1 As Vector3D = userGrillage.getPointByCode("rightPt1")
            st = userMatrixTransform.transformPoint(ptGrillageRight1)
            offsetX = positionPointNozzleX
            dictProjectionPoint = Pillar.selectPointForDraw(listPointGrillage, ProjectionPoint.projectView.Front)
            Dim drawGrillage = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            '============================================================================================================================================
            'рисуем подготовку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointPreparation, ProjectionPoint.projectView.Front)
            Dim drawPreparation = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0 + offsetX, y0, ProjectionPoint.projectView.Front)
            '============================================================================================================================================
            'рисуем сваи
            Dim numberRowPile As Integer = NUpD_CountRowsPile.Value
            Dim posBottomElement As Integer = y0
            If dictProjectionPointPile.ContainsKey(numberRowPile) = True Then
                Dim dictRowPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectionPointPile.Item(numberRowPile)
                If dictRowPile.Count > 0 Then
                    For i As Integer = 0 To dictRowPile.Count - 1
                        Dim arrayProjectPoint As ProjectionPoint() = dictRowPile.ElementAt(i).Value
                        Dim offsetPile As Double = x0 + offsetX
                        Dim userFindPile As PilePillar = PilePillar.findPileToArray(numberRowPile, i + 1, arrayPile)
                        If IsNothing(userFindPile) = False Then
                            Dim offsetDist As Double = 0
                            If i = 0 Then
                                offsetDist = userFindPile.OffsetX
                            ElseIf i = dictRowPile.Count - 1 Then
                                If (userGrillage.Lenght - userFindPile.OffsetX) > 0 Then
                                    offsetDist = userGrillage.Lenght - userFindPile.OffsetX
                                Else
                                    offsetDist = userNozzle.Lenght - userFindPile.OffsetX
                                End If
                            End If
                            Dim boolDrawPile As Boolean = PilePillar.drawPilePictureBox(bmp, arrayProjectPoint, userFindPile, k, offsetPile, y0, ProjectionPoint.projectView.Front, i + 1, offsetDist)
                        End If
                    Next i
                End If
            End If
            '==========================================================================================================================================
            'рисуем ростверк или насадку сверху
            'забираем точки для проекции Bottom
            Dim widthBox As Integer = 0
            If userGrillage._elementBridgePoint.ListPointModel.Count > 3 Then
                widthBox = userGrillage.Width * k
                y0 = widthPB - 10
                dictProjectionPoint = Pillar.selectPointForDraw(listPointGrillage, ProjectionPoint.projectView.Bottom)
            Else
                widthBox = userNozzle.Width * k
                y0 = widthPB - 10
                dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Bottom)
            End If
            'рисуем насадку
            Dim boolDrawTopElement = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Bottom)
            If dictProjectionPointPile.Count > 0 Then
                For i As Integer = 0 To dictProjectionPointPile.Count - 1
                    Dim dictRowPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectionPointPile.ElementAt(i).Value
                    If dictRowPile.Count > 0 Then
                        For j As Integer = 0 To dictRowPile.Count - 1
                            Dim userFindPile As PilePillar = PilePillar.findPileToArray(numberRowPile, i + 1, arrayPile)
                            Dim arrayProjectPoint As ProjectionPoint() = dictRowPile.ElementAt(j).Value
                            Dim pjectPoint As Vector3D = arrayProjectPoint(0).projectPoint
                            Dim size As Single = userFindPile.Width * k
                            If size = 0 Then
                                size = userFindPile.Diameter * k
                            End If
                            Dim offsetXPile As Integer = x0 + pjectPoint.X * k - size / 2
                            Dim offsetYPile As Integer = y0 + pjectPoint.Y * k - size / 2
                            Using brush As New SolidBrush(colorPen)
                                g.FillEllipse(brush, offsetXPile, offsetYPile, size, size)
                            End Using
                        Next j
                    End If
                Next i
            End If
        End Using
        PictureBox1.Image = bmp
    End Sub
    'вид сбоку слева
    Private Sub PaintViewLeftPillar()
        If IsNothing(userNozzle) = True Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вид сбоку
        Dim heightPB As Integer = PictureBox2.Height
        Dim widthPB As Integer = PictureBox2.Width
        Dim heightPillar As Double = 0
        Dim widthPillar As Double = 0
        If IsNothing(userNozzle) = False Then
            heightPillar += userNozzle.SecondHeight
            widthPillar += userNozzle.Width
        End If
        If IsNothing(userLeftHand) = False Then
            widthPillar += userLeftHand.LengthTop
        End If
        'определяем масштаб по Y
        If IsNothing(userCabinetWall) = False Then
            heightPillar += userCabinetWall.CenterHeight
        End If
        If IsArray(arrayRack) = True Then
            Dim userRack As RackPillar = arrayRack(0)
            heightPillar += userRack.Height
        End If
        If IsNothing(userGrillage) = False Then
            heightPillar += userGrillage.Height
        End If
        If IsNothing(userPreparation) = False Then
            heightPillar += userPreparation.Height
        End If
        If IsArray(arrayUserPile) = True Then
            Dim userPile As PilePillar = arrayPile(0)
            heightPillar += userPile.Height
        End If
        Dim k As Double = MathFunction.FuncTrimDigitFloo(heightPB / (heightPillar), 3)
        k = k * NUpD_ScaleLeft.Value
        'начальная точка рисования
        heightPillar = (heightPillar) * k
        Dim x0 As Integer = (widthPB) / 2 + NUpD_dxLeft.Value * k
        Dim y0 As Integer = (heightPB - heightPillar) / 2 + NUpD_dyLeft.Value * k
        Dim bmp As New Bitmap(PictureBox2.Width, PictureBox2.Height)
        Dim borderRigel As Integer = Integer.Parse(1)
        Dim colorPen As Color = Color.Black
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(Color.White)
            'забираем точки для проекции left
            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Left)
            'рисуем насадку
            Dim drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем подферменник
            If dictProjectPointPrevSubFermenters.Count > 0 Then
                'забираем точки для проекции back
                Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointPrevSubFermenters.First.Value, ProjectionPoint.projectView.Left)
                'рисуем элемент
                Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Left)
            ElseIf dictProjectPointNextSubFermenters.Count > 0 Then
                'забираем точки для проекции back
                Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointNextSubFermenters.First.Value, ProjectionPoint.projectView.Left)
                'рисуем элемент
                Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Left)
            End If
            '============================================================================================================================================
            'рисуем шкафную стенку (основной элемент)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointCabinetWall, ProjectionPoint.projectView.Left)
            'рисуем элемент
            Dim drawCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            'рисуем шкафную стенку (зуб упора)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointPlateCabinetWall, ProjectionPoint.projectView.Left)
            'рисуем элемент
            Dim drawPlateCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем откосное крыло (основной элемент)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftHand, ProjectionPoint.projectView.Back)
            'рисуем элемент
            Dim drawLeftHand = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            'рисуем карниз
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftHandCornice, ProjectionPoint.projectView.Back)
            'рисуем элемент
            Dim drawLeftHandCornice = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            '============================================================================================================================================
            'рисуем обратный открылок (основной элемент)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointLeftPostcard, ProjectionPoint.projectView.Front)
            'рисуем элемент
            Dim drawLeftPostcard = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем стойку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRack.Item(1), ProjectionPoint.projectView.Left)
            'рисуем элемент
            Dim drawLeftRask = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем Ростверк
            dictProjectionPoint = Pillar.selectPointForDraw(listPointGrillage, ProjectionPoint.projectView.Left)
            'рисуем элемент
            Dim drawGrillage = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем Подготовку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointPreparation, ProjectionPoint.projectView.Left)
            'рисуем элемент
            Dim drawPreparation = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем Сваи
            Dim numberColumnPile As Integer = NUpD_CountColumnsPile.Value
            For i As Integer = 0 To dictProjectionPointPile.Count - 1
                Dim dictRowPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectionPointPile.ElementAt(i).Value
                If dictRowPile.Count > 0 Then
                    If dictRowPile.ContainsKey(numberColumnPile) = True Then
                        Dim arrayProjectPoint As ProjectionPoint() = dictRowPile.ElementAt(i).Value
                        Dim userFindPile As PilePillar = PilePillar.findPileToArray(i + 1, numberColumnPile, arrayPile)
                        If IsNothing(userFindPile) = False Then
                            Dim offsetDist As Double = 0
                            If i = 0 Then
                                offsetDist = userFindPile.OffsetY
                            ElseIf i = dictProjectionPointPile.Count - 1 Then
                                If (userGrillage.Width - userFindPile.OffsetY) > 0 Then
                                    offsetDist = userGrillage.Width - userFindPile.OffsetY
                                Else
                                    offsetDist = userGrillage.Width - userFindPile.OffsetY
                                End If
                            End If
                            Dim boolDrawPile As Boolean = PilePillar.drawPilePictureBox(bmp, arrayProjectPoint, userFindPile, k, x0, y0, ProjectionPoint.projectView.Left, i + 1, offsetDist)
                        End If
                    End If
                End If
            Next i
            'рисуем проектную поверхность
            Dim axisLineSurface As DwgLine = New DwgLine
            Dim pointRightNozzle As Vector3D = userNozzle.getPointByCode("rightPt1")
            Dim pointRightHand As Vector3D = userLeftHand.getPointByCode("leftPt2")
            axisLineSurface.StartPoint = pointRightNozzle.Pos
            axisLineSurface.EndPoint = pointRightHand.Pos
            Dim delta As Double = (axisLineSurface.Length - userNozzle.Width / 2) * k
            Dim boolDrawSurface = Pillar.drawSurfaceLine(bmp, projectSurface, axisLineSurface, userMatrixTransform.DeltaZ, k, x0 - delta, y0, False, False)
            'рисуем поверхность земли
            Dim boolDrawEgSurface = Pillar.drawSurfaceLine(bmp, egSurface, axisLineSurface, userMatrixTransform.DeltaZ, k, x0 - delta, y0, False, True)
        End Using
        PictureBox2.Image = bmp
    End Sub
    'вид сбоку справа
    Private Sub PaintViewRightPillar()
        If IsNothing(userNozzle) = True Then
            Exit Sub
        End If
        Dim heightPB As Integer = PictureBox5.Height
        Dim widthPB As Integer = PictureBox5.Width
        Dim heightPillar As Double = 0
        Dim widthPillar As Double = 0
        If IsNothing(userNozzle) = False Then
            heightPillar += userNozzle.SecondHeight
            widthPillar += userNozzle.Width
        End If
        If IsNothing(userLeftHand) = False Then
            widthPillar += userLeftHand.LengthTop
        End If
        'определяем масштаб по Y
        If IsNothing(userCabinetWall) = False Then
            heightPillar += userCabinetWall.CenterHeight
        End If
        If IsArray(arrayRack) = True Then
            Dim userRack As RackPillar = arrayRack(0)
            heightPillar += userRack.Height
        End If
        If IsNothing(userGrillage) = False Then
            heightPillar += userGrillage.Height
        End If
        If IsNothing(userPreparation) = False Then
            heightPillar += userPreparation.Height
        End If
        If IsArray(arrayUserPile) = True Then
            Dim userPile As PilePillar = arrayPile(0)
            heightPillar += userPile.Height
        End If
        Dim k As Double = MathFunction.FuncTrimDigitFloo(heightPB / (heightPillar), 3)
        k = k * +NUpD_ScaleRight.Value
        heightPillar = (heightPillar) * k
        Dim x0 As Integer = (widthPB) / 2 + +NUpD_dxRight.Value
        Dim y0 As Integer = (heightPB - heightPillar) / 2 + NUpD_dyRight.Value
        Dim bmp As New Bitmap(PictureBox5.Width, PictureBox5.Height)
        Dim borderRigel As Integer = Integer.Parse(1)
        Dim colorPen As Color = Color.Black
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.Clear(Color.White)
            'забираем точки для проекции left
            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Right)
            'рисуем насадку
            Dim drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            '============================================================================================================================================
            'рисуем подферменник
            If dictProjectPointPrevSubFermenters.Count > 0 Then
                'забираем точки для проекции back
                Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointPrevSubFermenters.Last.Value, ProjectionPoint.projectView.Right)
                'рисуем элемент
                Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Right)
            ElseIf dictProjectPointNextSubFermenters.Count > 0 Then
                'забираем точки для проекции back
                Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointNextSubFermenters.Last.Value, ProjectionPoint.projectView.Right)
                'рисуем элемент
                Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Right)
            End If
            '============================================================================================================================================
            'рисуем шкафную стенку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointCabinetWall, ProjectionPoint.projectView.Right)
            'рисуем элемент
            Dim drawCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            'рисуем шкафную стенку (зуб упора)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointPlateCabinetWall, ProjectionPoint.projectView.Right)
            'рисуем элемент
            Dim drawPlateCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            '============================================================================================================================================
            'рисуем откосное крыло (основной элемент)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightHand, ProjectionPoint.projectView.Back)
            'рисуем элемент
            Dim drawRightHand = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            'рисуем карниз
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightHandCornice, ProjectionPoint.projectView.Back)
            'рисуем элемент
            Dim drawRightHandCornice = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем обратный открылок (основной элемент)
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRightPostcard, ProjectionPoint.projectView.Front)
            'рисуем элемент
            Dim drawRightPostcard = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
            '============================================================================================================================================
            'рисуем стойку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointRack.Item(listPointRack.Count - 1), ProjectionPoint.projectView.Right)
            'рисуем элемент
            Dim drawLeftRask = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            '============================================================================================================================================
            'рисуем Ростверк
            dictProjectionPoint = Pillar.selectPointForDraw(listPointGrillage, ProjectionPoint.projectView.Right)
            'рисуем элемент
            Dim drawGrillage = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            '============================================================================================================================================
            'рисуем Подготовку
            dictProjectionPoint = Pillar.selectPointForDraw(listPointPreparation, ProjectionPoint.projectView.Right)
            'рисуем элемент
            Dim drawPreparation = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
            '============================================================================================================================================
            'рисуем Сваи
            Dim numberColumnPile As Integer = NUpD_CountColumnsPile.Value
            For i As Integer = 0 To dictProjectionPointPile.Count - 1
                Dim dictRowPile As Dictionary(Of Integer, ProjectionPoint()) = dictProjectionPointPile.ElementAt(i).Value
                If dictRowPile.Count > 0 Then
                    If dictRowPile.ContainsKey(numberColumnPile) = True Then
                        Dim arrayProjectPoint As ProjectionPoint() = dictRowPile.ElementAt(i).Value
                        Dim userFindPile As PilePillar = PilePillar.findPileToArray(i + 1, numberColumnPile, arrayPile)
                        If IsNothing(userFindPile) = False Then
                            Dim offsetDist As Double = 0
                            If i = 0 Then
                                offsetDist = userFindPile.OffsetY
                            ElseIf i = dictProjectionPointPile.Count - 1 Then
                                If (userGrillage.Width - userFindPile.OffsetY) > 0 Then
                                    offsetDist = userGrillage.Width - userFindPile.OffsetY
                                Else
                                    offsetDist = userGrillage.Width - userFindPile.OffsetY
                                End If
                            End If
                            Dim boolDrawPile As Boolean = PilePillar.drawPilePictureBox(bmp, arrayProjectPoint, userFindPile, k, x0, y0, ProjectionPoint.projectView.Left, i + 1, offsetDist)
                        End If
                    End If
                End If
            Next i
            'рисуем проектную поверхность
            Dim axisLineSurface As DwgLine = New DwgLine
            Dim pointRightNozzle As Vector3D = userNozzle.getPointByCode("rightPt2")
            Dim pointRightHand As Vector3D = userLeftHand.getPointByCode("leftPt2")
            axisLineSurface.StartPoint = pointRightNozzle.Pos
            axisLineSurface.EndPoint = pointRightHand.Pos
            Dim delta As Double = (axisLineSurface.Length - userNozzle.Width / 2) * k
            Dim boolDrawSurface = Pillar.drawSurfaceLine(bmp, projectSurface, axisLineSurface, userMatrixTransform.DeltaZ, k, x0 - delta, y0, False, False)
            'рисуем поверхность земли
            Dim boolDrawEgSurface = Pillar.drawSurfaceLine(bmp, egSurface, axisLineSurface, userMatrixTransform.DeltaZ, k, x0 - delta, y0, False, True)
        End Using
        PictureBox5.Image = bmp
    End Sub
    'вид по элементно
    'Private Sub PaintViewTopPillar()
    '    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '    'вид сбоку
    '    Dim heightPB As Integer = PictureBox3.Height
    '    Dim widthPB As Integer = PictureBox3.Width
    '    Dim activeTabText As String = TabControl1.SelectedTab.Text 'получаем активную вкладку
    '    'удлинняем ось опоры
    '    If axisLinePillar.Length = 0 Then
    '        Exit Sub
    '    End If
    '    Dim tempAxisLineNozzle As DwgLine = axisLinePillar.Clone
    '    Dim boolExtend As Boolean = BridgeGeometry.extendBeam(tempAxisLineNozzle, 2, 2)
    '    'переводим все в мм
    '    Dim orignLenghtNozzle As Integer = tempAxisLineNozzle.Length * 1000
    '    If orignLenghtNozzle = 0 Then Exit Sub
    '    'вычисляем коэффициент по горизонтали по оси Х
    '    Dim k As Double = MathFunction.FuncTrimDigitFloo(widthPB / orignLenghtNozzle, 3)
    '    k = k * 1000
    '    'определяем масштаб по вертикали (по оси Y)'двойка - это запас)
    '    Dim heightPillar As Integer = userNozzle.SecondHeight + userCabinetWall.Width + userPile.Height + 2
    '    If heightPillar = 0 Then Exit Sub
    '    If ChB_CreateRack.Checked = False Then
    '        heightPillar += userRack.Height
    '    End If
    '    If ChB_CreateGrillage.Checked = False Then
    '        heightPillar += userGrillage.Height
    '    End If
    '    If ChB_CreatePreparation.Checked = False Then
    '        heightPillar += userPreparation.Height
    '    End If
    '    Dim tempheightPillar = heightPillar * 1000
    '    Dim k2 As Double = MathFunction.FuncTrimDigitFloo(heightPB / tempheightPillar, 3)
    '    k2 = k2 * 1000
    '    If k2 < k Then k = k2
    '    k = k * NUpD_ScaleFront.Value
    '    'начальные координаты для рисования
    '    Dim lenghtAxisPillar As Integer = axisLinePillar.Length * k
    '    'крайняя левая точка
    '    Dim x0 As Integer = (widthPB - lenghtAxisPillar) / 2 + NUpD_dxFront.Value
    '    'рассчитываем вертикаль
    '    heightPillar *= k
    '    Dim y0 As Integer = 100 + NUpD_dyFront.Value
    '    'настройки рисования
    '    Dim borderRigel As Integer = Integer.Parse(1)
    '    Dim bmp As New Bitmap(PictureBox1.Width, PictureBox1.Height)
    '    Dim colorPen As Color = Color.Black
    '    Using g As Graphics = Graphics.FromImage(bmp)
    '        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
    '        g.Clear(Color.White)
    '        If activeTabText Like "Насадка" Then
    '            '==========================================================================================================================================
    '            'рисуем насадку
    '            'забираем точки для проекции Top
    '            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Top)
    '            'рисуем насадку
    '            Dim drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Top)
    '            y0 = y0 + userNozzle.SecondHeight * k + 100
    '            'забираем точки для проекции Bottom
    '            dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Bottom)
    '            'рисуем насадку
    '            drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Bottom)
    '            y0 = y0 + userNozzle.SecondHeight * k + 100
    '            'забираем точки для проекции Front
    '            dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Front)
    '            'рисуем насадку
    '            drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Front)
    '            y0 = y0 + userNozzle.Width * k + 100
    '            'забираем точки для проекции Back
    '            dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Back)
    '            'рисуем насадку
    '            drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Back)
    '            y0 = y0 + userNozzle.Width * k + 100
    '            'забираем точки для проекции Left
    '            dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Left)
    '            'рисуем насадку
    '            drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
    '            y0 = y0 + userNozzle.Width * k + 100
    '            'забираем точки для проекции Right
    '            dictProjectionPoint = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Right)
    '            'рисуем насадку
    '            drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Right)
    '        ElseIf activeTabText Like "Подферменник" Then
    '            '==========================================================================================================================================
    '            'рисуем насадку
    '            'забираем точки для проекции Top
    '            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointNozzle, ProjectionPoint.projectView.Top)
    '            'рисуем насадку
    '            Dim drawNozzle = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Top)
    '            '============================================================================================================================================
    '            'рисуем подферменник
    '            If dictProjectPointPrevSubFermenters.Count > 0 Then
    '                For i As Integer = 0 To dictProjectPointPrevSubFermenters.Count - 1
    '                    'забираем точки для проекции back
    '                    Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointPrevSubFermenters.ElementAt(i).Value, ProjectionPoint.projectView.Top)
    '                    'рисуем элемент
    '                    Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Top)
    '                Next i
    '            ElseIf dictProjectPointNextSubFermenters.Count > 0 Then
    '                For i As Integer = 0 To dictProjectPointNextSubFermenters.Count - 1
    '                    'забираем точки для проекции back
    '                    Dim dictProjectionPointSubFerm As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(dictProjectPointNextSubFermenters.ElementAt(i).Value, ProjectionPoint.projectView.Top)
    '                    'рисуем элемент
    '                    Dim drawSubFerm = Pillar.drawCounterPictureBox(bmp, dictProjectionPointSubFerm, k, x0, y0, ProjectionPoint.projectView.Top)
    '                Next i
    '            End If
    '        ElseIf activeTabText Like "Шкафная стенка" Then
    '            '============================================================================================================================================
    '            'рисуем землю

    '            'рисуем шкафную стенку
    '            Dim dictProjectionPoint As Dictionary(Of String, ProjectionPoint) = Pillar.selectPointForDraw(listPointCabinetWall, ProjectionPoint.projectView.Left)
    '            'рисуем элемент
    '            Dim drawCabinetWall = Pillar.drawCounterPictureBox(bmp, dictProjectionPoint, k, x0, y0, ProjectionPoint.projectView.Left)
    '        End If
    '    End Using
    '    PictureBox3.Image = bmp
    'End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменить количество стоек
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NUpD_CountRack.ValueChanged
        Dim countRack As Integer = NUpD_CountRack.Value
        Dim countColumnsRack As Integer = DGV_Rack.ColumnCount
        If countColumnsRack > 1 Then
            If countRack > countColumnsRack - 1 Then
                Dim dataColl As DataGridViewColumn = DGV_Rack.Columns.Item(1)
                Dim newDataColl As DataGridViewColumn = New DataGridViewColumn()
                newDataColl = dataColl.Clone()
                Dim posRow As Integer = DGV_Rack.Columns.Add(newDataColl)
                DGV_Rack.Columns.Item(countRack).HeaderText = "Стойка №" & countRack
                For i As Integer = 0 To DGV_Rack.RowCount - 1
                    Dim tagRow As String = DGV_Rack.Rows(i).Tag
                    If tagRow Like "number" Then
                        DGV_Rack.Rows(i).Cells(countRack).Value = countRack
                    Else
                        Dim copyData As String = DGV_Rack.Rows(i).Cells(1).Value
                        DGV_Rack.Rows(i).Cells(countRack).Value = copyData
                    End If
                Next i
            ElseIf countRack < countColumnsRack - 1 Then
                If countColumnsRack >= countRack + 1 Then
                    DGV_Rack.Columns.RemoveAt(countRack + 1)
                End If
            End If
        End If

    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'заполнить datagrid со сваями (Применить схему)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'делаем разбор строк
        Dim pileRowsFieldDiagram As String = TxtB_PileRowDiagram.Text
        Dim pileColumnFieldDiagram As String = TxtB_PileCollDiagram.Text
        If IsNothing(pileRowsFieldDiagram) = True Then
            MsgBox("Не удалось прочитать схему раскладки свай!!!")
            Exit Sub
        End If
        If IsNothing(pileColumnFieldDiagram) = True Then
            MsgBox("Не удалось прочитать схему раскладки свай!!!")
            Exit Sub
        End If
        'делаем разбор строки 1
        Dim countRowsPile As Integer = 0
        Dim countColumnsPile As Integer = 0
        If pileColumnFieldDiagram.Trim.Length > 0 And pileRowsFieldDiagram.Trim.Length Then
            Dim arrayStr As String() = pileRowsFieldDiagram.Split("+")
            If IsArray(arrayStr) = True Then
                For i As Integer = 0 To arrayStr.Length - 1
                    Dim pileDiagram As String = arrayStr(i)
                    Dim arrayD As String() = pileDiagram.Split("*")
                    If arrayD.Length = 2 Then
                        Dim countPileTemp As Integer = Val(arrayD(0))
                        countRowsPile += countPileTemp
                    End If
                Next i
            End If
            arrayStr = pileColumnFieldDiagram.Split("+")
            If IsArray(arrayStr) = True Then
                For i As Integer = 0 To arrayStr.Length - 1
                    Dim pileDiagram As String = arrayStr(i)
                    Dim arrayD As String() = pileDiagram.Split("*")
                    If arrayD.Length = 2 Then
                        Dim countPileTemp As Integer = Val(arrayD(0))
                        countColumnsPile += countPileTemp
                    End If
                Next i
            End If
        End If
        If countColumnsPile = 0 Then
            countColumnsPile = 1
        End If
        If countRowsPile = 0 Then
            countRowsPile = 1
        End If
        If userPile.Height = 0 Or userPile.Width = 0 Then
            Dim arrayPile As PilePillar() = PilePillar.readPropertiesPile(numberPillar, numberSubPillar, DGV_Piles)
        End If
        NUpD_CountRowsPile.Minimum = 1
        NUpD_CountRowsPile.Maximum = countColumnsPile
        NUpD_CountColumnsPile.Minimum = 1
        NUpD_CountColumnsPile.Maximum = countRowsPile
        Dim countPiles As Integer = 1
        If DGV_Piles.Columns.Count > 2 Then
            For k As Integer = DGV_Piles.ColumnCount - 1 To 2 Step -1
                DGV_Piles.Columns.RemoveAt(k)
            Next k
        End If
        If countRowsPile > 0 And countColumnsPile > 0 Then
            dictUserPile = New Dictionary(Of Integer, Dictionary(Of Integer, PilePillar))
        End If
        For i As Integer = 1 To countColumnsPile
            Dim dictPile As Dictionary(Of Integer, PilePillar) = New Dictionary(Of Integer, PilePillar)
            For j As Integer = 1 To countRowsPile
                Dim newUserPile As PilePillar = New PilePillar()
                newUserPile.NumberPillar = userPile.NumberPillar
                newUserPile.NumberSubPillars = userPile.NumberSubPillars
                newUserPile.NumberRow = i
                newUserPile.NumberColumn = j
                newUserPile.AngleX = 90
                newUserPile.AngleY = 90
                newUserPile.Height = userPile.Height
                newUserPile.Diameter = userPile.Diameter
                newUserPile.Width = userPile.Width
                newUserPile.TopSeal = userPile.TopSeal
                newUserPile.Expand = userPile.Expand
                newUserPile.HeightExpand = userPile.HeightExpand
                newUserPile.WidthExpand = userPile.WidthExpand
                newUserPile.HeightDownExpand = userPile.HeightDownExpand
                newUserPile.DegExpand = userPile.DegExpand
                newUserPile.PileInRack = userPile.PileInRack
                newUserPile.Type = userPile.Type
                newUserPile.NameModel = userPile.NameModel
                dictPile.Add(j, newUserPile)
                '=========================================================================================
                If countPiles > 1 Then
                    Dim dataColl As DataGridViewColumn = DGV_Piles.Columns.Item(1)
                    Dim newDataColl As DataGridViewColumn = New DataGridViewColumn()
                    newDataColl = dataColl.Clone()
                    Dim posRow As Integer = DGV_Piles.Columns.Add(newDataColl)
                    DGV_Piles.Columns.Item(countPiles).HeaderText = "Свая (Ряд №" & i & " Стобец №" & j & ")"
                    For k As Integer = 0 To DGV_Piles.RowCount - 1
                        Dim tag As String = DGV_Piles.Rows.Item(k).Tag
                        If tag Like "numberRow" Then
                            DGV_Piles.Rows(k).Cells(countPiles).Value = newUserPile.NumberRow
                        ElseIf tag Like "numberColl" Then
                            DGV_Piles.Rows(k).Cells(countPiles).Value = newUserPile.NumberColumn
                        Else
                            Dim copyData As String = DGV_Piles.Rows(k).Cells(1).Value
                            DGV_Piles.Rows(k).Cells(countPiles).Value = copyData
                        End If
                    Next k
                    countPiles += 1
                Else
                    DGV_Piles.Columns.Item(1).HeaderText = "Свая (Ряд №" & i & " Стобец №" & j & ")"
                    countPiles += 1
                End If
            Next j
            dictUserPile.Add(i, dictPile)
        Next i
    End Sub

    'смещение рада
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рассчитать угол по ряду
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim offsetRow As Double = NUpD_OffsetRowsPile.Value
        Dim numberRow As Integer = NUpD_CountRowsPile.Value
        Dim indexLenghtPile As Integer = -1
        Dim indexAngleY As Integer = -1
        Dim indexRow As Integer = -1
        If DGV_Piles.RowCount > 1 Then
            For i As Integer = 0 To DGV_Piles.RowCount - 1
                Dim tagRow As String = DGV_Piles.Rows.Item(i).Tag
                If tagRow Like "bridge_piles_length" Then
                    indexLenghtPile = i
                ElseIf tagRow Like "bridge_piles_angleY" Then
                    indexAngleY = i
                ElseIf tagRow Like "numberRow" Then
                    indexRow = i
                End If
            Next i
        End If
        If indexLenghtPile > -1 And indexAngleY > -1 And indexRow > -1 Then
            If DGV_Piles.ColumnCount > 1 Then
                For i As Integer = 1 To DGV_Piles.ColumnCount - 1
                    Dim numbRow As Integer = CInt(DGV_Piles.Rows(indexRow).Cells(i).Value)
                    If numbRow = numberRow Then
                        Dim pileLenght As Double = CDbl(DGV_Piles.Rows(indexLenghtPile).Cells(i).Value)
                        Dim lenOffset As Double = pileLenght / offsetRow
                        If pileLenght > 0 Then
                            Dim d As Double = pileLenght / lenOffset
                            Dim a As Double = Math.Atan(d)
                            Dim angle As Double = Math.PI / 2 - a
                            Dim deg As Double = (angle * 180) / Math.PI
                            deg = 90 - deg
                            DGV_Piles.Rows(indexAngleY).Cells(i).Value = Math.Round(deg, 2)
                        End If
                    End If
                Next i
            End If
            MsgBox("Данные успешно записаны!!!")
        Else
            MsgBox("Не удалось записать смещение для выбранного ряда!!!")
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'смещение столбца
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim offsetColumn As Double = NUpD_OffsetColumnsPile.Value
        Dim numberColumn As Integer = NUpD_CountColumnsPile.Value
        Dim indexLenghtPile As Integer = -1
        Dim indexAngleX As Integer = -1
        Dim indexColl As Integer = -1
        If DGV_Piles.RowCount > 1 Then
            For i As Integer = 0 To DGV_Piles.RowCount - 1
                Dim tagRow As String = DGV_Piles.Rows.Item(i).Tag
                If tagRow Like "bridge_piles_length" Then
                    indexLenghtPile = i
                ElseIf tagRow Like "bridge_piles_angleX" Then
                    indexAngleX = i
                ElseIf tagRow Like "numberColl" Then
                    indexColl = i
                End If
            Next i
        End If
        If indexLenghtPile > -1 And indexColl > -1 And indexAngleX > -1 Then
            If DGV_Piles.ColumnCount > 1 Then
                For i As Integer = 1 To DGV_Piles.ColumnCount - 1
                    Dim numbColl As Integer = CInt(DGV_Piles.Rows(indexColl).Cells(i).Value)
                    If numbColl = numberColumn Then
                        Dim pileLenght As Double = CDbl(DGV_Piles.Rows(indexLenghtPile).Cells(i).Value)
                        Dim lenOffset As Double = pileLenght / offsetColumn
                        If pileLenght > 0 Then
                            Dim d As Double = pileLenght / lenOffset
                            Dim a As Double = Math.Atan(d)
                            Dim angle As Double = Math.PI / 2 - a
                            Dim deg As Double = (angle * 180) / Math.PI
                            deg = 90 - deg
                            DGV_Piles.Rows(indexAngleX).Cells(i).Value = Math.Round(deg, 2)
                        End If
                    End If
                Next i
            End If
            MsgBox("Данные успешно записаны!!!")
        Else
            MsgBox("Не удалось записать смещение для выбранного ряда!!!")
        End If
    End Sub

    'отметка земли
    Private Sub ChB_EgSurface_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_ElevationLand.CheckedChanged
        If ChB_ElevationLand.Checked = True Then
            CB_EgSurface.Enabled = False
            NUpD_ElevationLand.Enabled = True
        Else
            CB_EgSurface.Enabled = True
            NUpD_ElevationLand.Enabled = False
        End If
    End Sub

    Private Sub DGV_CabinetWall_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_CabinetWall.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_CabinetWall.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_CabinetWall.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_CabinetWall.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_Grillage_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_Grillage.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_Grillage.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_Grillage.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_Grillage.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_LeftHand_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_LeftHand.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_LeftHand.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_LeftHand.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_LeftHand.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_LeftPostcard_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_LeftPostcard.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_LeftPostcard.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_LeftPostcard.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_LeftPostcard.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_Nozzle_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_Nozzle.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If e.ColumnIndex < 1 Then
            Return
        End If
        If TypeOf DGV_Nozzle.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_Nozzle.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_Nozzle.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_Piles_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_Piles.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_Piles.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_Piles.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_Piles.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_Preparation_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_Preparation.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_Preparation.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_Preparation.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_Preparation.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_Rack_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_Rack.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_Rack.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_Rack.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_Rack.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_RightHand_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_RightHand.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_RightHand.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_RightHand.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_RightHand.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_RightPostcard_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_RightPostcard.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_RightPostcard.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_RightPostcard.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_RightPostcard.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub DGV_SubFermenter_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DGV_SubFermenter.CellValidating
        ' Проверяем, что это ячейка текстового типа
        If e.RowIndex < 0 Or e.ColumnIndex < 1 Then
            Return ' Игнорируем заголовки
        End If
        If TypeOf DGV_SubFermenter.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn Then
            Dim value As String = e.FormattedValue.ToString()
            ' Если в значении есть запятая
            If value.Contains(",") Then
                ' Заменяем запятую на точку
                Dim correctedValue As String = value.Replace(",", ".")
                If IsNumeric(correctedValue) Then
                    ' Если после замены значение является числом, сохраняем его
                    DGV_SubFermenter.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = correctedValue
                Else
                    DGV_SubFermenter.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
                End If
            End If
        End If
    End Sub

    Private Sub ChB_EgeParallel_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_EgeParallel.CheckedChanged

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'не создавать стойки
    Private Sub ChB_CreateRack_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_CreateRack.CheckedChanged
        If ChB_CreateRack.Checked = True Then
            ChB_CreateGrillage.Checked = True
        Else
            ChB_CreateGrillage.Checked = False
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка вставить сваю в стойку
    Private Sub ChB_InsertPileInRack_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_InsertPileInRack.CheckedChanged
        If ChB_InsertPileInRack.Checked = True Then
            TxtB_PileCollDiagram.Enabled = False
            TxtB_PileRowDiagram.Enabled = False
            'смотрим количество стоек
            Dim countRack As Integer = NUpD_CountRack.Value
            'удаляем лишние столбцы
            If DGV_Piles.ColumnCount > 2 Then
                For i As Integer = DGV_Piles.ColumnCount - 1 To 2 Step -1
                    DGV_Piles.Columns.RemoveAt(i)
                Next i
            End If

            For i As Integer = 1 To countRack
                If i = 1 Then
                    DGV_Piles.Columns.Item(i).HeaderText = "Свая (Ряд №1, Столбец №" & i & ")"
                Else
                    Dim dataColl As DataGridViewColumn = DGV_Piles.Columns.Item(1)
                    Dim newDataColl As DataGridViewColumn = New DataGridViewColumn()
                    newDataColl = dataColl.Clone()
                    Dim posRow As Integer = DGV_Piles.Columns.Add(newDataColl)
                    DGV_Piles.Columns.Item(i).HeaderText = "Свая (Ряд №1, Столбец №" & i & ")"
                End If
                For j As Integer = 0 To DGV_Piles.RowCount - 1
                    Dim tagRow As String = DGV_Piles.Rows(j).Tag
                    If tagRow Like "numberRow" Then
                        DGV_Piles.Rows(j).Cells(i).Value = 1
                    ElseIf tagRow Like "numberColl" Then
                        DGV_Piles.Rows(j).Cells(i).Value = i
                    Else
                        Dim copyData As String = DGV_Piles.Rows(j).Cells(1).Value
                        DGV_Piles.Rows(j).Cells(i).Value = copyData
                    End If
                Next j
            Next i
            ChB_CreateGrillage.Checked = True
            ChB_CreatePreparation.Checked = True
        Else
            TxtB_PileCollDiagram.Enabled = True
            TxtB_PileRowDiagram.Enabled = True
            If ChB_CreateRack.Checked = False Then
                ChB_CreateGrillage.Checked = False
                ChB_CreatePreparation.Checked = False
            End If
        End If
    End Sub

    'выбрана насадка
    Private Sub DGV_Nozzle_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Nozzle.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_Nozzle.Rows(indexRow).Tag
                nameElement = "Nozzle"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_CabinetWall_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_CabinetWall.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_CabinetWall.Rows(indexRow).Tag
                nameElement = "CabinetWall"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_Grillage_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Grillage.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_Grillage.Rows(indexRow).Tag
                nameElement = "Grillage"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_LeftHand_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_LeftHand.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_LeftHand.Rows(indexRow).Tag
                nameElement = "LeftHand"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_LeftPostcard_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_LeftPostcard.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_LeftPostcard.Rows(indexRow).Tag
                nameElement = "LeftPostcard"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_RightHand_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_RightHand.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_RightHand.Rows(indexRow).Tag
                nameElement = "RightHand"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_RightPostcard_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_RightPostcard.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_RightPostcard.Rows(indexRow).Tag
                nameElement = "RightPostcard"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_Rack_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Rack.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_Rack.Rows(indexRow).Tag
                nameElement = "Rack"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_Preparation_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Preparation.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_Preparation.Rows(indexRow).Tag
                nameElement = "Preparation"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub DGV_Piles_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Piles.CellClick
        Try
            Dim indexRow As Integer = e.RowIndex
            If indexRow > -1 Then
                tagSelect = DGV_Piles.Rows(indexRow).Tag
                nameElement = "Piles"
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub
    'заполнить таблицу подферменников согласно 1 элемента
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If DGV_SubFermenter.ColumnCount > 2 Then
            If DGV_SubFermenter.RowCount > 2 Then
                For i As Integer = 0 To DGV_SubFermenter.RowCount - 1
                    Dim nameField As String = DGV_SubFermenter.Rows(i).Tag
                    Dim value As String = ""
                    If nameField Like "numberProlet" Then
                        Continue For
                    ElseIf nameField Like "numberRow" Then
                        For j As Integer = 1 To DGV_SubFermenter.ColumnCount - 1
                            Dim numberRow As Integer = DGV_SubFermenter.Columns.Item(j).Tag
                            DGV_SubFermenter.Rows(i).Cells(j).Value = CalculationBeams.getConditionalRow(numberRow)
                        Next j
                        Continue For
                    Else
                        value = DGV_SubFermenter.Rows(i).Cells(1).Value
                    End If
                    For j As Integer = 2 To DGV_SubFermenter.ColumnCount - 1
                        DGV_SubFermenter.Rows(i).Cells(j).Value = value
                    Next j
                Next i
            End If
        End If
    End Sub

    Private Sub DGV_Nozzle_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV_Nozzle.CellContentClick

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles ChB_LeftPostcsrdFixedLenght.CheckedChanged

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click

    End Sub
End Class