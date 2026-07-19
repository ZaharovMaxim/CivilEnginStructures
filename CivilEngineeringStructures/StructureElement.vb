Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Interop
Imports CivilEnginStructures.StructureElement
Imports Microsoft.Office.Core
Imports NetTopologySuite.Algorithm
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Topomatic.Cad.Foundation
Imports Topomatic.Crs.Rail
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class StructureElement
    Public Enum classBridge
        <Description("Мостовое сооружение")> Bridge = 0
        <Description("Мостовое полотно")> BridgeDesk = 1
        <Description("Крайние опоры")> LastPillars = 2
        <Description("Промежуточные опоры")> MiddlePillars = 3
        <Description("Пролетное строение")> SpanStructures = 4
        <Description("Прочие элементы")> OtherElements = 10
    End Enum
    Public Enum classStructure
        <Description("Мостовое сооружение")> Bridges = 0
        <Description("Крайняя опора")> LastPillar = 1
        <Description("Промежуточная опора")> MiddlePillar = 2
        <Description("Насадка")> NozzlePillar = 3
        <Description("Ригель")> RigelPillar = 4
        <Description("Группа подферменников")> GroupSubFermenters = 5
        <Description("Подферменник")> SubFermenters = 6
        <Description("Шкафная стенка")> CabinetWallPillar = 7
        <Description("Откосное крыло левое")> HandLeftPillar = 8
        <Description("Откосное крыло правое")> HandRightPillar = 9
        <Description("Обратный открылок левый")> PostcardLeftPillar = 10
        <Description("Обратный открылок правый")> PostcardRightPillar = 11
        <Description("Группа стоек")> GroupRackPillar = 12
        <Description("Стойка")> RackPillar = 13
        <Description("Ростверк")> GrillagePillar = 14
        <Description("Подготовка")> PreparationPillar = 15
        <Description("Группа свай")> GroupPilePillar = 16
        <Description("Свая")> PilePillar = 17
        <Description("Балка двутавровая")> BeamI = 18
        <Description("Участок омоноличивания балок")> SitesBeamsMonolit = 19
        <Description("Участок омоноличивания опоры")> SitesPillarMonolit = 20
        <Description("Граница сооружения")> BoundBridge = 21
        <Description("Иной объект")> OtherObject = 22
    End Enum

    ' Второй enum (со смещением на 100)
    Public Enum typeObject
        <Description("Ось мостового сооружения")> axisBridge = 100
        <Description("Ось траектории раскладки балок")> axisTrajectoryPlacementBeams = 101
        <Description("Ось опирания балок")> axisPillarBeams = 102
        <Description("Ось балки")> axisBeam = 103
        <Description("Ось опоры")> axisPillar = 104
        <Description("Ось насадки")> axisNozzle = 105
        <Description("Ось ригеля")> axisRigel = 106
        <Description("Ось подферменника")> axisSubFermenters = 107
        <Description("Ось стойки")> axisRack = 108
        <Description("Ось ростверка")> axisGrillage = 109
        <Description("Ось ледореза")> axisIcecutter = 110
        <Description("Ось подготовки")> axisPreparation = 111
        <Description("Ось сваи")> axisPile = 112
        <Description("Ось левого откосного крыла")> axisLeftHand = 113
        <Description("Ось правого откосного крыла")> axisRightHand = 114
        <Description("Ось левого обратного открылка")> axisLeftPostcard = 115
        <Description("Ось правого обратного открылка")> axisRightPostcard = 116
        <Description("Ось шкафной стенки")> axisCabinetWall = 117
        <Description("Ось участка омоноличивания балок")> axisSiteMonolitBeams = 118
        <Description("Ось участка омоноличивания опоры")> axisSiteMonolitPillar = 119

        <Description("Контур насадки (верх)")> contourNozzleTop = 120
        <Description("Контур насадки (низ)")> contourNozzleBottom = 121
        <Description("Линия начала размещения шкафной стенки")> contourNozzleCabinetWall = 122
        <Description("Контур левой консоли")> contourNozzleLeftConsole = 123
        <Description("Контур правой консоли")> contourNozzleRightConsole = 124
        <Description("Контур шкафной стенки (верх)")> contourCabinetWallTop = 125
        <Description("Контур шкафной стенки (низ)")> contourCabinetWallBottom = 126
        <Description("Контур зуба упора (верх)")> contourCabinetWallPlateTop = 127
        <Description("Контур зуба упора (низ)")> contourCabinetWallPlateBottom = 128

        <Description("Контур ростверка (верх)")> counterGrillageTop = 129
        <Description("Контур ростверка (низ)")> counterGrillageBottom = 130
        <Description("Контур левого откосного крыла (верх)")> counterLeftHandTop = 131
        <Description("Контур левого откосного крыла (низ)")> counterLeftHandBottom = 132
        <Description("Контур карниза левого откосного крыла (верх)")> counterLeftHandCorniceTop = 133
        <Description("Контур карниза левого откосного крыла (низ)")> counterLeftHandCorniceBottom = 134
        <Description("Контур правого откосного крыла (верх)")> counterRightHandTop = 135
        <Description("Контур правого откосного крыла (низ)")> counterRightHandBottom = 136
        <Description("Контур карниза правого откосного крыла (верх)")> counterRightHandCorniceTop = 137
        <Description("Контур карниза правого откосного крыла (низ)")> counterRightHandCorniceBottom = 138
        <Description("Контур левого обратного открылка (верх)")> counterLeftPostcardTop = 139
        <Description("Контур левого обратного открылка (низ)")> counterLeftPostcardBottom = 140
        <Description("Контур правого обратного открылка (верх)")> counterRightPostcardTop = 141
        <Description("Контур правого обратного открылка (низ)")> counterRightPostcardBottom = 142
        <Description("Контур сваи (верх)")> counterPileTop = 143
        <Description("Контур сваи (низ)")> counterPileBottom = 144
        <Description("Контур ригеля (верх)")> counterRigelTop = 145
        <Description("Контур ригеля (низ)")> counterRigelBottom = 146
        <Description("Контур подферменника (верх)")> counterSubFermentersTop = 147
        <Description("Контур подферменника (низ)")> counterSubFermentersBottom = 148
        <Description("Контур стойки (верх)")> counterRackTop = 149
        <Description("Контур стойки (низ)")> counterRackBottom = 150
        <Description("Контур ледореза (верх)")> counterIcecutterTop = 151
        <Description("Контур ледореза (низ)")> counterIcecutterBottom = 152
        <Description("Контур подготовки (верх)")> counterPreparationTop = 153
        <Description("Контур подготовки (низ)")> counterPreparationBottom = 154
        <Description("Верх грани плиты балки")> counterTopBeam = 155
        <Description("Низ ребра балки")> counterBottomBeam = 156
        <Description("Контур участка омоноличивания балок (верх)")> counterSiteMonolitBeamsTop = 157
        <Description("Контур участка омоноличивания балок (низ)")> counterSiteMonolitBeamsBottom = 158
        <Description("Контур участка омоноличивания опор (верх)")> counterSiteMonolitPillarTop = 159
        <Description("Контур участка омоноличивания опор (низ)")> counterSiteMonolitPillarBottom = 160
        <Description("Область участка омоноличивания балок")> hatchSiteMonolitBeams = 161
        <Description("Область участка омоноличивания опор")> hatchSiteMonolitPillar = 162

        <Description("Балка двутавровая (модель)")> modelBeam = 163
        <Description("Насадка (модель)")> modelNozzle = 164
        <Description("Ригель (модель)")> modelRigel = 165
        <Description("Шкафная стенка (модель)")> modelCabinetWall = 166
        <Description("Зуб упора (модель)")> modelCabinetWallPlate = 167
        <Description("Ростверк (модель)")> modelGrillage = 168
        <Description("Свая (модель)")> modelPile = 169
        <Description("Стойка (модель)")> modelRack = 170
        <Description("Откосное крыло левое (модель)")> modelLeftHand = 171
        <Description("Карниз откосного крыла левого (модель)")> modelLeftHandCornice = 172
        <Description("Откосное крыло правое (модель)")> modelRightHand = 173
        <Description("Карниз откосного крыла правого (модель)")> modelRightHandCornice = 174
        <Description("Обратный открылок левый (модель)")> modelLeftPostcard = 175
        <Description("Обратный открылок правый (модель)")> modelRightPostcard = 176
        <Description("Подферменник (модель)")> modelSubFermenters = 177
        <Description("Подготовка (модель)")> modelPreparation = 178
        <Description("Ледорез (модель)")> modelIcecutter = 179
        <Description("Участок омоноличивания балок")> modelSiteMonolitBeams = 180
        <Description("Участок омоноличивания опор")> modelSiteMonolitPillar = 181

        <Description("Границы мостового сооружения")> boundaresBridge = 182
        <Description("Иной объект")> OtherElement = 183
    End Enum
    Public Enum tableXRecords
        PROJECT_BRIDGE
        PROJECT_STRUCTURES
    End Enum
    Private _label As String
    Private _classBridge As classBridge
    Private _classStructure As classStructure
    Private _name As typeObject
    Private _description As String
    Private _keyParameter As String
    Private _idStructure As String
    Private _idElement As String
    Private _note As String
    Private _idObject As String
    Private _entity As DwgEntity
    ' Конструктор по умолчанию
    Public Sub New()
        _label = String.Empty
        _classBridge = classBridge.OtherElements
        _classStructure = classStructure.OtherObject
        _name = typeObject.OtherElement
        _description = String.Empty
        _keyParameter = String.Empty
        _idStructure = String.Empty
        _idElement = String.Empty
        _note = String.Empty
        _idObject = String.Empty
        _entity = Nothing
    End Sub

    ' Конструктор с параметрами
    Public Sub New(label As String, ClassObject As classStructure, name As typeObject, description As String, keyParameter As String,
                   idStructure As String, idElement As String, note As String, idObject As String)
        _label = label
        _classStructure = ClassObject
        _name = name
        _description = description
        _keyParameter = keyParameter
        _idStructure = idStructure
        _idElement = idElement
        _note = note
        _idObject = idObject
    End Sub
    ' Свойство label
    Public Property Label As String
        Get
            Return _label
        End Get
        Set(value As String)
            _label = value
        End Set
    End Property
    ' Свойство 
    Public Property ClassBridgeObject As classBridge
        Get
            Return _classBridge
        End Get
        Set(value As classBridge)
            _classStructure = value
        End Set
    End Property

    ' Свойство Class
    Public Property ClassObject As classStructure
        Get
            Return _classStructure
        End Get
        Set(value As classStructure)
            _classStructure = value
        End Set
    End Property
    ' Свойство Name
    Public Property Name As typeObject
        Get
            Return _name
        End Get
        Set(value As typeObject)
            _name = value
        End Set
    End Property
    ' Свойство Description (бывшее desk)
    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = value
        End Set
    End Property
    ' Свойство KeyParameter
    Public Property KeyParameter As String
        Get
            Return _keyParameter
        End Get
        Set(value As String)
            _keyParameter = value
        End Set
    End Property
    ' Свойство IdBridge
    Public Property IdStructure As String
        Get
            Return _idStructure
        End Get
        Set(value As String)
            _idStructure = value
        End Set
    End Property
    ' Свойство IdElement
    Public Property IdElement As String
        Get
            Return _idElement
        End Get
        Set(value As String)
            _idElement = value
        End Set
    End Property
    ' Свойство Note
    Public Property Note As String
        Get
            Return _note
        End Get
        Set(value As String)
            _note = value
        End Set
    End Property
    ' Свойство IdObject
    Public Property IdObject As String
        Get
            Return _idObject
        End Get
        Set(value As String)
            _idObject = value
        End Set
    End Property
    ' Свойство IdObject
    '<JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
    Public Property DWGEntity As DwgEntity
        Get
            Return _entity
        End Get
        Set(value As DwgEntity)
            _entity = value
        End Set
    End Property

    ' Метод для проверки валидности
    Public Function IsValid() As Boolean
        Return Not String.IsNullOrWhiteSpace(_name) AndAlso
               Not String.IsNullOrWhiteSpace(_idStructure)
    End Function
    Public Function getNameStructure(ByVal deskElement As String) As StructureElement.typeObject
        getNameStructure = Nothing
        If IsNothing(deskElement) = False Then
            If deskElement.Trim.Length > 0 Then
                Select Case deskElement
                    Case "Главная ось сооружения"
                        Return StructureElement.typeObject.axisPillar
                    Case "Граница сооружения"
                        Return StructureElement.typeObject.boundaresBridge
                    Case Else
                        Return Nothing
                End Select
            End If
        End If
    End Function
    'получить достук элементу типа МОСТ
    Public Function getBridge() As Bridges
        Dim userBridge As Bridges = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userBridge = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Bridges)(KeyParameter)
            'userBridge.DWGEntity = DWGEntity
        End If
        Return userBridge
    End Function
    'получить достук элементу типа ОСЬ ОПОРЫ
    Public Function getPillar() As Pillar
        Dim userPillar As Pillar = Nothing
        If IsNothing(KeyParameter) = False Then
            If FuncGSON.IsValidJson(KeyParameter) = True Then
                userPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Pillar)(KeyParameter)
            End If
        End If
        Return userPillar
    End Function
    'получить достук элементу типа ОСЬ ОПИРАНИЯ БАЛОК
    Public Function getAxisBeamsPillar() As AxisBeamsPillars
        Dim userAxisBeams As AxisBeamsPillars = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userAxisBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisBeamsPillars)(KeyParameter)
            'userAxisBeams.EntityAxisBeams = DWGEntity
        End If
        Return userAxisBeams
    End Function
    'получить достук элементу типа БАЛКА
    Public Function getBeamI() As BeamI
        Dim userBeam As BeamI = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of BeamI)(KeyParameter)
            'userBeam.EntityBeamI = DWGEntity
        End If
        Return userBeam
    End Function
    'получить достук элементу типа КОНТУР БАЛКИ
    Public Function getCounterBeam() As CounterBeam
        Dim userCounterBeam As CounterBeam = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCounterBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CounterBeam)(KeyParameter)
        End If
        Return userCounterBeam
    End Function

    'получить достук элементу типа БАЛКА
    Public Function getModelBeam() As ModelBeam
        Dim userModelBeam As ModelBeam = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userModelBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ModelBeam)(KeyParameter)
        End If
        Return userModelBeam
    End Function

    'получить достук элементу типа Ось раскладки балок
    Public Function getAxisPlacementBeams() As TrajectoryPlacementBeams
        Dim userAxisPlacementBeams As TrajectoryPlacementBeams = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userAxisPlacementBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of TrajectoryPlacementBeams)(KeyParameter)
        End If
        Return userAxisPlacementBeams
    End Function

    'получить достук элементу типа Ось омоноличивания балок
    Public Function getMonolitSiteBeams() As SiteMonolitBeams
        Dim userSiteMonolitBeams As SiteMonolitBeams = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userSiteMonolitBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SiteMonolitBeams)(KeyParameter)
        End If
        Return userSiteMonolitBeams
    End Function
    Public Function getHatchMonolitSiteBeams() As HatchSiteMonolitBeams
        Dim userHatchSiteMonolit As HatchSiteMonolitBeams = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userHatchSiteMonolit = Newtonsoft.Json.JsonConvert.DeserializeObject(Of HatchSiteMonolitBeams)(KeyParameter)
        End If
        Return userHatchSiteMonolit
    End Function
    Public Function getCounterMonolitSiteBeams() As CounterSiteMonolitBeams
        Dim userCounterSiteMonolitBeams As CounterSiteMonolitBeams = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCounterSiteMonolitBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CounterSiteMonolitBeams)(KeyParameter)
        End If
        Return userCounterSiteMonolitBeams
    End Function
    Public Function getModelSiteMonolitBeams() As ModelSiteMonolitBeams
        Dim userModelSiteMonolitBeams As ModelSiteMonolitBeams = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userModelSiteMonolitBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ModelSiteMonolitBeams)(KeyParameter)
        End If
        Return userModelSiteMonolitBeams
    End Function
    Public Function getNozzlePillar() As NozzlePillar
        Dim userNozzlePillar As NozzlePillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userNozzlePillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of NozzlePillar)(KeyParameter)
        End If
        Return userNozzlePillar
    End Function
    Public Function getNozzleCounterPillar() As NozzleContour
        Dim userNozzleCounterPillar As NozzleContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userNozzleCounterPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of NozzleContour)(KeyParameter)
        End If
        Return userNozzleCounterPillar
    End Function
    Public Function getNozzleModelPillar() As NozzleModel
        Dim userNozzleModelPillar As NozzleModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userNozzleModelPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of NozzleModel)(KeyParameter)
        End If
        Return userNozzleModelPillar
    End Function
    Public Function getCabinetWallPillar() As CabinetWallPillar
        Dim userCabinetWallPillar As CabinetWallPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCabinetWallPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CabinetWallPillar)(KeyParameter)
        End If
        Return userCabinetWallPillar
    End Function
    Public Function getCabinetWallCounterPillar() As CabinetWallContour
        Dim userCabinetWallCounterPillar As CabinetWallContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCabinetWallCounterPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CabinetWallContour)(KeyParameter)
        End If
        Return userCabinetWallCounterPillar
    End Function
    Public Function getCabinetWallModelPillar() As CabinetWallModel
        Dim userCabinetWallModelPillar As CabinetWallModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCabinetWallModelPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CabinetWallModel)(KeyParameter)
        End If
        Return userCabinetWallModelPillar
    End Function
    Public Function getGrillagePillar() As GrillagePillar
        Dim userGrillagePillar As GrillagePillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userGrillagePillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of GrillagePillar)(KeyParameter)
        End If
        Return userGrillagePillar
    End Function
    Public Function getGrillageCounterPillar() As GrillageContour
        Dim userCounterPillar As GrillageContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userCounterPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of GrillageContour)(KeyParameter)
        End If
        Return userCounterPillar
    End Function
    Public Function getGrillageModelPillar() As GrillageModel
        Dim userModelPillar As GrillageModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userModelPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of GrillageModel)(KeyParameter)
        End If
        Return userModelPillar
    End Function
    Public Function getHandPillar() As HandPillar
        Dim userHandPillar As HandPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userHandPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of HandPillar)(KeyParameter)
        End If
        Return userHandPillar
    End Function
    Public Function getHandCounterPillar() As HandContour
        Dim userHandCounter As HandContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userHandCounter = Newtonsoft.Json.JsonConvert.DeserializeObject(Of HandContour)(KeyParameter)
        End If
        Return userHandCounter
    End Function
    Public Function getHandModelPillar() As HandModel
        Dim userHandModel As HandModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userHandModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of HandModel)(KeyParameter)
        End If
        Return userHandModel
    End Function
    Public Function getPostcardPillar() As PostcardPillar
        Dim userPostcardPillar As PostcardPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPostcardPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PostcardPillar)(KeyParameter)
        End If
        Return userPostcardPillar
    End Function
    Public Function getPostcardCounterPillar() As PostcardContour
        Dim userPostcardCounter As PostcardContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPostcardCounter = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PostcardContour)(KeyParameter)
        End If
        Return userPostcardCounter
    End Function
    Public Function getPostcardModelPillar() As PostcardModel
        Dim userPostcardModel As PostcardModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPostcardModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PostcardModel)(KeyParameter)
        End If
        Return userPostcardModel
    End Function
    Public Function getPilePillar() As PilePillar
        Dim userPilePillar As PilePillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPilePillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PilePillar)(KeyParameter)
        End If
        Return userPilePillar
    End Function
    Public Function getPileCounterPillar() As PileContour
        Dim userPileCounter As PileContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPileCounter = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PileContour)(KeyParameter)
        End If
        Return userPileCounter
    End Function
    Public Function getPileModelPillar() As PileModel
        Dim userPileModel As PileModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPileModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PileModel)(KeyParameter)
        End If
        Return userPileModel
    End Function
    Public Function getPreparationPillar() As PreparationPillar
        Dim userPreparationPillar As PreparationPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPreparationPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PreparationPillar)(KeyParameter)
        End If
        Return userPreparationPillar
    End Function
    Public Function getPreparationContourPillar() As PreparationContour
        Dim userPreparationContour As PreparationContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPreparationContour = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PreparationContour)(KeyParameter)
        End If
        Return userPreparationContour
    End Function
    Public Function getPreparationModelPillar() As PreparationModel
        Dim userPreparationModel As PreparationModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userPreparationModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of PreparationModel)(KeyParameter)
        End If
        Return userPreparationModel
    End Function
    Public Function getRackPillar() As RackPillar
        Dim userRackPillar As RackPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRackPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RackPillar)(KeyParameter)
        End If
        Return userRackPillar
    End Function
    Public Function getContourRackPillar() As RackContour
        Dim userRackContour As RackContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRackContour = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RackContour)(KeyParameter)
        End If
        Return userRackContour
    End Function
    Public Function getRackModelPillar() As RackModel
        Dim userRackModel As RackModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRackModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RackModel)(KeyParameter)
        End If
        Return userRackModel
    End Function
    Public Function getRigelPillar() As RigelPillar
        Dim userRigelPillar As RigelPillar = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRigelPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RigelPillar)(KeyParameter)
        End If
        Return userRigelPillar
    End Function
    Public Function getRigelContourPillar() As RigelContour
        Dim userRigelContour As RigelContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRigelContour = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RigelContour)(KeyParameter)
        End If
        Return userRigelContour
    End Function
    Public Function getRigelModelPillar() As RigelModel
        Dim userRigelModel As RigelModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userRigelModel = Newtonsoft.Json.JsonConvert.DeserializeObject(Of RigelModel)(KeyParameter)
        End If
        Return userRigelModel
    End Function
    Public Function getSubFermenters() As SubFermenters
        Dim userSubFermenters As SubFermenters = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userSubFermenters = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SubFermenters)(KeyParameter)
        End If
        Return userSubFermenters
    End Function
    Public Function getSubFermentersContour() As SubFermenterContour
        Dim userSubFermenters As SubFermenterContour = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userSubFermenters = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SubFermenterContour)(KeyParameter)
        End If
        Return userSubFermenters
    End Function
    Public Function getSubFermentersModel() As SubFermenterModel
        Dim userSubFermenters As SubFermenterModel = Nothing
        If FuncGSON.IsValidJson(KeyParameter) = True Then
            userSubFermenters = Newtonsoft.Json.JsonConvert.DeserializeObject(Of SubFermenterModel)(KeyParameter)
        End If
        Return userSubFermenters
    End Function
    'функция получает парамерры класса выбранного 
    Public Shared Function getStructureElement(Of T As Class)(ByVal entity As DwgEntity) As T
        If IsNothing(entity) = False Then
            Dim elementBridge As StructureElement = New StructureElement()
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(entity, elementBridge)
            If boolFindData = True Then
                Dim keyParam As String = elementBridge.KeyParameter
                Dim result As T = Newtonsoft.Json.JsonConvert.DeserializeObject(Of T)(keyParam)
                'elementBridge.DWGEntity = entity
                Return result
            End If
        End If
        Return Nothing
    End Function

    'функция удаляет элемент
    Public Shared Function isValidateDataStructure(ByRef listData As List(Of StructureElement)) As StructureElement
        Dim result As StructureElement = Nothing
        If IsNothing(listData) = False Then
            For i As Integer = 0 To listData.Count - 1
                Dim tempData As StructureElement = listData.Item(i)
                If i = 0 Then
                    result = tempData
                Else
                    Try
                        Dim userObject As DwgEntity = tempData.DWGEntity
                        Dim drawDoc As Drawing = userObject.Drawing
                        drawDoc.ActiveSpace.Entities.Remove(userObject)
                    Catch ex As System.Exception
                    End Try
                End If
            Next i
        End If
        Return result
    End Function

    Public Shared Function GetDescription(value As System.Enum) As String
        Dim fieldInfo = value.GetType().GetField(value.ToString())
        Dim attributes = DirectCast(fieldInfo.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        If attributes IsNot Nothing AndAlso attributes.Length > 0 Then
            Return attributes(0).Description
        End If
        Return value.ToString()
    End Function

    Public Shared Function GetDescriptions(enumType As Type) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)()
        For Each value As System.Enum In [Enum].GetValues(enumType)
            Dim description = GetDescription(value)
            result.Add(value.ToString(), description)
        Next
        Return result
    End Function

    Public Shared Function ConvertToEnum(Of T As Structure)(value As Object, Optional defaultValue As T = Nothing) As T
        ' Проверяем, что T действительно enum
        If Not GetType(T).IsEnum Then
            Return defaultValue
        End If

        ' Если значение Nothing или DBNull
        If value Is Nothing OrElse IsDBNull(value) Then
            Return defaultValue
        End If

        Try
            ' Если значение уже нужного типа
            If TypeOf value Is T Then
                Return DirectCast(value, T)
            End If

            ' Проверяем строку
            If TypeOf value Is String Then
                Dim strValue As String = value.ToString().Trim()

                ' Если строка пустая
                If String.IsNullOrEmpty(strValue) Then
                    Return defaultValue
                End If

                ' Пробуем преобразовать по имени
                Dim result As T
                If [Enum].TryParse(strValue, True, result) Then
                    ' Проверяем, что значение определено в enum
                    If [Enum].IsDefined(GetType(T), result) Then
                        Return result
                    End If
                End If

                ' Пробуем преобразовать по числовому значению (если строка - число)
                Dim numericValue As Integer
                If Integer.TryParse(strValue, numericValue) Then
                    If [Enum].IsDefined(GetType(T), numericValue) Then
                        Return DirectCast([Enum].ToObject(GetType(T), numericValue), T)
                    End If
                End If
            End If

            ' Если значение - число
            If TypeOf value Is Integer OrElse TypeOf value Is Long OrElse
               TypeOf value Is Short OrElse TypeOf value Is Byte Then
                Dim numericValue As Integer = Convert.ToInt32(value)
                If [Enum].IsDefined(GetType(T), numericValue) Then
                    Return DirectCast([Enum].ToObject(GetType(T), numericValue), T)
                End If
            End If

        Catch ex As Exception
        End Try

        Return defaultValue
    End Function

    Public Shared Function ConvertToClassStructure(value As Object, Optional defaultValue As classStructure = classStructure.OtherObject) As classStructure
        Return ConvertToEnum(Of classStructure)(value, defaultValue)
    End Function

    ' Для typeObject
    Public Shared Function ConvertToTypeObject(value As Object, Optional defaultValue As typeObject = typeObject.OtherElement) As typeObject
        Return ConvertToEnum(Of typeObject)(value, defaultValue)
    End Function

    ' Для typeRack
    Public Shared Function ConvertToTypeRack(value As Object, Optional defaultValue As RackPillar.TypeRack = RackPillar.TypeRack.None) As RackPillar.TypeRack
        Return ConvertToEnum(Of RackPillar.TypeRack)(value, defaultValue)
    End Function
End Class

Public Class ObjectFactory
    ' Словарь соответствий между enum и типом объекта
    Private Shared ReadOnly _typeMappings As New Dictionary(Of typeObject, Type) From {
        {typeObject.axisBridge, GetType(Bridges)},
        {typeObject.axisTrajectoryPlacementBeams, GetType(TrajectoryPlacementBeams)},
        {typeObject.axisPillarBeams, GetType(AxisBeamsPillars)},
        {typeObject.axisBeam, GetType(BeamI)},
        {typeObject.axisPillar, GetType(Pillar)},
        {typeObject.axisNozzle, GetType(NozzlePillar)},
        {typeObject.axisRigel, GetType(RigelPillar)},
        {typeObject.axisSubFermenters, GetType(SubFermenters)},
        {typeObject.axisRack, GetType(RackPillar)},
        {typeObject.axisGrillage, GetType(GrillagePillar)},
        {typeObject.axisIcecutter, GetType(IcecutterPillar)},
        {typeObject.axisPreparation, GetType(PreparationPillar)},
        {typeObject.axisPile, GetType(PilePillar)},
        {typeObject.axisLeftHand, GetType(HandPillar)},
        {typeObject.axisRightHand, GetType(HandPillar)},
        {typeObject.axisLeftPostcard, GetType(PostcardPillar)},
        {typeObject.axisRightPostcard, GetType(PostcardPillar)},
        {typeObject.axisCabinetWall, GetType(CabinetWallPillar)},
        {typeObject.axisSiteMonolitBeams, GetType(SiteMonolitBeams)},
        {typeObject.contourNozzleTop, GetType(NozzleContour)},
        {typeObject.contourNozzleBottom, GetType(NozzleContour)},
        {typeObject.contourNozzleCabinetWall, GetType(NozzleContour)},
        {typeObject.contourNozzleLeftConsole, GetType(NozzleContour)},
        {typeObject.contourNozzleRightConsole, GetType(NozzleContour)},
        {typeObject.contourCabinetWallTop, GetType(CabinetWallContour)},
        {typeObject.contourCabinetWallBottom, GetType(CabinetWallContour)},
        {typeObject.contourCabinetWallPlateTop, GetType(CabinetWallContour)},
        {typeObject.contourCabinetWallPlateBottom, GetType(CabinetWallContour)},
        {typeObject.counterGrillageTop, GetType(GrillageContour)},
        {typeObject.counterGrillageBottom, GetType(GrillageContour)},
        {typeObject.counterLeftHandTop, GetType(HandContour)},
        {typeObject.counterLeftHandBottom, GetType(HandContour)},
        {typeObject.counterLeftHandCorniceTop, GetType(HandContour)},
        {typeObject.counterLeftHandCorniceBottom, GetType(HandContour)},
        {typeObject.counterRightHandTop, GetType(HandContour)},
        {typeObject.counterRightHandBottom, GetType(HandContour)},
        {typeObject.counterRightHandCorniceTop, GetType(HandContour)},
        {typeObject.counterRightHandCorniceBottom, GetType(HandContour)},
        {typeObject.counterLeftPostcardTop, GetType(PostcardContour)},
        {typeObject.counterLeftPostcardBottom, GetType(PostcardContour)},
        {typeObject.counterRightPostcardTop, GetType(PostcardContour)},
        {typeObject.counterRightPostcardBottom, GetType(PostcardContour)},
        {typeObject.counterPileTop, GetType(PileContour)},
        {typeObject.counterPileBottom, GetType(PileContour)},
        {typeObject.counterRigelTop, GetType(RigelContour)},
        {typeObject.counterRigelBottom, GetType(RigelContour)},
        {typeObject.counterSubFermentersTop, GetType(SubFermenterContour)},
        {typeObject.counterSubFermentersBottom, GetType(SubFermenterContour)},
        {typeObject.counterRackTop, GetType(RackContour)},
        {typeObject.counterRackBottom, GetType(RackContour)},
        {typeObject.counterIcecutterTop, GetType(IcecutterCounter)},
        {typeObject.counterIcecutterBottom, GetType(IcecutterCounter)},
        {typeObject.counterPreparationTop, GetType(PreparationContour)},
        {typeObject.counterPreparationBottom, GetType(PreparationContour)},
        {typeObject.counterTopBeam, GetType(CounterBeam)},
        {typeObject.counterBottomBeam, GetType(CounterBeam)},
        {typeObject.counterSiteMonolitBeamsTop, GetType(SiteMonolitBeams)},
        {typeObject.counterSiteMonolitBeamsBottom, GetType(SiteMonolitBeams)},
        {typeObject.hatchSiteMonolitBeams, GetType(HatchSiteMonolitBeams)},
        {typeObject.modelBeam, GetType(ModelBeam)},
        {typeObject.modelNozzle, GetType(NozzleModel)},
        {typeObject.modelRigel, GetType(RigelModel)},
        {typeObject.modelCabinetWall, GetType(CabinetWallModel)},
        {typeObject.modelCabinetWallPlate, GetType(CabinetWallModel)},
        {typeObject.modelGrillage, GetType(GrillageModel)},
        {typeObject.modelPile, GetType(PileModel)},
        {typeObject.modelRack, GetType(RackModel)},
        {typeObject.modelLeftHand, GetType(HandModel)},
        {typeObject.modelLeftHandCornice, GetType(HandModel)},
        {typeObject.modelRightHand, GetType(HandModel)},
        {typeObject.modelRightHandCornice, GetType(HandModel)},
        {typeObject.modelLeftPostcard, GetType(PostcardModel)},
        {typeObject.modelRightPostcard, GetType(PostcardModel)},
        {typeObject.modelSubFermenters, GetType(SubFermenterModel)},
        {typeObject.modelPreparation, GetType(PreparationModel)},
        {typeObject.modelIcecutter, GetType(IcecutterModel)},
        {typeObject.modelSiteMonolitBeams, GetType(ModelSiteMonolitBeams)},
        {typeObject.boundaresBridge, GetType(BoundaryStructure)},
        {typeObject.OtherElement, GetType(OtherElement)}
    }

    Public Shared Function CreateObject(typeEnum As typeObject, jsonData As String) As Object
        ' Проверка входных данных
        If String.IsNullOrEmpty(jsonData) Then
            System.Diagnostics.Debug.WriteLine("JSON данные пустые")
            Return Nothing
        End If

        ' Получаем тип
        Dim targetType As Type = GetTargetType(typeEnum)
        If targetType Is Nothing Then
            System.Diagnostics.Debug.WriteLine($"Тип {typeEnum} не найден в словаре")
            Return Nothing
        End If

        Try
            Return JsonConvert.DeserializeObject(jsonData, targetType)
        Catch ex As JsonReaderException
            System.Diagnostics.Debug.WriteLine($"Ошибка парсинга JSON для {typeEnum}: {ex.Message}")
            Return Nothing
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Ошибка десериализации для {typeEnum}: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Создает объект и сразу отображает в PropertyGrid
    ''' </summary>
    Public Shared Sub LoadToPropertyGrid(typeEnum As typeObject, jsonData As String, propertyGrid As PropertyGrid)
        If propertyGrid Is Nothing Then
            System.Diagnostics.Debug.WriteLine("PropertyGrid не может быть Nothing")
            Return
        End If

        Dim obj As Object = CreateObject(typeEnum, jsonData)

        If obj IsNot Nothing Then
            propertyGrid.SelectedObject = obj
        Else
            propertyGrid.SelectedObject = Nothing
            System.Diagnostics.Debug.WriteLine($"Не удалось загрузить объект для типа {typeEnum}")
        End If
    End Sub

    ''' <summary>
    ''' Проверяет, поддерживается ли тип
    ''' </summary>
    Public Shared Function IsTypeSupported(typeEnum As typeObject) As Boolean
        Return _typeMappings.ContainsKey(typeEnum)
    End Function

    ''' <summary>
    ''' Получает тип .NET для enum
    ''' </summary>
    Public Shared Function GetTargetType(typeEnum As typeObject) As Type
        Dim targetType As Type = Nothing
        If _typeMappings.TryGetValue(typeEnum, targetType) Then
            Return targetType
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Получает все поддерживаемые типы
    ''' </summary>
    Public Shared Function GetSupportedTypes() As List(Of typeObject)
        Return _typeMappings.Keys.ToList()
    End Function

    ''' <summary>
    ''' Получает все типы с их описаниями
    ''' </summary>
    Public Shared Function GetSupportedTypesWithDescriptions() As Dictionary(Of typeObject, String)
        Dim result As New Dictionary(Of typeObject, String)()

        For Each key As typeObject In _typeMappings.Keys
            result.Add(key, GetEnumDescription(key))
        Next

        Return result
    End Function

    ''' <summary>
    ''' Получает описание enum из атрибута Description
    ''' </summary>
    Public Shared Function GetEnumDescription(typeEnum As typeObject) As String
        Dim enumType As Type = GetType(typeObject)
        Dim memberInfo As MemberInfo = enumType.GetMember(typeEnum.ToString()).FirstOrDefault()

        If memberInfo IsNot Nothing Then
            Dim descAttr As DescriptionAttribute = memberInfo.GetCustomAttribute(Of DescriptionAttribute)()
            If descAttr IsNot Nothing Then
                Return descAttr.Description
            End If
        End If

        Return typeEnum.ToString()
    End Function

    ''' <summary>
    ''' Добавляет новое соответствие (для расширения)
    ''' </summary>
    Public Shared Sub AddMapping(typeEnum As typeObject, targetType As Type)
        If targetType Is Nothing Then
            System.Diagnostics.Debug.WriteLine("targetType не может быть Nothing")
            Return
        End If

        If Not _typeMappings.ContainsKey(typeEnum) Then
            _typeMappings.Add(typeEnum, targetType)
        Else
            System.Diagnostics.Debug.WriteLine($"Тип {typeEnum} уже существует в словаре")
        End If
    End Sub

    ''' <summary>
    ''' Обновляет существующее соответствие
    ''' </summary>
    Public Shared Sub UpdateMapping(typeEnum As typeObject, targetType As Type)
        If targetType Is Nothing Then
            System.Diagnostics.Debug.WriteLine("targetType не может быть Nothing")
            Return
        End If

        If _typeMappings.ContainsKey(typeEnum) Then
            _typeMappings(typeEnum) = targetType
        Else
            _typeMappings.Add(typeEnum, targetType)
        End If
    End Sub

    ''' <summary>
    ''' Удаляет соответствие
    ''' </summary>
    Public Shared Sub RemoveMapping(typeEnum As typeObject)
        If _typeMappings.ContainsKey(typeEnum) Then
            _typeMappings.Remove(typeEnum)
        End If
    End Sub

    ''' <summary>
    ''' Очищает словарь (использовать с осторожностью!)
    ''' </summary>
    Public Shared Sub ClearAllMappings()
        _typeMappings.Clear()
    End Sub
End Class

<Serializable>
Public Class PointsCollections
    Private _centerPoint As List(Of Vector3D)
    Private _axisPoints As List(Of Vector3D)
    Private _listPointModel As Dictionary(Of Integer, PointStructure)
    Private _listPointModelSecond As Dictionary(Of Integer, PointStructure)

    Public Sub New()
        _centerPoint = New List(Of Vector3D) From {Nothing, Nothing}
        _axisPoints = New List(Of Vector3D) From {Nothing, Nothing}
        _listPointModel = New Dictionary(Of Integer, PointStructure)()
        _listPointModelSecond = New Dictionary(Of Integer, PointStructure)()
    End Sub
    ' ось элемента
    <Browsable(False)>
    Public Property CenterTopPoint() As Vector3D
        Get
            Return _centerPoint(0)
        End Get
        Set(value As Vector3D)
            _centerPoint(0) = value
        End Set
    End Property

    <Browsable(False)>
    Public Property CenterBottomPoint() As Vector3D
        Get
            Return _centerPoint(1)
        End Get
        Set(value As Vector3D)
            _centerPoint(1) = value
        End Set
    End Property
    ' ось элемента
    <Browsable(False)>
    Public Property StartAxisPoint() As Vector3D
        Get
            Return _axisPoints(0)
        End Get
        Set(value As Vector3D)
            _axisPoints(0) = value
        End Set
    End Property

    <Browsable(False)>
    Public Property EndAxisPoint() As Vector3D
        Get
            Return _axisPoints(1)
        End Get
        Set(value As Vector3D)
            _axisPoints(1) = value
        End Set
    End Property
    'координаты основного контура элемента
    <Browsable(False)>
    Public Property ListPointModel As Dictionary(Of Integer, PointStructure)
        Get
            Return _listPointModel
        End Get
        Set(value As Dictionary(Of Integer, PointStructure))
            _listPointModel = value
        End Set
    End Property

    'координаты вспомогательного контура элемента
    <Browsable(False)>
    Public Property ListPointSecondModel As Dictionary(Of Integer, PointStructure)
        Get
            Return _listPointModelSecond
        End Get
        Set(value As Dictionary(Of Integer, PointStructure))
            _listPointModelSecond = value
        End Set
    End Property
End Class