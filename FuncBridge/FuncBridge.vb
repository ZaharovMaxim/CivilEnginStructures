Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Net.Security
Imports System.Security.AccessControl
Imports System.Security.RightsManagement
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.Alg.Prf
Imports Topomatic.Alg.Road.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Sfc
Imports Topomatic.Sites.Core
Imports Topomatic.Smt
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Точки
Public Class pointBridge
    Public Property x As Double
    Public Property y As Double
    Public Property z As Double
    Public Property dx As Double
    Public Property dy As Double
    Public Property dz As Double
    Public Property h2 As Double
    Public Property code As String
    Public Sub New(xElement As Double, yElement As Double, zElement As Double, dxElement As Double, dyElement As Double, dzElement As Double, h2Element As Double, codeElement As String)
        Me.x = xElement
        Me.y = yElement
        Me.z = zElement
        Me.dx = dxElement
        Me.dy = dyElement
        Me.dz = dzElement
        Me.h2 = h2Element
        Me.code = codeElement
    End Sub
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'сооружение
Public Class Bridge
    Public name As String 'имя сооружения/////
    Public countProlet As Integer 'число пролетов
    Public countLeftRows As Integer 'число рядов влево от оси
    Public countRightRows As Integer 'число рядов вправо от оси
    Public centreAxisBeam As Boolean 'наличие осевого ряда
    Public typeBridge As Integer '0-путепровод с фиксированными балками, 1-путепровод с балками индивидуального проектирования,2-путепровод с расчетным максимальным зазором
    Public dimLeftStructure As Double 'ширина сооружения влево
    Public dimRightStructure As Double 'ширина сооружения вправо
    Public startPosition As Double 'пикет начальной раскладки мостового сооружения (не зависит от горизонтального смещения)
    Public offsetHPosition As Double 'горизонтальное смещение (вдоль оси)
    Public offsetHTPosition As Double 'горизонтальное смещение (поперек оси)
    Public offsetVPosition As Double 'вертикальное смещение
    Public lenght As Double 'длина сооружения
    Public startStation As Double 'фактический пикет начала раскладки
    Public endStation As Double 'фактический пикет конца раскладки
    Public nameSurface As String 'имя поверхности проектной
    Public nameEgSurface As String 'имя поверхности земли
    Public nameAlignment As String 'имя трассы
    Public IDPolyline As String 'идентификатор полилинии (пока не определено)
    Public Sub New()
        Me.name = ""
        countProlet = 1
        countRightRows = 0
        countLeftRows = 0
        centreAxisBeam = True
        typeBridge = 0
        dimLeftStructure = 0
        dimRightStructure = 0
        startPosition = 0
        offsetHPosition = 0
        offsetVPosition = 0
        offsetHTPosition = 0
        lenght = 0
        startStation = 0
        endStation = 0
        nameSurface = ""
        nameAlignment = ""
        nameEgSurface = ""
        IDPolyline = ""
    End Sub

End Class
'ось опоры (горизонтальная)
Public Class AxisHorizontalPillars
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ось опоры горизонтальная
    'расчетные параметры
    Public number As Integer 'номер опоры
    Public defining As Boolean 'опора явлется определяющей
    Public leftClearence As Double 'левый зазор
    Public rightClearence As Double 'правый зазор
    Public monolitSite As Double 'участок омоноличивания 
    Public station As Double 'пикет пересечения с трассой
    Public twoRigels As Boolean 'число ригелей (1 или 2)
    Public singleLevelBeams As Boolean 'ригели в один уровень
    Public deltaLenghtRigel As Double 'расстояние между смежными ригелями
    Public lenghtLeftAxis As Double 'длина влево от оси
    Public lenghtRightAxis As Double 'длина вправо от оси

    Public Sub New()
        number = 0
        defining = False
        station = 0
        leftClearence = 0
        rightClearence = 0
        monolitSite = 0
        lenghtLeftAxis = 0
        lenghtRightAxis = 0
        twoRigels = False
        singleLevelBeams = True
        deltaLenghtRigel = 0
    End Sub
End Class
'ось опирания балки
Public Class AxisBeamsPillar
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ось опирания балок
    Public numberPillar As Integer 'номер опоры
    Public numberColl As Integer 'номер пролета
    Public station As Double 'пикет пересечения с трассой
    Public leftLenght As Double 'длина влево от оси
    Public rightLenght As Double 'длина вправо от оси
    Public leftElevation As Double 'отметка влево
    Public rightElevation As Double 'отметка вправо
    Public Sub New()
        numberPillar = 0
        numberColl = 0
        station = 0
        leftLenght = 0
        rightLenght = 0
        leftElevation = 0
        rightElevation = 0
    End Sub
End Class
'балка
Public Class Beams
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'параметры балки ГОСТ
    Public model As String 'модель
    Public nameAlbum As String 'альбом
    Public fullLenght As Double 'полная длина болки
    Public widthTop As Double 'ширина по верху
    Public widthBottom As Double 'ширина по низу
    Public heightTopPlate As Double 'толщина плиты
    Public height As Double 'полная высота балки
    Public a As Double 'расстояние от начала балки до точки опирания в начале балки
    Public b As Double 'расстояние от конца балки до точки опирания в конце балки
    Public verticalOffsetRibZone As Double '
    Public horizontalOffsetRibZone As Double
    Public deltaRib As Double
    Public radiusTop As Double
    Public radiusBottom As Double
    Public gWidth As Double
    Public startLenghtMonolith As Double 'участок омоличивания в начале балки
    Public endLenghtMonolith As Double 'участок омоличивания в начале балки
    Public mass As Double 'масса
    Public concrete As String 'марка бетона
    Public volume As String 'объем балки
    Public steel As String 'масса стали
    Public modelTLS As String 'имя модели TLC
    Public nameSurface As String 'проектная поверхность
    Public nameAlignment As String 'проектная трасса
    Public IDPolyline As String 'id оси раскладки балок
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'расчетные параметры
    Public numberProlet As Integer 'номер пролета
    Public numberRow As Integer 'номер ряда
    Public otherNumber As String 'дополнительный номер (для балок вставленных вручную, с доп.рядом) пока не используется
    Public clearence As Double 'зазор с предыдущей балкой
    Public dEarth As Double 'толщина покрытия над балкой
    Public offsetStartSurface 'толщина покрытия в точке А
    Public offsetEndSurface 'толщина покрытия в точе B
    Public axisStartOffset As Double 'смещение от оси в начале балки
    Public axisEndOffset As Double 'смещение от оси в конце балки
    Public startPK As Double 'пикет начала балки
    Public endPK As Double 'пикет конца балки
    Public axisCentre As Boolean 'балка является осевой
    Public Sub New()
        model = ""
        nameAlbum = ""
        fullLenght = 0
        widthTop = 0
        widthBottom = 0
        heightTopPlate = 0.15
        height = 0
        a = 0.3
        b = 0.3
        verticalOffsetRibZone = 0
        horizontalOffsetRibZone = 0
        deltaRib = 0
        radiusTop = 0
        radiusBottom = 0
        gWidth = 0
        startLenghtMonolith = 0
        endLenghtMonolith = 0
        mass = 0
        concrete = ""
        volume = 0
        steel = ""
        modelTLS = ""
        nameSurface = ""
        nameAlignment = ""
        IDPolyline = ""
        numberProlet = 0
        numberRow = 0
        otherNumber = ""
        clearence = 0
        dEarth = 0.1
        axisStartOffset = 0.1
        axisEndOffset = 0.1
        axisStartOffset = 0
        axisEndOffset = 0
        startPK = 0
        endPK = 0
        axisCentre = False
    End Sub
End Class
'участок омоноличивания балки
Public Class MonotitingLongSite
    Public numberColl As Integer 'номер пролета
    Public numberLeftBeam As Integer 'номер левой балки
    Public numberRightBeam As Integer 'номер правой балки
    Public thickness As Double 'толщина
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberColl = 0
        numberLeftBeam = 0
        numberRightBeam = 0
        thickness = 0.15
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
End Class
'участок омоноличивания опоры
Public Class MonotitingPillarSite
    Public numberPillar As Integer 'номер опоры
    Public perimeter As Double 'периметр
    Public thickness As Double 'толщина
    Public area As Double 'площадь
    Public volume As Double 'объем
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        perimeter = 0
        thickness = 0.15
        area = 0
        volume = 0
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
End Class
'вспомогательные элементы балки
Public Class ElementsBridge
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'параметры балки ГОСТ
    Public name As String
    Public numberRows As Integer
    Public numberProlet As Integer
End Class
Public Class ElementsBeams
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'параметры балки ГОСТ
    Public name As String
    Public numberProlet As Integer
    Public numberRows As Integer
    Public Sub New()
        name = ""
        numberProlet = 0
        numberRows = 0
    End Sub
End Class


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'опоры мостовых сооружений
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'насадка
Public Class Nozzle
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер насадки
    Public lenght As Double 'длина насадки 
    Public width As Double 'ширина насадки
    Public firstHeight As Double 'высота насадки в начале в зоне расположения балок
    Public secondHeight As Double 'высота насадки в месте расположения шкафной стенки
    Public outletLeftBeam As Double 'выпуск насадки влево за левую балку
    Public outletRightBeam As Double 'выпуск насадки вправо за правую балку
    Public lenghtLeftConsole As Double 'длина левой консоли
    Public lenghtRightConsole As Double 'длина правой консоли
    Public heightLeftConsole As Double 'высота торца ригеля слева
    Public heightRightConsole As Double 'высота торца ригеля слева
    Public minElevationBeams As Double 'наименьшее расстояние от балки до насадки
    Public widthCabinetWall As Double 'ширина горизонтальной площадки под шкафную стенку
    Public topElevation As Double 'отметка верха насадки
    Public bottomElevation As Double 'отметка низа насадки
    Public leftDirection As Double 'левый угол скоса короткой стороны
    Public rightDirection As Double 'правый угол скоса короткой стороны
    Public offsetLeftRack As Double 'смещение начало левой стойки
    Public offsetRightRack As Double 'смещение начало правой стойки
    Public countRacks As Integer 'количество стоек
    Public stepRacks As Double 'шаг расстановки стоек
    Public pileRowsFieldDiagram As String 'схема расстановки свай ряда
    Public pileColumnFieldDiagram As String 'схема расстановки свай столбца
    Public listPointAxis As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        lenght = 0
        width = 0
        firstHeight = 0
        secondHeight = 0
        outletLeftBeam = 0
        outletRightBeam = 0
        lenghtLeftConsole = 0
        lenghtRightConsole = 0
        heightLeftConsole = 0
        heightRightConsole = 0
        minElevationBeams = 0
        widthCabinetWall = 0
        topElevation = 0
        bottomElevation = 0
        leftDirection = 0
        rightDirection = 0
        offsetLeftRack = 0
        offsetRightRack = 0
        countRacks = 0
        stepRacks = 0
        pileRowsFieldDiagram = ""
        pileColumnFieldDiagram = ""
        listPointAxis = New Dictionary(Of Integer, pointBridge)
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "lenght" Then
            lenght = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "firstHeight" Then
            firstHeight = value
        ElseIf name Like "secondHeight" Then
            secondHeight = value
        ElseIf name Like "outletLeftBeam" Then
            outletLeftBeam = value
        ElseIf name Like "outletRightBeam" Then
            outletRightBeam = value
        ElseIf name Like "lenghtLeftConsole" Then
            lenghtLeftConsole = value
        ElseIf name Like "lenghtRightConsole" Then
            lenghtRightConsole = value
        ElseIf name Like "heightLeftConsole" Then
            heightLeftConsole = value
        ElseIf name Like "heightRightConsole" Then
            heightRightConsole = value
        ElseIf name Like "minElevationBeams" Then
            minElevationBeams = value
        ElseIf name Like "widthCabinetWall" Then
            widthCabinetWall = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "leftDirection" Then
            leftDirection = value
        ElseIf name Like "rightDirection" Then
            rightDirection = value
        ElseIf name Like "offsetLeftRack" Then
            offsetLeftRack = value
        ElseIf name Like "offsetRightRack" Then
            offsetRightRack = value
        ElseIf name Like "countRacks" Then
            countRacks = value
        ElseIf name Like "stepRacks" Then
            stepRacks = value
        ElseIf name Like "pileRowsFieldDiagram" Then
            pileRowsFieldDiagram = value
        ElseIf name Like "pileColumnFieldDiagram" Then
            pileColumnFieldDiagram = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'ригель
Public Class Rigel
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер ригеля
    Public lenght As Double 'полная длина ригеля 
    Public width As Double 'ширина ригеля
    Public height As Double 'высота ригеля
    Public outletLeftBeam As Double 'минимальный выпуск ригеля влево за габарит сооружения
    Public outletRightBeam As Double 'выпуск выпуск вправо ригеля за габарит сооружения
    Public heightLeftConsole As Double 'высота торца ригеля слева
    Public heightRightConsole As Double 'высота торца ригеля слева
    Public lenghtLeftConsole As Double 'длина левой консоли
    Public lenghtRightConsole As Double 'длина правой консоли
    Public minElevationBeams As Double 'расстояние от балки до ригеля
    Public leftDirection As Double 'дирекционное направление левой грани ригеля
    Public rightDirection As Double 'дирекционное направление правой грани ригеля
    Public topElevation As Double 'отметка верха
    Public bottomElevation As Double 'отметка низа
    Public axisOffset As Double 'смещение центра ригеля относительно оси
    Public offsetLeftRack As Double
    Public offsetRightRack As Double
    Public countRack As Integer 'количество стоек
    Public stepRack As Double 'шаг расстановки стоек
    Public heightDrain As Double 'высота слива
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        lenght = 0
        width = 0
        height = 0
        minElevationBeams = 0
        leftDirection = 0
        rightDirection = 0
        topElevation = 0
        bottomElevation = 0
        axisOffset = 0
        heightLeftConsole = 0
        heightRightConsole = 0
        lenghtLeftConsole = 0
        lenghtRightConsole = 0
        outletLeftBeam = 0
        outletRightBeam = 0
        offsetLeftRack = 0
        offsetRightRack = 0
        countRack = 0
        stepRack = 0
        heightDrain = 0
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "lenght" Then
            lenght = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "height" Then
            height = value
        ElseIf name Like "minElevationBeams" Then
            minElevationBeams = value
        ElseIf name Like "leftDirection" Then
            leftDirection = value
        ElseIf name Like "rightDirection" Then
            rightDirection = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "axisOffset" Then
            axisOffset = value
        ElseIf name Like "heightLeftConsole" Then
            heightLeftConsole = value
        ElseIf name Like "heightRightConsole" Then
            heightRightConsole = value
        ElseIf name Like "lenghtLeftConsole" Then
            lenghtLeftConsole = value
        ElseIf name Like "lenghtRightConsole" Then
            lenghtRightConsole = value
        ElseIf name Like "outletLeftBeam" Then
            outletLeftBeam = value
        ElseIf name Like "outletRightBeam" Then
            outletRightBeam = value
        ElseIf name Like "offsetLeftRack" Then
            offsetLeftRack = value
        ElseIf name Like "offsetRightRack" Then
            offsetRightRack = value
        ElseIf name Like "countRack" Then
            countRack = value
        ElseIf name Like "stepRack" Then
            stepRack = value
        ElseIf name Like "heightDrain" Then
            heightDrain = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'подферменник
Public Class SubFermenter
    Public numberPillar As Integer 'номер опоры
    Public numberSubPillar As Integer 'номер подопоры
    Public numberRow As Integer 'номер ряда
    Public numberColl As Integer 'номер столбца
    Public lenght As Double 'длина 
    Public width As Double 'ширина 
    Public dHeight As Double 'высота
    Public topElevation As Double 'отметка верха 
    Public rotation As Double
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public centerPoint As Vector3D 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        numberSubPillar = 0
        numberRow = 0
        numberColl = 0
        lenght = 0
        width = 0
        dHeight = 0
        topElevation = 0
        rotation = 0
        centerPoint = New Vector3D(0, 0, 0)
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "numberSubPillar" Then
            numberSubPillar = value
        ElseIf name Like "numberRow" Then
            numberRow = value
        ElseIf name Like "numberColl" Then
            numberColl = value
        ElseIf name Like "length" Then
            lenght = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "dHeight" Then
            dHeight = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "rotation" Then
            rotation = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'шкафная стенка
Public Class CabinetWall
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер 
    Public width As Double 'ширина 
    Public lenght As Double 'полная длина
    Public height As Double 'высота по умолчанию
    Public elevationOffsetProjectSurface As Double 'заглубление
    Public elevationTopPlate As Double 'отметка верха плиты
    Public outletLeftNozzle As Double 'выпуск насадки влево за габарит насадки
    Public outletRightNozzle As Double 'выпуск насадки вправо за габарит насадки
    Public heightTopPl As Double
    Public fullLenghtPl As Double
    Public lenghtPl As Double
    Public widthPl As Double
    Public fixedHeight As Boolean 'зафиксировать высоту
    Public listPointModelWall As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public listPointModelPlate As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        width = 0
        lenght = 0
        height = 0
        elevationTopPlate = 0
        elevationOffsetProjectSurface = 0
        outletLeftNozzle = 0
        outletRightNozzle = 0
        heightTopPl = 0
        fullLenghtPl = 0
        lenghtPl = 0
        widthPl = 0
        fixedHeight = False
        listPointModelWall = New Dictionary(Of Integer, pointBridge)
        listPointModelPlate = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub

    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "elevationOffsetProjectSurface" Then
            elevationOffsetProjectSurface = value
        ElseIf name Like "outletLeftNozzle" Then
            outletLeftNozzle = value
        ElseIf name Like "outletRightNozzle" Then
            outletRightNozzle = value
        ElseIf name Like "fixedHeight" Then
            fixedHeight = CBool(value)
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'откосные крылья
Public Class Hand
    Public numberPillar As Integer 'номер опоры
    Public numberSubPillar As Integer 'номер подопоры
    Public number As Integer 'номер откосного крыла
    Public width As Double 'Толщина крыла без учета карниза
    Public lengthTop As Double  'длина крыла по верху
    Public lenghtBottom As Double = 0 'длина крыла по низу
    Public heightTop As Double = 0 'высота крыла от верха насадки
    Public heightBottom As Double 'выпуск крыла вниз по торцу насадки
    Public heightTopFace As Double 'высота крыла по фасаду
    Public heightBottomFace As Double 'Длина по торцу крыла сзади (вертикальная линия по дальнему концу)
    Public heightCornice As Double 'высота карниза
    Public widthCornice As Double 'высота карниза
    Public listPointModelHand As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public listPointModelCornice As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public type As String 'тип (Left, Right)
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        numberSubPillar = 0
        number = 0
        width = 0
        lengthTop = 0
        lenghtBottom = 0
        heightTop = 0
        heightBottom = 0
        heightTopFace = 0
        heightBottomFace = 0
        heightCornice = 0
        widthCornice = 0
        listPointModelHand = New Dictionary(Of Integer, pointBridge)
        listPointModelCornice = New Dictionary(Of Integer, pointBridge)
        type = ""
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            numberSubPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "lengthTop" Then
            lengthTop = value
        ElseIf name Like "lenghtBottom" Then
            lenghtBottom = value
        ElseIf name Like "heightTop" Then
            heightTop = value
        ElseIf name Like "heightBottom" Then
            heightBottom = value
        ElseIf name Like "heightTopFace" Then
            heightTopFace = value
        ElseIf name Like "heightBottomFace" Then
            heightBottomFace = value
        ElseIf name Like "heightCornice" Then
            heightCornice = value
        ElseIf name Like "widthCornice" Then
            widthCornice = value
        ElseIf name Like "type" Then
            type = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'обратные открылки
Public Class Postcard
    Public numberPillar As Integer 'номер опоры
    Public numberSubPillar As Integer 'номер подопоры
    Public number As Integer 'номер откосного крыла
    Public length As Double = 0 'Длина открылки
    Public width As Double 'Толщина открылка
    Public lenghtCabinetWall As Double 'длина горизонтальной части у шкафной стенки
    Public heightCabinetWall As Double 'высота открылка у шкафной стенки
    Public heightEndNozzle As Double 'высота открылка у торца насадки
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public type As String 'тип (Left, Right)
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        numberSubPillar = 0
        number = 0
        length = 0
        width = 0
        lenghtCabinetWall = 0
        heightCabinetWall = 0
        heightEndNozzle = 0
        type = ""
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            numberSubPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "length" Then
            length = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "lenghtCabinetWall" Then
            lenghtCabinetWall = value
        ElseIf name Like "heightCabinetWall " Then
            heightCabinetWall = value
        ElseIf name Like "heightEndNozzle" Then
            heightEndNozzle = value
        ElseIf name Like "type" Then
            type = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'Стойка
Public Class Rack
    Public numberPillar As Integer 'номер опоры
    Public numberSubPillars As Integer 'номер подопоры
    Public number As Integer 'номер стойки
    Public diameter As Double 'диаметр или длина стойки вдоль насадки или ригеля
    Public height As Double 'высота стойки
    Public topSeal As Double 'величина заделки в насадку или ригель
    Public bottomSeal As Double 'величина заделки в ростверк
    Public widthRackTop As Double 'ширина стойки в уровне насадки или ригеля
    Public widthRackBottom As Double 'ширина стойки в уровне ростверка
    Public offsetAxisTopX As Double  'смещение верха вдоль элемента относительно рассчетного значения ригеля или насадки
    Public offsetAxisTopY As Double  'смещение поперек элемента верха относительно левого края ригеля или насадки
    'Public offsetAxisBottomX As Double  'смещение верха вдоль элемента относительно рассчетного значения ригеля или насадки
    'Public offsetAxisBottomY As Double  'смещение поперек элемента верха относительно левого края ригеля или насадки
    'Public offsetEdgeTop As Double 'отступ стойки от края ригеля или насадки (слева)
    'Public offsetEdgeBottom As Double 'отступ стойки от края ростверка (справа)
    Public topElevation As Double 'отметка верха
    Public bottomElevation As Double 'отметка низа
    Public edgesParallel As Boolean 'Грани стоек  паралельны граням насадки
    Public fixedHeight As Boolean 'высота стойки задается пользователем
    Public rotation As Double 'угол поворота
    Public pointCenter As pointBridge
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public type As String
    Public model As String 'имя модели

    Public Sub New()
        numberPillar = 0
        numberSubPillars = 0
        number = 0
        diameter = 0
        widthRackTop = 0
        widthRackBottom = 0
        'offsetEdgeTop = 0
        'offsetEdgeBottom = 0
        offsetAxisTopX = 0
        offsetAxisTopY = 0
        'offsetAxisBottomX = 0
        'offsetAxisBottomY = 0
        height = 0
        topSeal = 0
        bottomSeal = 0
        topElevation = 0
        bottomElevation = 0
        rotation = 0
        edgesParallel = True
        fixedHeight = False
        pointCenter = New pointBridge(0, 0, 0, 0, 0, 0, 0, "")
        listPointModel = New Dictionary(Of Integer, pointBridge)
        type = ""
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "numberSubPillars" Then
            numberSubPillars = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "diameter" Then
            diameter = value
        ElseIf name Like "widthRackTop" Then
            widthRackTop = value
        ElseIf name Like "widthRackBottom" Then
            widthRackBottom = value
        ElseIf name Like "offsetEdgeTop" Then
            'offsetEdgeTop = value
        ElseIf name Like "offsetEdgeBottom" Then
            'offsetEdgeBottom = value
        ElseIf name Like "offsetAxisTopX" Then
            offsetAxisTopX = value
        ElseIf name Like "offsetAxisTopY" Then
            offsetAxisTopY = value
        ElseIf name Like "offsetAxisBottomX" Then
            'offsetAxisBottomX = value
        ElseIf name Like "offsetAxisBottomY" Then
            'offsetAxisBottomY = value
        ElseIf name Like "height" Then
            height = value
        ElseIf name Like "topSeal" Then
            topSeal = value
        ElseIf name Like "bottomSeal" Then
            bottomSeal = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "edgesParallel" Then
            edgesParallel = value
        ElseIf name Like "fixedHeight" Then
            fixedHeight = value
        ElseIf name Like "type" Then
            type = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'массивное тело
Public Class Icecutter
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер тела
    Public elevation As Double 'отметка верха
    Public downElevation As Double 'отметка низа
    Public axisOffset As Double = 0 'смещение относительно оси трассы
    Public depthGrillage 'смещение от ростверка (ввех положительное число)
    Public leftOffsetRack As Double = 0 '
    Public rightOffstRack As Double = 0
    Public station As Double 'пикет
    Public offsetStation As Double 'смещение
    Public rotation As Double 'угол поворота
    Public fullLenght As Double 'полная длина
    Public width As Double 'ширина
    Public height As Double 'высота
    Public radiusFace As Double
    Public radiusTop As Double
    Public heightDrain As Double
    Public fixedPosition As Boolean 'фиксированное положение
    Public type As String
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        elevation = 0
        downElevation = 0
        station = 0
        axisOffset = 0
        offsetStation = 0
        leftOffsetRack = 0
        rightOffstRack = 0
        rotation = 0
        fullLenght = 0
        width = 0
        height = 0
        radiusFace = 0
        radiusTop = 0
        heightDrain = 0
        fixedPosition = False
        type = ""
        model = ""
    End Sub
End Class
'ростверк
Public Class Grillage
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер ростверка
    Public lenght As Double 'длина ростверка
    Public width As Double 'ширина
    Public height As Double 'высота
    Public offsetCenter As Double 'смещение ростверка относительно оси опоры
    Public depthFoundation As Double 'глубина заложения относительно земли
    Public offsetLeftRack As Double 'отступ ростверка слева от крайней опоры
    Public offsetRightRack As Double 'отступ ростверка справа от крайней опоры
    Public offsetEgeRack As Double 'отступ края ростверка от стойки (вперд назад)
    Public pileRowsFieldDiagram As String 'схема расстановки свай для ряда
    Public pileColumnFieldDiagram As String 'схема расстановки свай для столбца
    Public topElevation As Double 'отметка верха ростверка
    Public bottomElevation As Double 'отметка низа ростверка
    Public edgesParallelNozzle As Boolean 'линии торцов ростверка параллельны торцам насадки
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        lenght = 0
        width = 0
        height = 0
        offsetCenter = 0
        depthFoundation = 0
        pileRowsFieldDiagram = ""
        pileColumnFieldDiagram = ""
        topElevation = 0
        bottomElevation = 0
        offsetLeftRack = 0
        offsetRightRack = 0
        offsetEgeRack = 0
        edgesParallelNozzle = False
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "lenght" Then
            lenght = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "height" Then
            height = value
        ElseIf name Like "offsetCenter" Then
            offsetCenter = value
        ElseIf name Like "depthFoundation" Then
            depthFoundation = value
        ElseIf name Like "offsetLeftRack" Then
            offsetLeftRack = value
        ElseIf name Like "offsetRightRack" Then
            offsetRightRack = value
        ElseIf name Like "pileRowsFieldDiagram" Then
            pileRowsFieldDiagram = value
        ElseIf name Like "pileColumnFieldDiagram" Then
            pileColumnFieldDiagram = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "offsetEgeRack" Then
            offsetEgeRack = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'подготовка
Public Class Preparation
    Public numberPillar As Integer 'номер опоры
    Public number As Integer 'номер подготовки
    Public topElevation As Double 'отметка верха
    Public bottomElevation As Double 'отметка низа
    Public lenght As Double 'длина влево
    Public width As Double 'ширина
    Public topLenght As Double
    Public bottomLenght As Double
    Public topWidth As Double
    Public bottomWidth As Double
    Public height As Double 'высота
    Public offsetLenghtTop As Double 'смещение сверху ростверка или насадки
    Public offsetLenghtBottom As Double 'смещение снизу ростверка или насадки
    Public material As String 'материал
    Public listPointModel As Dictionary(Of Integer, pointBridge) 'номер вершины, СПИСОК=x,y,z.h1,h2,code
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        number = 0
        topElevation = 0
        bottomElevation = 0
        lenght = 0
        width = 0
        topLenght = 0
        bottomLenght = 0
        topWidth = 0
        bottomWidth = 0
        height = 0
        offsetLenghtTop = 0
        offsetLenghtBottom = 0
        material = ""
        listPointModel = New Dictionary(Of Integer, pointBridge)
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "number" Then
            number = value
        ElseIf name Like "lenght" Then
            lenght = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "height" Then
            height = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "offsetLenghtTop" Then
            offsetLenghtTop = value
        ElseIf name Like "offsetLenghtBottom" Then
            offsetLenghtBottom = value
        ElseIf name Like "material" Then
            material = value
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class
'свая
Public Class Pile
    Public numberPillar As Integer 'номер опоры
    Public numberSubPillars As Integer 'номер подопоры
    Public numberRow As Integer 'номер ряда
    Public numberColl As Integer 'номер столбца
    Public diameter As Double 'длина стороны сваи
    Public width As Double 'ширина стороны сваи
    Public height As Double 'высота сваи
    Public topSeal As Double 'высота заделки в ростверк или насадку
    Public offsetX As Double 'смещение сваи вдоль ростверка
    Public offsetY As Double 'смещение сваи поперек ростверка
    Public offsetRow As Double 'смещение низа сваи в ряде относительно центра
    Public offsetColumn As Double 'смещение низа сваи в столбце 
    Public angleX As Double
    Public angleY As Double
    Public rotation As Double 'угол поворота
    Public topElevation As Double 'отметка верха
    Public bottomElevation As Double 'отметка низа
    Public expand As Boolean = False 'наличие уширения
    Public heightExpand As Double 'высота уширения
    Public widthExpand As Double 'диаметр уширения
    Public heightDownExpand As Double 'высота от низа сваи до уширения
    Public degExpand As Double 'угол уширения
    Public pileInRack As Boolean 'стойка в свае
    Public type As String 'призматическая\буровая
    Public pointCenter As pointBridge
    Public model As String 'имя модели
    Public Sub New()
        numberPillar = 0
        numberSubPillars = 0
        numberRow = 0
        numberColl = 0
        diameter = 0
        width = 0
        height = 0
        topSeal = 0
        offsetX = 0
        offsetY = 0
        offsetRow = 0
        offsetColumn = 0
        angleX = 90
        angleY = 90
        rotation = 0
        topElevation = 0
        bottomElevation = 0
        expand = False
        heightExpand = 0.1
        widthExpand = 1
        heightDownExpand = 0.2
        degExpand = 10
        type = ""
        pileInRack = False
        pointCenter = New pointBridge(0, 0, 0, 0, 0, 0, 0, "")
        model = ""
    End Sub
    Public Function SetValue(ByVal name As String, ByRef value As String)
        If name Like "numberPillar" Then
            numberPillar = value
        ElseIf name Like "numberSubPillars" Then
            numberSubPillars = value
        ElseIf name Like "numberRow" Then
            numberRow = value
        ElseIf name Like "numberColl" Then
            numberColl = value
        ElseIf name Like "diameter" Then
            diameter = value
        ElseIf name Like "width" Then
            width = value
        ElseIf name Like "height" Then
            height = value
        ElseIf name Like "topSeal" Then
            topSeal = value
        ElseIf name Like "offsetX" Then
            offsetX = value
        ElseIf name Like "offsetY" Then
            offsetY = value
        ElseIf name Like "offsetRow" Then
            offsetRow = value
        ElseIf name Like "offsetColumn" Then
            offsetColumn = value
        ElseIf name Like "angleX" Then
            angleX = value
        ElseIf name Like "angleY" Then
            angleY = value
        ElseIf name Like "rotation" Then
            rotation = value
        ElseIf name Like "topElevation" Then
            topElevation = value
        ElseIf name Like "bottomElevation" Then
            bottomElevation = value
        ElseIf name Like "expand" Then
            expand = CBool(value)
        ElseIf name Like "heightExpand" Then
            heightExpand = value
        ElseIf name Like "widthExpand" Then
            widthExpand = value
        ElseIf name Like "heightDownExpand" Then
            heightDownExpand = value
        ElseIf name Like "degExpand" Then
            degExpand = value
        ElseIf name Like "type" Then
            type = value
        ElseIf name Like "pileInRack" Then
            pileInRack = CBool(value)
        ElseIf name Like "model" Then
            model = value
        End If
    End Function
End Class


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство горизонтальной оси опоры
Public Class PropertyAxisPillarsPanel
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private defining As Boolean
    <Browsable(True)>
    <Description("Определяющая опора (от данной оси опоры происходит раскладка балок)")>
    <Category("Свойства")>
    <DisplayName("Определяющая опора")>
    Public Property b_defining As Boolean
        Get
            Return defining
        End Get
        Set(ByVal value As Boolean)
            defining = value
        End Set
    End Property
    Private leftClearence As Double
    <Browsable(True)>
    <Description("Зазор слева в пролете, м")>
    <Category("Свойства")>
    <DisplayName("Левый зазор")>
    Public Property b_leftClearence As Double
        Get
            Return leftClearence
        End Get
        Set(ByVal value As Double)
            leftClearence = value
        End Set
    End Property
    Private rightClearence As Double
    <Browsable(True)>
    <Description("Зазор справа в пролете, м")>
    <Category("Свойства")>
    <DisplayName("Правый зазор")>
    Public Property b_rightClearence As Double
        Get
            Return rightClearence
        End Get
        Set(ByVal value As Double)
            rightClearence = value
        End Set
    End Property
    Private monolitSite As Double
    <Browsable(True)>
    <Description("Размер участка омоноличивания балки, м")>
    <Category("Свойства")>
    <DisplayName("Омоноличивание")>
    Public Property b_monolitSite As Double
        Get
            Return monolitSite
        End Get
        Set(ByVal value As Double)
            monolitSite = value
        End Set
    End Property
    Private stationPillars As String
    <Browsable(True)>
    <Description("Пикет пересечения с осью трассы, ПК+")>
    <Category("Пикет")>
    <DisplayName("Пикет")>
    <[ReadOnly](True)>
    Public Property b_stationPillars As String
        Get
            Return stationPillars
        End Get
        Set(ByVal value As String)
            stationPillars = value
        End Set
    End Property
    Private lenghtLeftAlignAxis As Double
    <Browsable(True)>
    <Description("Длина влево от оси трассы, м")>
    <Category("Свойства")>
    <DisplayName("Длина влево")>
    <[ReadOnly](True)>
    Public Property b_lenghtLeftAlignAxis As String
        Get
            Return lenghtLeftAlignAxis
        End Get
        Set(ByVal value As String)
            lenghtLeftAlignAxis = value
        End Set
    End Property
    Private lenghtRightAlignAxis As Double
    <Browsable(True)>
    <Description("Длина вправо от оси трассы, м")>
    <Category("Свойства")>
    <DisplayName("Длина вправо")>
    <[ReadOnly](True)>
    Public Property b_lenghtRightAlignAxis As String
        Get
            Return lenghtRightAlignAxis
        End Get
        Set(ByVal value As String)
            lenghtRightAlignAxis = value
        End Set
    End Property
    Private twoRigels As Boolean
    <Browsable(True)>
    <Description("Наличие двух отдельных ригелей в опоре")>
    <Category("Свойства")>
    <DisplayName("Два ригеля")>
    Public Property b_twoRigels As Boolean
        Get
            Return twoRigels
        End Get
        Set(ByVal value As Boolean)
            twoRigels = value
        End Set
    End Property
    Private singleLevelBeams As Boolean
    <Browsable(True)>
    <Description("Вставка ригелей в один уровень")>
    <Category("Свойства")>
    <DisplayName("Один уровень")>
    Public Property b_singleLevelBeams As Boolean
        Get
            Return singleLevelBeams
        End Get
        Set(ByVal value As Boolean)
            singleLevelBeams = value
        End Set
    End Property
    Private deltaLenghtRigel As Double
    <Browsable(True)>
    <Description("Расстояние между смежными ригелямиа")>
    <Category("Свойства")>
    <DisplayName("Зазор между ригелями")>
    <[ReadOnly](True)>
    Public Property b_deltaLenghtRigel As String
        Get
            Return deltaLenghtRigel
        End Get
        Set(ByVal value As String)
            deltaLenghtRigel = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство осей опирания
Public Class PropertyAxisBeamsPanel
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private numberColumn As Integer
    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Свойства")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property b_numberColumn As Integer
        Get
            Return numberColumn
        End Get
        Set(ByVal value As Integer)
            numberColumn = value
        End Set
    End Property
    Private leftLenght As Double
    <Browsable(True)>
    <Description("Длина оси между влево от оси трассы автомобильной дороги, м")>
    <Category("Свойства")>
    <DisplayName("Длина оси влево")>
    <[ReadOnly](True)>
    Public Property b_leftLenght As Double
        Get
            Return leftLenght
        End Get
        Set(ByVal value As Double)
            leftLenght = value
        End Set
    End Property
    Private rightLenght As Double
    <Browsable(True)>
    <Description("Длина оси между вправо от оси трассы автомобильной дороги, м")>
    <Category("Свойства")>
    <DisplayName("Длина оси вправо")>
    <[ReadOnly](True)>
    Public Property b_rightLenght As Double
        Get
            Return rightLenght
        End Get
        Set(ByVal value As Double)
            rightLenght = value
        End Set
    End Property
    Private startElevation As Double
    <Browsable(True)>
    <Description("Отметка левого конца оси опирания балока, м")>
    <Category("Свойства")>
    <DisplayName("Отметка лево")>
    <[ReadOnly](True)>
    Public Property b_startElevation As Double
        Get
            Return startElevation
        End Get
        Set(ByVal value As Double)
            startElevation = value
        End Set
    End Property
    Private endElevation As Double
    <Browsable(True)>
    <Description("Отметка правого конца оси опирания балока, м")>
    <Category("Свойства")>
    <DisplayName("Отметка право")>
    <[ReadOnly](True)>
    Public Property b_endElevation As Double
        Get
            Return endElevation
        End Get
        Set(ByVal value As Double)
            endElevation = value
        End Set
    End Property
    Private stationAxisBeams As String
    <Browsable(True)>
    <Description("Пикет пересечения с осью трассы автомобильной дороги, ПК+")>
    <Category("Пикет")>
    <DisplayName("Пикет пересечения")>
    <[ReadOnly](True)>
    Public Property b_stationAxisBeams As String
        Get
            Return stationAxisBeams
        End Get
        Set(ByVal value As String)
            stationAxisBeams = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойства балки
Public Class PropertyBeamsPanel
    Private numColumns As Integer
    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Местоположение")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property b_numColumns As Integer
        Get
            Return numColumns
        End Get
        Set(ByVal value As Integer)
            numColumns = value
        End Set
    End Property
    Private numRows As String
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Местоположение")>
    <DisplayName("Номер балки")>
    <[ReadOnly](True)>
    Public Property b_numRows As String
        Get
            Return numRows
        End Get
        Set(ByVal value As String)
            numRows = value
        End Set
    End Property
    Private startPK As String
    <Browsable(True)>
    <Description("Пикетажное положения начальной точки опирания, ПК+")>
    <Category("Местоположение")>
    <DisplayName("Начальный пикет")>
    <[ReadOnly](True)>
    Public Property b_startPK As String
        Get
            Return startPK
        End Get
        Set(ByVal value As String)
            startPK = value
        End Set
    End Property
    Private offsetStart As Double
    <Browsable(True)>
    <Description("Смещение балки от главной оси в начальной точки опирания, м")>
    <Category("Местоположение")>
    <DisplayName("Начальное смещение")>
    <[ReadOnly](True)>
    Public Property b_offsetStart As Double
        Get
            Return offsetStart
        End Get
        Set(ByVal value As Double)
            offsetStart = value
        End Set
    End Property
    Private endPK As String
    <Browsable(True)>
    <Description("Пикетажное положения конечной точки опирания ПК+")>
    <Category("Местоположение")>
    <DisplayName("Конечный пикет")>
    <[ReadOnly](True)>
    Public Property b_endPK As String
        Get
            Return endPK
        End Get
        Set(ByVal value As String)
            endPK = value
        End Set
    End Property
    Private offsetEnd As Double
    <Browsable(True)>
    <Description("Смещение балки от главной оси в конечной точки опирания, м")>
    <Category("Местоположение")>
    <DisplayName("Конечное смещение")>
    <[ReadOnly](True)>
    Public Property b_offsetEnd As Double
        Get
            Return offsetEnd
        End Get
        Set(ByVal value As Double)
            offsetEnd = value
        End Set
    End Property
    Private fullLenght As Double
    <Browsable(True)>
    <Description("Полная длина балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Длина балки")>
    <[ReadOnly](True)>
    Public Property b_fullLenght As Double
        Get
            Return fullLenght
        End Get
        Set(ByVal value As Double)
            fullLenght = value
        End Set
    End Property
    Private widthBottom As Double
    <Browsable(True)>
    <Description("Ширина ребра низа балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Ширина по низу")>
    <[ReadOnly](True)>
    Public Property b_widthBottom As Double
        Get
            Return widthBottom
        End Get
        Set(ByVal value As Double)
            widthBottom = value
        End Set
    End Property
    Private heightTopPlate As Double
    <Browsable(True)>
    <Description("Толщина верхней опорной плиты балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Толщина плиты")>
    <[ReadOnly](True)>
    Public Property b_heightTopPlate As Double
        Get
            Return heightTopPlate
        End Get
        Set(ByVal value As Double)
            heightTopPlate = value
        End Set
    End Property
    Private widthTop As Double
    <Browsable(True)>
    <Description("Ширина плиты балки по верху, м")>
    <Category("Параметры балки")>
    <DisplayName("Ширина по верху")>
    <[ReadOnly](True)>
    Public Property b_widthTop As Double
        Get
            Return widthTop
        End Get
        Set(ByVal value As Double)
            widthTop = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Высота")>
    <[ReadOnly](True)>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private clearence As Double
    <Browsable(True)>
    <Description("Зазор между предыдущей балкой, м")>
    <Category("Параметры балки")>
    <DisplayName("Зазор")>
    <[ReadOnly](True)>
    Public Property b_clearence As Double
        Get
            Return clearence
        End Get
        Set(ByVal value As Double)
            clearence = value
        End Set
    End Property
    Private a As Double
    <Browsable(True)>
    <Description("Расстояние от начала балки до точки опирания, м")>
    <Category("Параметры балки")>
    <DisplayName("Начальная точка опирания")>
    <[ReadOnly](True)>
    Public Property b_a As Double
        Get
            Return a
        End Get
        Set(ByVal value As Double)
            a = value
        End Set
    End Property
    Private b As Double
    <Browsable(True)>
    <Description("Расстояние от конца балки до точки опирания, м")>
    <Category("Параметры балки")>
    <DisplayName("Конечная точка опирания")>
    <[ReadOnly](True)>
    Public Property b_b As Double
        Get
            Return b
        End Get
        Set(ByVal value As Double)
            b = value
        End Set
    End Property
    Private dEarth As Double
    <Browsable(True)>
    <Description("Минимальная толщина покрытия, м")>
    <Category("Параметры балки")>
    <DisplayName("Минимальная толщина покрытия")>
    <[ReadOnly](True)>
    Public Property b_dEarth As Double
        Get
            Return dEarth
        End Get
        Set(ByVal value As Double)
            dEarth = value
        End Set
    End Property
    Private monolitStartSites As Double
    <Browsable(True)>
    <Description("Участок омоличивания в начале балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Омоноличивание в начале")>
    <[ReadOnly](True)>
    Public Property b_monolitStartSites As Double
        Get
            Return monolitStartSites
        End Get
        Set(ByVal value As Double)
            monolitStartSites = value
        End Set
    End Property
    Private monolitEndSites As Double
    <Browsable(True)>
    <Description("Участок омоличивания в конце балки, м")>
    <Category("Параметры балки")>
    <DisplayName("Омоноличивание в конце")>
    <[ReadOnly](True)>
    Public Property b_monolitEndSites As Double
        Get
            Return monolitEndSites
        End Get
        Set(ByVal value As Double)
            monolitEndSites = value
        End Set
    End Property
    Private nameAlbum As String
    <Browsable(True)>
    <Description("Наименование альбома балок")>
    <Category("Описание")>
    <DisplayName("Альбом")>
    <[ReadOnly](True)>
    Public Property b_nameAlbum As String
        Get
            Return nameAlbum
        End Get
        Set(ByVal value As String)
            nameAlbum = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Марка балки")>
    <Category("Описание")>
    <DisplayName("Марка балки")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property

    'Private myList5 As List(Of String) = New List(Of String)()
    '<Browsable(True)>
    '<Description("Наименование альбома балок")>
    '<Category("Примечание")>
    '<DisplayName("Альбом2")>
    '<Editor("System.Windows.Forms.Design.StringCollectionEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", "System.Drawing.Design.UITypeEditor,System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")>
    '<TypeConverter(GetType(BeamsAlbumsTypeConverter))>
    'Public Property b_MyList5 As List(Of String)
    '    Get
    '        Return myList5
    '    End Get
    '    Set(ByVal value As List(Of String))
    '        myList5 = value
    '    End Set

    'End Property

    'Private nameAlbum As String
    '<DisplayName("Наименование альбома балок")>
    '<Description("Альбом")>
    '<Category("Описание")>
    'Public Property b_nameAlbum As String
    '    Get
    '        Return nameAlbum
    '    End Get
    '    Set(ByVal value As String)
    '        nameAlbum = value
    '    End Set
    'End Property

    'Public Class BeamsAlbumsTypeConverter
    '    Inherits StringConverter
    '    Dim arrayAlbum As String() = {""}
    '    Public Sub New(ByRef arrayAlbumBeams As String())
    '        arrayAlbum = arrayAlbumBeams
    '    End Sub
    '    Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
    '        Return True
    '    End Function
    '    Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
    '        Return True
    '    End Function
    '    Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
    '        Return New StandardValuesCollection(arrayAlbum)
    '    End Function
    'End Class
    'Private myList As String() = {}
    '<Browsable(True)>
    '<Description("Альбом")>
    '<Category("Примечание")>
    '<DisplayName("Альбом")>
    'Public Property b_MyList As String()
    '    Get
    '        Return myList
    '    End Get
    '    Set(ByVal value As String())
    '        myList = value
    '    End Set
    'End Property
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство ряда балок
Public Class PropertyRowsPanel
    Private numBeams As Integer
    <Browsable(True)>
    <Description("Количество балок в ряде, шт")>
    <Category("Параметры сооружения")>
    <DisplayName("Количество балок")>
    <[ReadOnly](True)>
    Public Property b_numBeams As Integer
        Get
            Return numBeams
        End Get
        Set(ByVal value As Integer)
            numBeams = value
        End Set
    End Property
    Private dEarth As Double
    <Browsable(True)>
    <Description("Минимальная толщина покрытия, м")>
    <Category("Параметры сооружения")>
    <DisplayName("Минимальная толщина покрытия")>
    Public Property b_dEarth As Double
        Get
            Return dEarth
        End Get
        Set(ByVal value As Double)
            dEarth = value
        End Set
    End Property
    Private offsetAxis As Double
    <Browsable(True)>
    <Description("Смещение ряда балок относительно главной оси сооружения, м")>
    <Category("Параметры сооружения")>
    <DisplayName("Смещение от оси")>
    Public Property b_offsetAxis As Double
        Get
            Return offsetAxis
        End Get
        Set(ByVal value As Double)
            offsetAxis = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойства пролета
Public Class PropertyColumnPanel
    Private numberColl As Integer
    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Параметры сооружения")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property b_numberColl As Integer
        Get
            Return numberColl
        End Get
        Set(ByVal value As Integer)
            numberColl = value
        End Set
    End Property
    Private countBeams As Integer
    <Browsable(True)>
    <Description("Количество балок в пролете, шт")>
    <Category("Параметры сооружения")>
    <DisplayName("Количество балок")>
    <[ReadOnly](True)>
    Public Property b_countBeams As Integer
        Get
            Return countBeams
        End Get
        Set(ByVal value As Integer)
            countBeams = value
        End Set
    End Property



End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов верха балки
Public Class PropertyTopElementsBeamPanel
    Private offsetElements As Double
    <Browsable(True)>
    <Description("Смещение элемента относительно оси балки, шт")>
    <Category("Параметры сооружения")>
    <DisplayName("Смещение от оси")>
    <[ReadOnly](True)>
    Public Property b_offsetElements As Double
        Get
            Return offsetElements
        End Get
        Set(ByVal value As Double)
            offsetElements = value
        End Set
    End Property
    Private ElevationA As Double
    <Browsable(True)>
    <Description("Отметка в точке опирания (а), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка (а)")>
    <[ReadOnly](True)>
    Public Property b_ElevationA As Double
        Get
            Return ElevationA
        End Get
        Set(ByVal value As Double)
            ElevationA = value
        End Set
    End Property
    Private ElevationB As Double
    <Browsable(True)>
    <Description("Отметка в точке опирания (b), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка (b)")>
    <[ReadOnly](True)>
    Public Property b_ElevationB As Double
        Get
            Return ElevationB
        End Get
        Set(ByVal value As Double)
            ElevationB = value
        End Set
    End Property
    Private ElevationEarthA As Double
    <Browsable(True)>
    <Description("Отметка поверхности в точке опирания (а), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка поверхности (а)")>
    <[ReadOnly](True)>
    Public Property b_ElevationEarthA As Double
        Get
            Return ElevationEarthA
        End Get
        Set(ByVal value As Double)
            ElevationEarthA = value
        End Set
    End Property
    Private ElevationEarthB As Double
    <Browsable(True)>
    <Description("Отметка поверхности в точке опирания (b), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка поверхности (b)")>
    <[ReadOnly](True)>
    Public Property b_ElevationEarthB As Double
        Get
            Return ElevationEarthB
        End Get
        Set(ByVal value As Double)
            ElevationEarthB = value
        End Set
    End Property
    Private dEarthA As Double
    <Browsable(True)>
    <Description("Толщина покрытия в точке опирания (а), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Толщина покрытия в точке (a)")>
    <[ReadOnly](True)>
    Public Property b_dEarth As Double
        Get
            Return dEarthA
        End Get
        Set(ByVal value As Double)
            dEarthA = value
        End Set
    End Property
    Private dEarthВ As Double
    <Browsable(True)>
    <Description("Толщина покрытия в точке опирания (b), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Толщина покрытия в точке (b)")>
    <[ReadOnly](True)>
    Public Property b_dEarthВ As Double
        Get
            Return dEarthВ
        End Get
        Set(ByVal value As Double)
            dEarthВ = value
        End Set
    End Property
    Private clearence As Double
    <Browsable(True)>
    <Description("Зазор с предыдущей балкой, м")>
    <Category("Параметры сооружения")>
    <DisplayName("Зазор")>
    <[ReadOnly](True)>
    Public Property b_clearence As Double
        Get
            Return clearence
        End Get
        Set(ByVal value As Double)
            clearence = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов низа балки
Public Class PropertyDownElementsBeamPanel
    Private offsetElements As Double
    <Browsable(True)>
    <Description("Смещение элемента относительно оси балки, шт")>
    <Category("Параметры сооружения")>
    <DisplayName("Смещение от оси")>
    <[ReadOnly](True)>
    Public Property b_offsetElements As Double
        Get
            Return offsetElements
        End Get
        Set(ByVal value As Double)
            offsetElements = value
        End Set
    End Property
    Private ElevationA As Double
    <Browsable(True)>
    <Description("Отметка в точке опирания (а), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка (а)")>
    <[ReadOnly](True)>
    Public Property b_ElevationA As Double
        Get
            Return ElevationA
        End Get
        Set(ByVal value As Double)
            ElevationA = value
        End Set
    End Property
    Private ElevationB As Double
    <Browsable(True)>
    <Description("Отметка в точке опирания (b), м")>
    <Category("Параметры сооружения")>
    <DisplayName("Отметка (b)")>
    <[ReadOnly](True)>
    Public Property b_ElevationB As Double
        Get
            Return ElevationB
        End Get
        Set(ByVal value As Double)
            ElevationB = value
        End Set
    End Property
    Private clearence As Double
    <Browsable(True)>
    <Description("Зазор с предыдущей балкой, м")>
    <Category("Параметры сооружения")>
    <DisplayName("Зазор")>
    <[ReadOnly](True)>
    Public Property b_clearence As Double
        Get
            Return clearence
        End Get
        Set(ByVal value As Double)
            clearence = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов НАСАДКА
Public Class PropertyNozzle
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private lenght As Double
    <Browsable(True)>
    <Description("Длина насадки")>
    <Category("Параметры")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lenght As Double
        Get
            Return lenght
        End Get
        Set(ByVal value As Double)
            lenght = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина насадки, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private widthCabinetWall As Double
    <Browsable(True)>
    <Description("Ширина горизонтальной площадки для размещения шнафной стенки, м")>
    <Category("Параметры")>
    <DisplayName("Ширина площадки шкафной стенки")>
    Public Property b_widthCabinetWall As Double
        Get
            Return widthCabinetWall
        End Get
        Set(ByVal value As Double)
            widthCabinetWall = value
        End Set
    End Property
    Private firstHeight As Double
    <Browsable(True)>
    <Description("Высота насадки по фасаду, м")>
    <Category("Параметры")>
    <DisplayName("Высота по фасаду")>
    Public Property b_firstHeight As Double
        Get
            Return firstHeight
        End Get
        Set(ByVal value As Double)
            firstHeight = value
        End Set
    End Property
    Private secondHeight As Double
    <Browsable(True)>
    <Description("Высота насадки в зоне шкафной стенки, м")>
    <Category("Параметры")>
    <DisplayName("Высота у шкафной стенки")>
    Public Property b_secondHeight As Double
        Get
            Return secondHeight
        End Get
        Set(ByVal value As Double)
            secondHeight = value
        End Set
    End Property
    Private minElevationBeams As Double
    <Browsable(True)>
    <Description("Зазор по вертикали, между насадкой и балкой которая имеет наименьшую отметку, м")>
    <Category("Свойства")>
    <DisplayName("Просвет с балкой")>
    <[ReadOnly](True)>
    Public Property b_minElevationBeams As Double
        Get
            Return minElevationBeams
        End Get
        Set(ByVal value As Double)
            minElevationBeams = value
        End Set
    End Property
    Private outletLeftBeam As Double
    <Browsable(True)>
    <Description("Выпуск насадки за габарит крайней левой балки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск влево")>
    Public Property b_outletLeftBeam As Double
        Get
            Return outletLeftBeam
        End Get
        Set(ByVal value As Double)
            outletLeftBeam = value
        End Set
    End Property
    Private outletRightBeam As Double
    <Browsable(True)>
    <Description("Выпуск насадки за габарит крайней правой балки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск вправо")>
    Public Property b_outletRightBeam As Double
        Get
            Return outletRightBeam
        End Get
        Set(ByVal value As Double)
            outletRightBeam = value
        End Set
    End Property
    Private lenghtLeftConsole As Double
    <Browsable(True)>
    <Description("Длина левой консоли насадки, м")>
    <Category("Параметры")>
    <DisplayName("Длина левой консоли")>
    Public Property b_lenghtLeftConsole As Double
        Get
            Return lenghtLeftConsole
        End Get
        Set(ByVal value As Double)
            lenghtLeftConsole = value
        End Set
    End Property
    Private lenghtRightConsole As Double
    <Browsable(True)>
    <Description("Длина правой консоли насадки, м")>
    <Category("Параметры")>
    <DisplayName("Длина правой консоли")>
    Public Property b_lenghtRightConsole As Double
        Get
            Return lenghtRightConsole
        End Get
        Set(ByVal value As Double)
            lenghtRightConsole = value
        End Set
    End Property
    Private heightLeftConsole As Double
    <Browsable(True)>
    <Description("Высота торца левой консоли насадки, м")>
    <Category("Параметры")>
    <DisplayName("Высота левого торца")>
    Public Property b_heightLeftConsole As Double
        Get
            Return heightLeftConsole
        End Get
        Set(ByVal value As Double)
            heightLeftConsole = value
        End Set
    End Property
    Private heightRightConsole As Double
    <Browsable(True)>
    <Description("Высота торца правой консоли насадки, м")>
    <Category("Параметры")>
    <DisplayName("Высота правого торца")>
    Public Property b_heightRightConsole As Double
        Get
            Return heightRightConsole
        End Get
        Set(ByVal value As Double)
            heightRightConsole = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха насадки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа насадки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private offsetLeftRack As Double
    <Browsable(True)>
    <Description("Расстояние от края стойки до края насадки в влево, м")>
    <Category("Параметры")>
    <DisplayName("Отступ стойки влево")>
    <[ReadOnly](True)>
    Public Property b_offsetLeftRack As Double
        Get
            Return offsetLeftRack
        End Get
        Set(ByVal value As Double)
            offsetLeftRack = value
        End Set
    End Property
    Private offsetRightRack As Double
    <Browsable(True)>
    <Description("Расстояние от края стойки до края насадки в вправо, м")>
    <Category("Параметры")>
    <DisplayName("Отступ стойки вправо")>
    <[ReadOnly](True)>
    Public Property b_offsetRightRack As Double
        Get
            Return offsetRightRack
        End Get
        Set(ByVal value As Double)
            offsetRightRack = value
        End Set
    End Property
    Private countRack As Integer
    <Browsable(True)>
    <Description("Количество стоек в насадке, м")>
    <Category("Параметры")>
    <DisplayName("Количество стоек")>
    <[ReadOnly](True)>
    Public Property b_countRack As Integer
        Get
            Return countRack
        End Get
        Set(ByVal value As Integer)
            countRack = value
        End Set
    End Property
    Private stepRacks As Double
    <Browsable(True)>
    <Description("Шаг расстановки стоек в насадке, м")>
    <Category("Параметры")>
    <DisplayName("Шаг стоек")>
    <[ReadOnly](True)>
    Public Property b_stepRacks As Double
        Get
            Return stepRacks
        End Get
        Set(ByVal value As Double)
            stepRacks = value
        End Set
    End Property
    Private pileRowsFieldDiagram As String
    <Browsable(True)>
    <Description("Схема расстановки свай в ряде")>
    <Category("Свойства")>
    <DisplayName("Схема ряда свай")>
    <[ReadOnly](True)>
    Public Property b_pileRowsFieldDiagram As String
        Get
            Return pileRowsFieldDiagram
        End Get
        Set(ByVal value As String)
            pileRowsFieldDiagram = value
        End Set
    End Property
    Private pileColumnsFieldDiagram As String
    <Browsable(True)>
    <Description("Схема расстановки свай в столбце")>
    <Category("Свойства")>
    <DisplayName("Схема столбца свай")>
    <[ReadOnly](True)>
    Public Property b_pileColumnsFieldDiagram As String
        Get
            Return pileColumnsFieldDiagram
        End Get
        Set(ByVal value As String)
            pileColumnsFieldDiagram = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели Насадки")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов РИГЕЛИ
Public Class PropertyRigelS
    Private twoRigels As Boolean
    <Browsable(True)>
    <Description("Два ригеля в опоре")>
    <Category("Свойства")>
    <DisplayName("Два ригеля в опоре")>
    Public Property b_twoRigels As Boolean
        Get
            Return twoRigels
        End Get
        Set(ByVal value As Boolean)
            twoRigels = value
        End Set
    End Property

    Private singleLevelRigels As Boolean
    <Browsable(True)>
    <Description("Ригели в один уровень")>
    <Category("Свойства")>
    <DisplayName("Ригели в один уровень")>
    Public Property b_singleLevelRigels As Boolean
        Get
            Return singleLevelRigels
        End Get
        Set(ByVal value As Boolean)
            singleLevelRigels = value
        End Set
    End Property

    Private deltaLenghtRigel As Double
    <Browsable(True)>
    <Description("Расстояние между смежными ригелями")>
    <Category("Свойства")>
    <DisplayName("Расстояние между ригелями")>
    Public Property b_deltaLenghtRigel As Double
        Get
            Return deltaLenghtRigel
        End Get
        Set(ByVal value As Double)
            deltaLenghtRigel = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента Шкафная стенка
Public Class PropertyCabinetWall
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private lenght As Double
    <Browsable(True)>
    <Description("Длина шкафной стенки, м")>
    <Category("Параметры")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lenght As Double
        Get
            Return lenght
        End Get
        Set(ByVal value As Double)
            lenght = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина шкафной стенки, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private elevationOffsetProjectSurface As Double
    <Browsable(True)>
    <Description("Заглубление над проезжей частью автодороги, м")>
    <Category("Свойства")>
    <DisplayName("Заглубление")>
    Public Property b_elevationOffsetProjectSurface As Double
        Get
            Return elevationOffsetProjectSurface
        End Get
        Set(ByVal value As Double)
            elevationOffsetProjectSurface = value
        End Set
    End Property
    Private widthPl As Double
    <Browsable(True)>
    <Description("Ширина опорной части Зуба упора, м")>
    <Category("Параметры")>
    <DisplayName("Ширина зуба упора")>
    Public Property b_widthPl As Double
        Get
            Return widthPl
        End Get
        Set(ByVal value As Double)
            widthPl = value
        End Set
    End Property
    Private lenghtPl As Double
    <Browsable(True)>
    <Description("Высота зуба упора от верха опорной части, м")>
    <Category("Параметры")>
    <DisplayName("Высота зуба упора")>
    Public Property b_lenghtPl As Double
        Get
            Return lenghtPl
        End Get
        Set(ByVal value As Double)
            lenghtPl = value
        End Set
    End Property
    Private fullLenghtPl As Double
    <Browsable(True)>
    <Description("Высота зуба упора вдоль шкафной стенки, м")>
    <Category("Параметры")>
    <DisplayName("Высота у шкафной стенки")>
    Public Property b_fullLenghtPl As Double
        Get
            Return fullLenghtPl
        End Get
        Set(ByVal value As Double)
            fullLenghtPl = value
        End Set
    End Property
    Private heightTopPl As Double
    <Browsable(True)>
    <Description("Расстояние от насадки до верха опорной части Зуба упора, м")>
    <Category("Параметры")>
    <DisplayName("Высота от насадки до Зуба упора")>
    Public Property b_heightTopPl As Double
        Get
            Return heightTopPl
        End Get
        Set(ByVal value As Double)
            heightTopPl = value
        End Set
    End Property
    Private elevationTopPlate As Double
    <Browsable(True)>
    <Description("Отметка верха опорной части Зуба упора, м")>
    <Category("Параметры")>
    <DisplayName("Отметка Зуба упора")>
    <[ReadOnly](True)>
    Public Property b_elevationTopPlate As Double
        Get
            Return elevationTopPlate
        End Get
        Set(ByVal value As Double)
            elevationTopPlate = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели Насадки")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов РИГЕЛЬ
Public Class PropertyRigel
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private lenght As Double
    <Browsable(True)>
    <Description("Полная длина ригеля, м")>
    <Category("Параметры")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lenght As Double
        Get
            Return lenght
        End Get
        Set(ByVal value As Double)
            lenght = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина ригеля, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота ригеля, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private maxElevationBeams As Double
    <Browsable(True)>
    <Description("Зазор между ригелем и балкой которая имеет наименьшую отметку, м")>
    <Category("Свойства")>
    <DisplayName("Просвет с балкой")>
    <[ReadOnly](True)>
    Public Property b_maxElevationBeams As Double
        Get
            Return maxElevationBeams
        End Get
        Set(ByVal value As Double)
            maxElevationBeams = value
        End Set
    End Property
    Private outletLeftBeam As Double
    <Browsable(True)>
    <Description("Выпуск ригеля влево за габарит крайней балки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск влево")>
    Public Property b_outletLeftBeam As Double
        Get
            Return outletLeftBeam
        End Get
        Set(ByVal value As Double)
            outletLeftBeam = value
        End Set
    End Property
    Private outletRightBeam As Double
    <Browsable(True)>
    <Description("Выпуск ригеля вправо за габарит крайней балки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск вправо")>
    Public Property b_outletRightBeam As Double
        Get
            Return outletRightBeam
        End Get
        Set(ByVal value As Double)
            outletRightBeam = value
        End Set
    End Property
    Private lenghtLeftConsole As Double
    <Browsable(True)>
    <Description("Длина скоса влево, м")>
    <Category("Параметры")>
    <DisplayName("Длина скоса влево")>
    Public Property b_lenghtLeftConsole As Double
        Get
            Return lenghtLeftConsole
        End Get
        Set(ByVal value As Double)
            lenghtLeftConsole = value
        End Set
    End Property
    Private lenghtRightConsole As Double
    <Browsable(True)>
    <Description("Длина скоса вправо, м")>
    <Category("Параметры")>
    <DisplayName("Длина скоса вправо")>
    Public Property b_lenghtRightConsole As Double
        Get
            Return lenghtRightConsole
        End Get
        Set(ByVal value As Double)
            lenghtRightConsole = value
        End Set
    End Property
    Private heightLeftConsole As Double
    <Browsable(True)>
    <Description("Высота торца ригеля слева, м")>
    <Category("Параметры")>
    <DisplayName("Высота торца слева")>
    Public Property b_heightLeftConsole As Double
        Get
            Return heightLeftConsole
        End Get
        Set(ByVal value As Double)
            heightLeftConsole = value
        End Set
    End Property
    Private heightRightConsole As Double
    <Browsable(True)>
    <Description("Высота торца ригеля справа, м")>
    <Category("Параметры")>
    <DisplayName("Высота торца справа")>
    Public Property b_heightRightConsole As Double
        Get
            Return heightRightConsole
        End Get
        Set(ByVal value As Double)
            heightRightConsole = value
        End Set
    End Property
    Private axisOffsetRigel As Double
    <Browsable(True)>
    <Description("Смещение центра оси ригеля относительно оси опоры, м")>
    <Category("Свойства")>
    <DisplayName("Смещение от оси")>
    <[ReadOnly](True)>
    Public Property b_axisOffsetRigel As Double
        Get
            Return axisOffsetRigel
        End Get
        Set(ByVal value As Double)
            axisOffsetRigel = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа ригеля, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private offsetLeftRack As Double
    <Browsable(True)>
    <Description("Расстояние от края стойки до скоса ригеля влево, м")>
    <Category("Параметры")>
    <DisplayName("Отступ стойки влево")>
    <[ReadOnly](True)>
    Public Property b_offsetLeftRack As Double
        Get
            Return offsetLeftRack
        End Get
        Set(ByVal value As Double)
            offsetLeftRack = value
        End Set
    End Property
    Private offsetRightRack As Double
    <Browsable(True)>
    <Description("Расстояние от края стойки до скоса ригеля вправо, м")>
    <Category("Параметры")>
    <DisplayName("Отступ стойки вправо")>
    <[ReadOnly](True)>
    Public Property b_offsetRightRack As Double
        Get
            Return offsetRightRack
        End Get
        Set(ByVal value As Double)
            offsetRightRack = value
        End Set
    End Property
    Private countRack As Integer
    <Browsable(True)>
    <Description("Количество стоек в ригеле, м")>
    <Category("Параметры")>
    <DisplayName("Количество стоек")>
    <[ReadOnly](True)>
    Public Property b_countRack As Integer
        Get
            Return countRack
        End Get
        Set(ByVal value As Integer)
            countRack = value
        End Set
    End Property
    Private stepRack As Double
    <Browsable(True)>
    <Description("Шаг расстановки стоек в ригеле, м")>
    <Category("Параметры")>
    <DisplayName("Шаг стоек")>
    <[ReadOnly](True)>
    Public Property b_stepRack As Double
        Get
            Return stepRack
        End Get
        Set(ByVal value As Double)
            stepRack = value
        End Set
    End Property
    Private heightDrain As Boolean
    <Browsable(True)>
    <Description("Высота слива по центру ригеля")>
    <Category("Свойства")>
    <DisplayName("Высота слива")>
    Public Property b_heightDrain As Boolean
        Get
            Return heightDrain
        End Get
        Set(ByVal value As Boolean)
            heightDrain = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели ригеля")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов СТОЙКА
Public Class PropertyRack
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private numberRack As String
    <Browsable(True)>
    <Description("Обозначение стойки")>
    <Category("Свойства")>
    <DisplayName("Номер стойки")>
    <[ReadOnly](True)>
    Public Property b_numberRack As String
        Get
            Return numberRack
        End Get
        Set(ByVal value As String)
            numberRack = value
        End Set
    End Property
    Private diameter As Double
    <Browsable(True)>
    <Description("Диаметр или длина стороны стойки вдоль оси насадки или ригеля, м")>
    <Category("Параметры")>
    <DisplayName("Диаметр/длина")>
    Public Property b_diameter As Double
        Get
            Return diameter
        End Get
        Set(ByVal value As Double)
            diameter = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота стойки, м")>
    <Category("Параметры")>
    <DisplayName("Высота стойки")>
    <[ReadOnly](True)>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private widthRackTop As Double
    <Browsable(True)>
    <Description("Ширина стойки в уровне насадки или ригеля, м")>
    <Category("Параметры")>
    <DisplayName("Ширина по верху")>
    Public Property b_widthRackTop As Double
        Get
            Return widthRackTop
        End Get
        Set(ByVal value As Double)
            widthRackTop = value
        End Set
    End Property
    Private widthRackBottom As Double
    <Browsable(True)>
    <Description("Ширина стойки в уровне ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Ширина по низу")>
    Public Property b_widthRackBottom As Double
        Get
            Return widthRackBottom
        End Get
        Set(ByVal value As Double)
            widthRackBottom = value
        End Set
    End Property
    'Private offsetEdgeTop As Double
    '<Browsable(True)>
    '<Description("Отступ стойки от края ригеля или насадки, м")>
    '<Category("Параметры")>
    '<DisplayName("Отступ по верху")>
    '<[ReadOnly](True)>
    'Public Property b_offsetEdgeTop As Double
    '    Get
    '        Return offsetEdgeTop
    '    End Get
    '    Set(ByVal value As Double)
    '        offsetEdgeTop = value
    '    End Set
    'End Property
    'Private offsetEdgeBottom As Double
    '<Browsable(True)>
    '<Description("Отступ стойки от края ростверка, м")>
    '<Category("Параметры")>
    '<DisplayName("Отступ по низу")>
    '<[ReadOnly](True)>
    'Public Property b_offsetEdgeBottom As Double
    '    Get
    '        Return offsetEdgeBottom
    '    End Get
    '    Set(ByVal value As Double)
    '        offsetEdgeBottom = value
    '    End Set
    'End Property
    Private offsetAxisTopX As Double
    <Browsable(True)>
    <Description("Смещение оси стойки от левого края ригеля или насадки, м")>
    <Category("Параметры")>
    <DisplayName("Смещение верха оси по X")>
    Public Property b_offsetAxisTopX As Double
        Get
            Return offsetAxisTopX
        End Get
        Set(ByVal value As Double)
            offsetAxisTopX = value
        End Set
    End Property
    Private offsetAxisTopY As Double
    <Browsable(True)>
    <Description("Смещение оси стойки от оси ригеля или насадки, м")>
    <Category("Параметры")>
    <DisplayName("Смещение верха оси по Y")>
    Public Property b_offsetAxisTopY As Double
        Get
            Return offsetAxisTopY
        End Get
        Set(ByVal value As Double)
            offsetAxisTopY = value
        End Set
    End Property
    'Private offsetAxisBottomX As Double
    '<Browsable(True)>
    '<Description("Смещение оси стойки от левого края ростверка, м")>
    '<Category("Параметры")>
    '<DisplayName("Смещение верха оси по X")>
    'Public Property b_offsetAxisBottomX As Double
    '    Get
    '        Return offsetAxisBottomX
    '    End Get
    '    Set(ByVal value As Double)
    '        offsetAxisBottomX = value
    '    End Set
    'End Property
    'Private offsetAxisBottomY As Double
    '<Browsable(True)>
    '<Description("Смещение оси стойки от оси ростверка, м")>
    '<Category("Параметры")>
    '<DisplayName("Смещение верха оси по Y")>
    'Public Property b_offsetAxisBottomY As Double
    '    Get
    '        Return offsetAxisBottomY
    '    End Get
    '    Set(ByVal value As Double)
    '        offsetAxisBottomY = value
    '    End Set
    'End Property
    Private topSeal As Double
    <Browsable(True)>
    <Description("Величина заделки стойки в насадку или ригель, м")>
    <Category("Свойства")>
    <DisplayName("Заделка в насадку или ригель")>
    Public Property b_topSeal As Double
        Get
            Return topSeal
        End Get
        Set(ByVal value As Double)
            topSeal = value
        End Set
    End Property
    Private bottomSeal As Double
    <Browsable(True)>
    <Description("Величина заделки стойки в ростверк, м")>
    <Category("Свойства")>
    <DisplayName("Заделка в ростверк")>
    Public Property b_bottomSeal As Double
        Get
            Return bottomSeal
        End Get
        Set(ByVal value As Double)
            bottomSeal = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха стойки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа стойки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели стойки")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента МАССИВНОЕ ТЕЛО
Public Class PropertyIcecutter
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private numberIcecutter As String
    <Browsable(True)>
    <Description("Обозначение массивного тела")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_numberIcecutter As String
        Get
            Return numberIcecutter
        End Get
        Set(ByVal value As String)
            numberIcecutter = value
        End Set
    End Property
    Private elevationUPIcecutter As Double
    <Browsable(True)>
    <Description("Отметка верха массивного тела, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    Public Property b_elevationUPIcecutter As Double
        Get
            Return elevationUPIcecutter
        End Get
        Set(ByVal value As Double)
            elevationUPIcecutter = value
        End Set
    End Property
    Private elevationDownIcecutter As Double
    <Browsable(True)>
    <Description("Отметка низа массивного тела, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    Public Property b_elevationDownIcecutter As Double
        Get
            Return elevationDownIcecutter
        End Get
        Set(ByVal value As Double)
            elevationDownIcecutter = value
        End Set
    End Property
    Private depthGrillageIcecutter As Double
    <Browsable(True)>
    <Description("Смещение от ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Смещение от ростверка")>
    Public Property b_depthGrillageIcecutter As Double
        Get
            Return depthGrillageIcecutter
        End Get
        Set(ByVal value As Double)
            depthGrillageIcecutter = value
        End Set
    End Property
    Private axisOffset As Double
    <Browsable(True)>
    <Description("Сдвиг массивного тела относительно оси трассы автомобильной дороги, м")>
    <Category("Свойства")>
    <DisplayName("Сдвиг от оси")>
    Public Property b_axisOffset As Double
        Get
            Return axisOffset
        End Get
        Set(ByVal value As Double)
            axisOffset = value
        End Set
    End Property
    Private station As String
    <Browsable(True)>
    <Description("Пикетажное положения точки вставки массивного тела, ПК+")>
    <Category("Местоположение")>
    <DisplayName("Пикет")>
    <[ReadOnly](True)>
    Public Property b_station As String
        Get
            Return station
        End Get
        Set(ByVal value As String)
            station = value
        End Set
    End Property
    Private offsetStation As String
    <Browsable(True)>
    <Description("Смещение массивного тела относительно оси автомобильной дороги, м")>
    <Category("Местоположение")>
    <DisplayName("Смещение")>
    <[ReadOnly](True)>
    Public Property b_offsetStation As String
        Get
            Return offsetStation
        End Get
        Set(ByVal value As String)
            offsetStation = value
        End Set
    End Property
    'Private leftOffsetRosk As Double
    '<Browsable(True)>
    '<Description("Выпуск массивного тела влево от крайней стойки, м")>
    '<Category("Параметры")>
    '<DisplayName("Выпуск влево")>
    'Public Property b_leftOffsetRosk As Double
    '    Get
    '        Return leftOffsetRosk
    '    End Get
    '    Set(ByVal value As Double)
    '        leftOffsetRosk = value
    '    End Set
    'End Property
    'Private rightOffsetRosk As Double
    '<Browsable(True)>
    '<Description("Выпуск массивного тела вправо от крайней стойки, м")>
    '<Category("Параметры")>
    '<DisplayName("Выпуск вправо")>
    'Public Property b_rightOffsetRosk As Double
    '    Get
    '        Return rightOffsetRosk
    '    End Get
    '    Set(ByVal value As Double)
    '        rightOffsetRosk = value
    '    End Set
    'End Property
    Private fullLenght As Double
    <Browsable(True)>
    <Description("Полная длина массивного тела, м")>
    <Category("Параметры")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_fullLenght As Double
        Get
            Return fullLenght
        End Get
        Set(ByVal value As Double)
            fullLenght = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота массивного тела, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина массивного тела, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private radiusFace As Double
    <Browsable(True)>
    <Description("Радиус скругления фасадов, м")>
    <Category("Параметры")>
    <DisplayName("Радиус фасада")>
    Public Property b_radiusFace As Double
        Get
            Return radiusFace
        End Get
        Set(ByVal value As Double)
            radiusFace = value
        End Set
    End Property
    Private radiusTop As Double
    <Browsable(True)>
    <Description("Радиус скругления торцов, м")>
    <Category("Параметры")>
    <DisplayName("Радиус торцов")>
    Public Property b_radiusTop As Double
        Get
            Return radiusTop
        End Get
        Set(ByVal value As Double)
            radiusTop = value
        End Set
    End Property
    Private heightDrain As Double
    <Browsable(True)>
    <Description("Высота слива, м")>
    <Category("Параметры")>
    <DisplayName("Высота слива")>
    Public Property b_heightDrain As Double
        Get
            Return heightDrain
        End Get
        Set(ByVal value As Double)
            heightDrain = value
        End Set
    End Property
    Private typeIcecutter As String
    <Browsable(True)>
    <Description("Тип массивного тела")>
    <Category("Свойства")>
    <DisplayName("Тип")>
    <TypeConverter(GetType(PostTypeConverter))>
    Public Property b_typeIcecutter As String
        Get
            Return typeIcecutter
        End Get
        Set(ByVal value As String)
            typeIcecutter = value
        End Set
    End Property
    Private modelIcecutter As String
    <Browsable(True)>
    <Description("Имя модели массивного тела")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    Public Property b_modelIcecutter As String
        Get
            Return modelIcecutter
        End Get
        Set(ByVal value As String)
            modelIcecutter = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента РОСТВЕРК
Public Class PropertyGrillage
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private number As String
    <Browsable(True)>
    <Description("Обозначение ростверка")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_number As String
        Get
            Return number
        End Get
        Set(ByVal value As String)
            number = value
        End Set
    End Property
    Private lenght As Double
    <Browsable(True)>
    <Description("Длина ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lenght As Double
        Get
            Return lenght
        End Get
        Set(ByVal value As Double)
            lenght = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    <[ReadOnly](True)>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private offsetCenter As Double
    <Browsable(True)>
    <Description("Смещение относительно оси опоры, м")>
    <Category("Параметры")>
    <DisplayName("Смещение от оси")>
    Public Property b_offsetCenter As Double
        Get
            Return offsetCenter
        End Get
        Set(ByVal value As Double)
            offsetCenter = value
        End Set
    End Property
    Private depthFoundation As Double
    <Browsable(True)>
    <Description("Глубина заложения ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Заложение")>
    Public Property b_depthFoundation As Double
        Get
            Return depthFoundation
        End Get
        Set(ByVal value As Double)
            depthFoundation = value
        End Set
    End Property
    Private offsetLeftRack As Double
    <Browsable(True)>
    <Description("Выпуск ростверка влево от крайней стойки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск влево")>
    Public Property b_offsetLeftRack As Double
        Get
            Return offsetLeftRack
        End Get
        Set(ByVal value As Double)
            offsetLeftRack = value
        End Set
    End Property
    Private offsetRightRack As Double
    <Browsable(True)>
    <Description("Выпуск ростверка вправо от крайней стойки, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск вправо")>
    Public Property b_offsetRightRack As Double
        Get
            Return offsetRightRack
        End Get
        Set(ByVal value As Double)
            offsetRightRack = value
        End Set
    End Property
    Private offsetEgeRack As Double
    <Browsable(True)>
    <Description("Расстояние от стойки до края ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Расстояние от стойки до ростверка")>
    Public Property b_offsetEgeRack As Double
        Get
            Return offsetEgeRack
        End Get
        Set(ByVal value As Double)
            offsetEgeRack = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа ростверка, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private pileRowsFieldDiagram As String
    <Browsable(True)>
    <Description("Схема расстановки свай в ряде")>
    <Category("Параметры")>
    <DisplayName("Схема ряда")>
    <[ReadOnly](True)>
    Public Property b_pileRowsFieldDiagram As String
        Get
            Return pileRowsFieldDiagram
        End Get
        Set(ByVal value As String)
            pileRowsFieldDiagram = value
        End Set
    End Property
    Private pileColumnFieldDiagram As String
    <Browsable(True)>
    <Description("Схема расстановки свай в столбце")>
    <Category("Параметры")>
    <DisplayName("Схема столбца")>
    <[ReadOnly](True)>
    Public Property b_pileColumnFieldDiagram As String
        Get
            Return pileColumnFieldDiagram
        End Get
        Set(ByVal value As String)
            pileColumnFieldDiagram = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели ростверка")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента ПОДГОТОВКА
Public Class PropertyPreparation
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private number As String
    <Browsable(True)>
    <Description("Обозначение подготовки")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_number As String
        Get
            Return number
        End Get
        Set(ByVal value As String)
            number = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха подготовки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа подготовки, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private lenght As Double
    <Browsable(True)>
    <Description("Длина подготовки, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lenght As Double
        Get
            Return lenght
        End Get
        Set(ByVal value As Double)
            lenght = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина подготовки, м")>
    <Category("Параметры")>
    <DisplayName("Ширина")>
    <[ReadOnly](True)>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private offsetLenghtTop As Double
    <Browsable(True)>
    <Description("Выпуск подготовки за габарит ростверка или насадки влево, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск влево")>
    Public Property b_offsetLenghtTop As Double
        Get
            Return offsetLenghtTop
        End Get
        Set(ByVal value As Double)
            offsetLenghtTop = value
        End Set
    End Property
    Private offsetLenghtBottom As Double
    <Browsable(True)>
    <Description("Выпуск подготовки за габарит ростверка или насадки вправо, м")>
    <Category("Параметры")>
    <DisplayName("Выпуск вправо")>
    Public Property b_offsetLenghtBottom As Double
        Get
            Return offsetLenghtBottom
        End Get
        Set(ByVal value As Double)
            offsetLenghtBottom = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота подготовки, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели ростверка")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов СВАИ
Public Class PropertyPile
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private numberRow As Integer
    <Browsable(True)>
    <Description("Обозначение ряда свай")>
    <Category("Свойства")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property b_numberRow As Integer
        Get
            Return numberRow
        End Get
        Set(ByVal value As Integer)
            numberRow = value
        End Set
    End Property
    Private numberColl As Integer
    <Browsable(True)>
    <Description("Обозначение столбца свай")>
    <Category("Свойства")>
    <DisplayName("Номер столбца")>
    <[ReadOnly](True)>
    Public Property b_numberColl As Integer
        Get
            Return numberColl
        End Get
        Set(ByVal value As Integer)
            numberColl = value
        End Set
    End Property
    Private diameter As Double
    <Browsable(True)>
    <Description("Диаметр (длина стороны) сваи, м")>
    <Category("Параметры")>
    <DisplayName("Диаметр")>
    Public Property b_diameter As Double
        Get
            Return diameter
        End Get
        Set(ByVal value As Double)
            diameter = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Высота сваи, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private height As Double
    <Browsable(True)>
    <Description("Высота сваи, м")>
    <Category("Параметры")>
    <DisplayName("Высота")>
    Public Property b_height As Double
        Get
            Return height
        End Get
        Set(ByVal value As Double)
            height = value
        End Set
    End Property
    Private topSeal As Double
    <Browsable(True)>
    <Description("Высота заделки сваи в ростверк, м")>
    <Category("Параметры")>
    <DisplayName("Высота заделки")>
    Public Property b_topSeal As Double
        Get
            Return topSeal
        End Get
        Set(ByVal value As Double)
            topSeal = value
        End Set
    End Property
    Private ofsetX As Double
    <Browsable(True)>
    <Description("Смещение верха сваи вдоль ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Смещение верха (X)")>
    Public Property b_ofsetX As Double
        Get
            Return ofsetX
        End Get
        Set(ByVal value As Double)
            ofsetX = value
        End Set
    End Property
    Private ofsetY As Double
    <Browsable(True)>
    <Description("Смещение верха сваи поперек ростверка, м")>
    <Category("Параметры")>
    <DisplayName("Смещение верха (Y)")>
    Public Property b_ofsetY As Double
        Get
            Return ofsetY
        End Get
        Set(ByVal value As Double)
            ofsetY = value
        End Set
    End Property
    Private offsetRow As Double
    <Browsable(True)>
    <Description("Смещение низа сваи относительно верха сваи в ряде, м")>
    <Category("Параметры")>
    <DisplayName("Смещение низа сваи (X)")>
    Public Property b_offsetRow As Double
        Get
            Return offsetRow
        End Get
        Set(ByVal value As Double)
            offsetRow = value
        End Set
    End Property
    Private offsetColumn As Double
    <Browsable(True)>
    <Description("Смещение низа сваи относительно верха сваи в столбце, м")>
    <Category("Параметры")>
    <DisplayName("Смещение низа сваи (Y)")>
    Public Property b_offsetColumn As Double
        Get
            Return offsetColumn
        End Get
        Set(ByVal value As Double)
            offsetColumn = value
        End Set
    End Property
    Private topElevation As Double
    <Browsable(True)>
    <Description("Отметка верха Сваи, м")>
    <Category("Свойства")>
    <DisplayName("Отметка верха")>
    <[ReadOnly](True)>
    Public Property b_topElevation As Double
        Get
            Return topElevation
        End Get
        Set(ByVal value As Double)
            topElevation = value
        End Set
    End Property
    Private bottomElevation As Double
    <Browsable(True)>
    <Description("Отметка низа сваи, м")>
    <Category("Свойства")>
    <DisplayName("Отметка низа")>
    <[ReadOnly](True)>
    Public Property b_bottomElevation As Double
        Get
            Return bottomElevation
        End Get
        Set(ByVal value As Double)
            bottomElevation = value
        End Set
    End Property
    Private typePile As String
    <Browsable(True)>
    <Description("Тип сваи")>
    <Category("Свойства")>
    <DisplayName("Тип")>
    <TypeConverter(GetType(PostTypeConverter))>
    <[ReadOnly](True)>
    Public Property b_typePile As String
        Get
            Return typePile
        End Get
        Set(ByVal value As String)
            typePile = value
        End Set
    End Property
    Private modelPile As String
    <Browsable(True)>
    <Description("Имя модели сваи")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    <[ReadOnly](True)>
    Public Property b_modelPile As String
        Get
            Return modelPile
        End Get
        Set(ByVal value As String)
            modelPile = value
        End Set
    End Property
    Private expand As ExpandPile
    <Description("Наличие уширения сваи")>
    <Category("Свойства")>
    <DisplayName("Уширение")>
    Public Property b_Expand As ExpandPile
        Get
            Return expand
        End Get
        Set(ByVal value As ExpandPile)
            expand = value
        End Set
    End Property
    <TypeConverter(GetType(ExpandableObjectConverter))>
    Class ExpandPile
        Public Sub New(ByVal boolExpand As Boolean, ByVal heightExpand As Double, ByVal widthExpand As Double, ByVal heightDownExpand As Double, ByVal degExpand As Double)
            _heightExpand = heightExpand
            _widthExpand = widthExpand
            _heightDownExpand = heightDownExpand
            _degExpand = degExpand
        End Sub
        Private boolExpand As Boolean
        <Browsable(True)>
        <DisplayName("Уширение")>
        <Description("Наличие уширения")>
        Public Property b_boolExpand As Boolean
            Get
                Return boolExpand
            End Get
            Set(ByVal value As Boolean)
                boolExpand = value
            End Set
        End Property
        Private _heightExpand As Double
        <Browsable(True)>
        <DisplayName("Уширение")>
        <Description("Величина уширения сваи, м.")>
        Public Property b_heightExpand As Double
            Get
                Return _heightExpand
            End Get
            Set(ByVal value As Double)
                _heightExpand = value
            End Set
        End Property

        Private _widthExpand As Double
        <Browsable(True)>
        <DisplayName("Диаметр")>
        <Description("Диаметр уширения сваи, м.")>
        Public Property b_widthExpand As Double
            Get
                Return _widthExpand
            End Get
            Set(ByVal value As Double)
                _widthExpand = value
            End Set
        End Property

        Private _heightDownExpand As Double
        <Browsable(True)>
        <DisplayName("Высота уширения")>
        <Description("Высота от низа сваи до уширения")>
        Public Property b_heightDownExpand As Double
            Get
                Return _heightDownExpand
            End Get
            Set(ByVal value As Double)
                _heightDownExpand = value
            End Set
        End Property

        Private _degExpand As Double
        <Browsable(True)>
        <DisplayName("Угол уширения")>
        <Description("Угол уширения")>
        Public Property b_degExpand As Double
            Get
                Return _degExpand
            End Get
            Set(ByVal value As Double)
                _degExpand = value
            End Set
        End Property
        Public Overrides Function ToString() As String
            If boolExpand = True Then
                Return "Да"
            Else
                Return "Нет"
            End If
        End Function
    End Class
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов СВАИ (для ряда)
Public Class PropertyRowPiles
    Private numberRow As Integer
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_numberRow As Integer
        Get
            Return numberRow
        End Get
        Set(ByVal value As Integer)
            numberRow = value
        End Set
    End Property
    Private offsetX As Double
    <Browsable(True)>
    <Description("Смещение ряда свай вдоль ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение X")>
    Public Property b_offsetX As Double
        Get
            Return offsetX
        End Get
        Set(ByVal value As Double)
            offsetX = value
        End Set
    End Property
    Private offsetY As Double
    <Browsable(True)>
    <Description("Смещение ряда свай поперек ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение Y")>
    Public Property b_offsetY As Double
        Get
            Return offsetY
        End Get
        Set(ByVal value As Double)
            offsetY = value
        End Set
    End Property
    Private offsetRow As Double
    <Browsable(True)>
    <Description("Смещение низа ряда свай относительно верха, м")>
    <Category("Свойства")>
    <DisplayName("Смещение низа")>
    Public Property b_offsetRow As Double
        Get
            Return offsetRow
        End Get
        Set(ByVal value As Double)
            offsetRow = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элементов СВАИ (для столбца)
Public Class PropertyColumnPiles
    Private numberColumn As Integer
    <Browsable(True)>
    <Description("Номер столбца")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_numberColumn As Integer
        Get
            Return numberColumn
        End Get
        Set(ByVal value As Integer)
            numberColumn = value
        End Set
    End Property
    Private offsetX As Double
    <Browsable(True)>
    <Description("Смещение ряда свай вдоль ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение X")>
    Public Property b_offsetX As Double
        Get
            Return offsetX
        End Get
        Set(ByVal value As Double)
            offsetX = value
        End Set
    End Property
    Private offsetY As Double
    <Browsable(True)>
    <Description("Смещение ряда свай поперек ростверка или насадки, м")>
    <Category("Свойства")>
    <DisplayName("Смещение Y")>
    Public Property b_offsetY As Double
        Get
            Return offsetY
        End Get
        Set(ByVal value As Double)
            offsetY = value
        End Set
    End Property
    Private offsetColumn As Double
    <Browsable(True)>
    <Description("Смещение низа столбца свай относительно верха, м")>
    <Category("Свойства")>
    <DisplayName("Смещение низа")>
    Public Property b_offsetColumn As Double
        Get
            Return offsetColumn
        End Get
        Set(ByVal value As Double)
            offsetColumn = value
        End Set
    End Property
End Class

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента ОБРАТНЫЙ ОТКРЫЛОК
Public Class PropertyPostcard
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private number As String
    <Browsable(True)>
    <Description("Обозначение откосного крыла")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_number As String
        Get
            Return number
        End Get
        Set(ByVal value As String)
            number = value
        End Set
    End Property
    Private length As Double
    <Browsable(True)>
    <Description("Полная длина Обратного открылка, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_length As Double
        Get
            Return length
        End Get
        Set(ByVal value As Double)
            length = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина обратного открылка, м")>
    <Category("Свойства")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private lenghtCabinetWall As Double
    <Browsable(True)>
    <Description("Длина горизонтальной части у шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Длина горизонтальной части")>
    Public Property b_lenghtCabinetWall As Double
        Get
            Return lenghtCabinetWall
        End Get
        Set(ByVal value As Double)
            lenghtCabinetWall = value
        End Set
    End Property
    Private heightCabinetWall As Double
    <Browsable(True)>
    <Description("Высота обратного открылка у шкафной стенки, м")>
    <Category("Свойства")>
    <DisplayName("Высота у шкафной стенки")>
    Public Property b_heightCabinetWall As Double
        Get
            Return heightCabinetWall
        End Get
        Set(ByVal value As Double)
            heightCabinetWall = value
        End Set
    End Property
    Private heightEndNozzle As String
    <Browsable(True)>
    <Description("Высота обратного открылка у торца насадки, м")>
    <Category("Местоположение")>
    <DisplayName("Превышение по верху")>
    <[ReadOnly](True)>
    Public Property b_heightEndNozzle As String
        Get
            Return heightEndNozzle
        End Get
        Set(ByVal value As String)
            heightEndNozzle = value
        End Set
    End Property
    Private type As String
    <Browsable(True)>
    <Description("Тип")>
    <Category("Свойства")>
    <DisplayName("Тип")>
    Public Property b_type As String
        Get
            Return type
        End Get
        Set(ByVal value As String)
            type = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство элемента ОТКОСНЫЕ КРЫЛЬЯ
Public Class PropertyHand
    Private numberPillars As Integer
    <Browsable(True)>
    <Description("Номер опоры мостового сооружения")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property b_numberPillars As Integer
        Get
            Return numberPillars
        End Get
        Set(ByVal value As Integer)
            numberPillars = value
        End Set
    End Property
    Private number As String
    <Browsable(True)>
    <Description("Обозначение откосного крыла")>
    <Category("Свойства")>
    <DisplayName("Номер")>
    <[ReadOnly](True)>
    Public Property b_number As String
        Get
            Return number
        End Get
        Set(ByVal value As String)
            number = value
        End Set
    End Property
    Private lengthTop As Double
    <Browsable(True)>
    <Description("Полная длина откосного крыла по верху, м")>
    <Category("Свойства")>
    <DisplayName("Длина")>
    <[ReadOnly](True)>
    Public Property b_lengthTop As Double
        Get
            Return lengthTop
        End Get
        Set(ByVal value As Double)
            lengthTop = value
        End Set
    End Property
    Private width As Double
    <Browsable(True)>
    <Description("Ширина откосного крыла, м")>
    <Category("Свойства")>
    <DisplayName("Ширина")>
    Public Property b_width As Double
        Get
            Return width
        End Get
        Set(ByVal value As Double)
            width = value
        End Set
    End Property
    Private lenghtBottom As Double
    <Browsable(True)>
    <Description("Длина горизонтальной части нижней части крыла, м")>
    <Category("Свойства")>
    <DisplayName("Низ крыла")>
    Public Property b_lenghtBottom As Double
        Get
            Return lenghtBottom
        End Get
        Set(ByVal value As Double)
            lenghtBottom = value
        End Set
    End Property
    Private heightTop As Double
    <Browsable(True)>
    <Description("Высота откосного крыла от верха насадки (без учета карниза), м")>
    <Category("Свойства")>
    <DisplayName("Высота от верха")>
    Public Property b_heightTop As Double
        Get
            Return heightTop
        End Get
        Set(ByVal value As Double)
            heightTop = value
        End Set
    End Property
    Private heightBottom As String
    <Browsable(True)>
    <Description("Выпуск крыла вниз по грани насадки, м")>
    <Category("Местоположение")>
    <DisplayName("Выпуск вниз")>
    <[ReadOnly](True)>
    Public Property b_heightBottom As String
        Get
            Return heightBottom
        End Get
        Set(ByVal value As String)
            heightBottom = value
        End Set
    End Property
    Private heightTopFace As String
    <Browsable(True)>
    <Description("Высота крыла по фасаду, м")>
    <Category("Местоположение")>
    <DisplayName("Высота по фасаду")>
    <[ReadOnly](True)>
    Public Property b_heightTopFace As String
        Get
            Return heightTopFace
        End Get
        Set(ByVal value As String)
            heightTopFace = value
        End Set
    End Property
    Private heightBottomFace As Double
    <Browsable(True)>
    <Description("Длина по торцу крыла задней части крыла (вертикальная линия по дальнему концу), м")>
    <Category("Параметры")>
    <DisplayName("Длина по торцу крыла")>
    Public Property b_heightBottomFace As Double
        Get
            Return heightBottomFace
        End Get
        Set(ByVal value As Double)
            heightBottomFace = value
        End Set
    End Property
    Private heightCornice As Double
    <Browsable(True)>
    <Description("Высота карниза, м")>
    <Category("Параметры")>
    <DisplayName("Высота карниза")>
    Public Property b_heightCornice As Double
        Get
            Return heightCornice
        End Get
        Set(ByVal value As Double)
            heightCornice = value
        End Set
    End Property
    Private widthCornice As Double
    <Browsable(True)>
    <Description("Ширина карниза, м")>
    <Category("Параметры")>
    <DisplayName("Ширина карниза")>
    Public Property b_widthCornice As Double
        Get
            Return widthCornice
        End Get
        Set(ByVal value As Double)
            widthCornice = value
        End Set
    End Property
    Private type As String
    <Browsable(True)>
    <Description("Тип крыла")>
    <Category("Параметры")>
    <DisplayName("Тип крыла")>
    Public Property b_type As String
        Get
            Return type
        End Get
        Set(ByVal value As String)
            type = value
        End Set
    End Property
    Private model As String
    <Browsable(True)>
    <Description("Имя модели")>
    <Category("Свойства")>
    <DisplayName("Имя модели")>
    Public Property b_model As String
        Get
            Return model
        End Get
        Set(ByVal value As String)
            model = value
        End Set
    End Property
End Class

'класс для выпадающего списка
Class PostTypeConverter
    Inherits StringConverter
    Dim userList As List(Of String) = New List(Of String)
    Public typeObject As String = ""
    'Public Sub New(ByRef typeObj As String)
    '    '    typeObject = typeObj
    '    GetStandardValues()
    'End Sub
    Public Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overrides Function GetStandardValuesExclusive(ByVal context As ITypeDescriptorContext) As Boolean
        Return False
    End Function

    Public Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
        'Dim pan As PanelProjectBridge = New PanelProjectBridge
        If IsNothing(context) = True Then
            Return New StandardValuesCollection(userList)
        End If
        If context.PropertyDescriptor.Name = "b_typePile" Then
            userList.Clear()
            userList.Add("Призматические сваи")
            userList.Add("Буровые сваи")
        ElseIf context.PropertyDescriptor.Name = "b_typeIcecutter" Then
            userList.Clear()
            userList.Add("Нет")
            userList.Add("Тип 1")
            userList.Add("Тип 2")
        End If
        Return New StandardValuesCollection(userList)
    End Function
End Class

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'свойство вспомогательных элементов
Public Class PropertyTLCObject
    Private nameElement As String
    <Browsable(True)>
    <Description("Имя объекта")>
    <Category("Свойства")>
    <DisplayName("Имя")>
    <[ReadOnly](True)>
    Public Property b_nameElement As String
        Get
            Return nameElement
        End Get
        Set(ByVal value As String)
            nameElement = value
        End Set
    End Property
    Private keyParameters As String
    <Browsable(True)>
    <Description("Свойство элемента")>
    <Category("Свойства")>
    <DisplayName("Свойство")>
    <[ReadOnly](True)>
    Public Property b_keyParameters As String
        Get
            Return keyParameters
        End Get
        Set(ByVal value As String)
            keyParameters = value
        End Set
    End Property
End Class
Public Class FuncBridge
    'функция заносит в словарь все элементы мостового сооружения
    Public Shared Function FuncFindAllObjectBridge(ByRef activDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByRef dictionaryObject As Dictionary(Of String, String(,))) As Boolean
        FuncFindAllObjectBridge = False
        If IsNothing(activDocument) = False Then
            For Each acEnt As DwgEntity In activDocument.ActiveSpace.Entities
                If TypeOf acEnt Is DwgEntity Then
                    Dim acEntity As DwgEntity = acEnt
                    Dim arrayData As String(,) = {}
                    Dim boolReadXdata As Boolean = FuncXRecords.FuncReadXData(acEntity, "PROJECT_BRIDGE", arrayData)
                    Dim idTempBridge As String = MathFunction.FuncFindValueToArray2d(arrayData, "BrigeID", 0, 1)
                    If idTempBridge Like idBridge Then
                        Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                        Dim keyParam As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1)
                        Dim idElement As String = MathFunction.FuncFindValueToArray2d(arrayData, "ElementID", 0, 1)
                        If dictionaryObject.Count = 0 Then
                            Dim arrayWrite As String(,) = {}
                            ReDim Preserve arrayWrite(4, 0)
                            arrayWrite(0, 0) = nameObject
                            arrayWrite(1, 0) = keyParam
                            arrayWrite(2, 0) = idElement
                            arrayWrite(3, 0) = idBridge
                            arrayWrite(4, 0) = acEntity.ObjectID
                            dictionaryObject.Add(nameObject, arrayWrite)
                        Else
                            If dictionaryObject.ContainsKey(nameObject) = False Then
                                Dim arrayWrite As String(,) = {}
                                ReDim Preserve arrayWrite(4, 0)
                                arrayWrite(0, 0) = nameObject
                                arrayWrite(1, 0) = keyParam
                                arrayWrite(2, 0) = idElement
                                arrayWrite(3, 0) = idBridge
                                arrayWrite(4, 0) = acEntity.ObjectID
                                dictionaryObject.Add(nameObject, arrayWrite)
                            Else
                                Dim arrayWrite As String(,) = dictionaryObject.Item(nameObject)
                                Dim countElement As Integer = arrayWrite.GetUpperBound(1) + 1
                                ReDim Preserve arrayWrite(4, countElement)
                                arrayWrite(0, countElement) = nameObject
                                arrayWrite(1, countElement) = keyParam
                                arrayWrite(2, countElement) = idElement
                                arrayWrite(3, countElement) = idBridge
                                arrayWrite(4, countElement) = acEntity.ObjectID
                                dictionaryObject.Item(nameObject) = arrayWrite
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Function
    'функция ищет в словаре необходимый элемент по его LOCALID (объекты dwg)
    Public Shared Function FuncFindObjectAllDictionaryByLocalID(ByRef activDocument As Topomatic.Dwg.Drawing, ByRef dictionaryObject As Dictionary(Of String, String(,)), ByVal LocalID As String, ByVal nameObject As String, ByRef objectEntity As DwgEntity, Optional boolErase As Boolean = False, Optional param As String = "") As Boolean
        FuncFindObjectAllDictionaryByLocalID = False
        If IsNothing(dictionaryObject) = True Then Return False
        If dictionaryObject.Count = 0 Then Return False
        If IsNothing(LocalID) = True Then Return False
        If LocalID.Trim.Length < 2 Then Return False
        If IsNothing(nameObject) = True Then Return False
        If nameObject.Trim.Length = 0 Then Return False
        If dictionaryObject.ContainsKey(nameObject) = True Then
            Dim arrayData As String(,) = dictionaryObject.Item(nameObject)
            If IsArray(arrayData) = True Then
                For i As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim tempLocalId As String = arrayData(2, i)
                    If LocalID.Trim Like tempLocalId Then
                        If param.Trim.Length > 0 Then
                            If param Like "-1" Then
                                If Not (Val(arrayData(1, i))) < 0 Then
                                    Continue For
                                End If
                            ElseIf param Like "1" Then
                                If Not (Val(arrayData(1, i))) > 0 Then
                                    Continue For
                                End If
                            End If
                        End If
                        If IsNothing(arrayData(4, i)) = False Then
                            If arrayData(4, i).Trim.Length > 1 Then
                                Dim hgObject As UInteger = CUInt(arrayData(4, i))
                                If boolErase = True Then
                                    arrayData(4, i) = ""
                                End If
                                Dim tempObject As DwgEntity = Nothing
                                Dim boolFindObect As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgObject, tempObject)
                                If boolFindObect = True Then
                                    If tempObject.GetType.ToString Like objectEntity.GetType.ToString Then
                                        objectEntity = tempObject
                                        Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End Function
    'функция ищет в словаре необходимый элемент по его LOCALID и возвращает его имя
    Public Shared Function FuncFindNameToObjectAllDictionaryByLocalID(ByRef activDocument As Topomatic.Dwg.Drawing, ByRef dictionaryObject As Dictionary(Of String, String(,)), ByVal LocalID As String, Optional boolErase As Boolean = False) As Boolean
        FuncFindNameToObjectAllDictionaryByLocalID = False
        If IsNothing(dictionaryObject) = True Then Return False
        If dictionaryObject.Count > 0 Then Return False
        If IsNothing(LocalID) = True Then Return False
        If LocalID.Trim.Length < 2 Then Return False
        For i As Integer = 0 To dictionaryObject.Count - 1
            Dim arrayData As String(,) = dictionaryObject.ElementAt(i).Value
            If IsArray(arrayData) = True Then
                For j As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim tempLocalId As String = arrayData(2, j)
                    If LocalID.Trim Like tempLocalId Then
                        If IsNothing(arrayData(4, j)) = False Then
                            If arrayData(4, j).Trim.Length > 1 Then
                                Dim hgObject As UInteger = CUInt(arrayData(4, j))
                                If boolErase = True Then
                                    arrayData(4, j) = ""
                                End If
                                If activDocument.ActiveSpace.Entities.ContainsHandle(hgObject) = True Then
                                    Return arrayData(0, j)
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
        Next i
    End Function
    'функция возвращает высоты (по верху) 4 точек опирания балки
    Public Shared Function FuncReturnElevationPointsBeam(ByVal lineDownBalka As DwgLine, ByVal userBeam As Beams, ByRef arrayElev As Double(,), Optional surf As Surface = Nothing) As Boolean
        FuncReturnElevationPointsBeam = False
        If IsNothing(lineDownBalka) = True Then Exit Function
        If lineDownBalka.Length = 0 Then Exit Function
        If IsNothing(userBeam) = True Then Exit Function
        If IsNothing(surf) = True Then Exit Function
        Erase arrayElev
        ReDim arrayElev(2, 3)
        '1.делаем смещение балки вверх
        Dim b As Double = lineDownBalka.EndPoint.Z - lineDownBalka.StartPoint.Z
        Dim c As Double = lineDownBalka.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, userBeam.height)
        'получаем новые координаты верха балки
        Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineDownBalka.StartPoint.X, lineDownBalka.StartPoint.Y), lineDownBalka.Rotation, deltaXZ.X)
        Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineDownBalka.EndPoint.X, lineDownBalka.EndPoint.Y), lineDownBalka.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineDownBalka.StartPoint.Z + deltaXZ.Y)
        Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineDownBalka.EndPoint.Z + deltaXZ.Y)
        Dim axisLineTopBeam As DwgLine = New DwgLine()
        axisLineTopBeam.StartPoint = startPointUpBalkaPr
        axisLineTopBeam.EndPoint = endPointUpBalkaPr
        '====================================================================================================================
        'смещение верх лево
        Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection4, -1 * userBeam.widthTop / 2)
        If dbCollection4.Count > 0 Then
            Dim tempLine As DwgLine = dbCollection4.Item(0)
            arrayElev(0, 0) = Math.Round(tempLine.StartPoint.Z, 3) 'начало
            arrayElev(0, 2) = Math.Round(tempLine.EndPoint.Z, 3) 'конец
            If IsNothing(surf) = False Then
                Try
                    Dim startHPoint As Double = surf.GetElevation(tempLine.StartPoint)
                    Dim endHPoint As Double = surf.GetElevation(tempLine.EndPoint)
                    arrayElev(1, 0) = Math.Round(startHPoint, 3)
                    arrayElev(2, 0) = Math.Round(startHPoint - tempLine.StartPoint.Z, 3)
                    arrayElev(1, 2) = Math.Round(endHPoint, 3)
                    arrayElev(2, 2) = Math.Round(endHPoint - tempLine.EndPoint.Z, 3)
                Catch ex As System.NullReferenceException
                    Return False
                End Try
            End If
        End If
        'смещение верх право
        Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection3, userBeam.widthTop / 2)
        If dbCollection3.Count > 0 Then
            Dim tempLine As DwgLine = dbCollection3.Item(0)
            arrayElev(0, 1) = Math.Round(tempLine.StartPoint.Z, 3)
            arrayElev(0, 3) = Math.Round(tempLine.EndPoint.Z, 3)
            If IsNothing(surf) = False Then
                Try
                    Dim startHPoint As Double = surf.GetElevation(tempLine.StartPoint)
                    Dim endHPoint As Double = surf.GetElevation(tempLine.EndPoint)

                    arrayElev(1, 1) = Math.Round(startHPoint, 3)
                    arrayElev(2, 1) = Math.Round(startHPoint - tempLine.StartPoint.Z, 3)
                    arrayElev(1, 3) = Math.Round(endHPoint, 3)
                    arrayElev(2, 3) = Math.Round(endHPoint - tempLine.EndPoint.Z, 3)
                Catch ex As System.NullReferenceException
                    Return False
                End Try
            End If
        End If
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функции для расчета
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция строит ось балки, без проверки ее по допуску на зазор
    Public Shared Function FuncCreateAxisBeam(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByVal startAxisPillar As DwgLine, ByVal deltaLenghtStartAxisPillars As Double, ByVal axisPolyline3D As Polyline3D, ByVal userBeam As Beams, ByVal leftBeam As Boolean, ByRef newLineShortBeam As DwgLine, Optional lenghtShortBeam As Double = 0, Optional endAxisBeamsPillar As DwgLine = Nothing, Optional endAxisPillar As DwgLine = Nothing, Optional surf As Surface = Nothing) As Boolean
        If IsNothing(startAxisPillar) = True Then
            Return False
        End If
        If IsNothing(axisPolyline3D) = True Then
            Return False
        End If
        Dim angle As Double = startAxisPillar.Rotation
        Dim angleReverse As Double = startAxisPillar.Rotation + Math.PI
        If angleReverse > Math.PI * 2 Then
            angleReverse -= Math.PI * 2
        End If
        '==============================================================================================================================
        'крайняя левая балка
        Dim boolFindIntersectPoint As Boolean = False
        Dim startIntersectPoint2D As Vector2D = New Vector2D(0, 0)
        Dim startPointPrBeam As Vector3D = New Vector3D(0, 0, 0)
        Dim endPointPrBeam As Vector3D = New Vector3D(0, 0, 0)
        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPolyline3D, startAxisPillar.StartPoint.Pos, startAxisPillar.EndPoint.Pos)
        If pointIntersectCollection.Count > 0 Then
            startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
            Dim startDistPrBeam As Double = -1
            Dim off As Double = -1
            Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
            If boolDist = True Then
                startDistPrBeam += deltaLenghtStartAxisPillars
                startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
            End If
            'вычисляем верх балки
            Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
            Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
            Dim startAlignPointLeftElevation As Double = 0
            Dim startAlignPointRightElevation As Double = 0
            Try
                startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                If startAlignPointLeftElevation > startAlignPointRightElevation Then
                    startAlignPointLeftElevation = startAlignPointRightElevation
                End If
            Catch ex As System.NullReferenceException
            End Try
            startPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            boolFindIntersectPoint = True
        Else
            'нет пересечения пытаемся удлиннить ось
            Dim newAxisPillar As DwgLine = New DwgLine
            newAxisPillar.StartPoint = startAxisPillar.StartPoint
            newAxisPillar.EndPoint = startAxisPillar.EndPoint
            Dim boolExtendLine As Boolean = FuncBridge.FuncExtendBearm(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
            'заново вытаемся найти пересечение
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
            If pointIntersectCollection.Count > 0 Then
                startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                Dim startDistPrBeam As Double = -1
                Dim off As Double = -1
                Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                If boolDist = True Then
                    startDistPrBeam += deltaLenghtStartAxisPillars
                    startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
                End If
                'вычисляем верх балки
                Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
                Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
                Dim startAlignPointLeftElevation As Double = 0
                Dim startAlignPointRightElevation As Double = 0
                Try
                    startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                    startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                    If startAlignPointLeftElevation > startAlignPointRightElevation Then
                        startAlignPointLeftElevation = startAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                startPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                boolFindIntersectPoint = True
            End If
        End If
        If boolFindIntersectPoint = False Then
            Return False
        End If
        '=========================================================================================================================
        'вычисляем второй конец балки
        If lenghtShortBeam > 0 Then
            endPointPrBeam = FuncFindEndPointPrBeam(axisPolyline3D, startPointPrBeam, lenghtShortBeam, surf, userBeam)
        ElseIf IsNothing(endAxisBeamsPillar) = False Then
            angle = endAxisBeamsPillar.Rotation
            angleReverse = endAxisBeamsPillar.Rotation + Math.PI
            If angleReverse > Math.PI * 2 Then
                angleReverse -= Math.PI * 2
            End If
            pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, endAxisBeamsPillar.StartPoint.Pos, endAxisBeamsPillar.EndPoint.Pos)
            If pointIntersectCollection.Count > 0 Then
                startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                'вычисляем верх балки
                Dim startDistPrBeam As Double = -1
                Dim off As Double = -1
                Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
                Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
                Dim startAlignPointLeftElevation As Double = 0
                Dim startAlignPointRightElevation As Double = 0
                Try
                    startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                    startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                    If startAlignPointLeftElevation > startAlignPointRightElevation Then
                        startAlignPointLeftElevation = startAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
            Else
                'нет пересечения пытаемся удлиннить ось
                Dim newAxisPillar As DwgLine = New DwgLine
                newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                Dim boolExtendLine As Boolean = FuncBridge.FuncExtendBearm(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                If pointIntersectCollection.Count > 0 Then
                    startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                    'вычисляем верх балки
                    Dim startDistPrBeam As Double = -1
                    Dim off As Double = -1
                    Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                    Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
                    Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
                    Dim startAlignPointLeftElevation As Double = 0
                    Dim startAlignPointRightElevation As Double = 0
                    Try
                        startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                        startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                        If startAlignPointLeftElevation > startAlignPointRightElevation Then
                            startAlignPointLeftElevation = startAlignPointRightElevation
                        End If
                    Catch ex As System.NullReferenceException
                    End Try
                    endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                End If
            End If
            If boolFindIntersectPoint = False Then
                Return False
            End If
        ElseIf IsNothing(endAxisPillar) = False Then
            'смотрим зазор
            Dim keyParam As String = FuncXRecords.FuncReadValueXData(endAxisPillar, "KeyParameters", "PROJECT_BRIDGE")
            If keyParam.Trim.Length > 0 Then
                Dim userAxisPillar As AxisHorizontalPillars = Nothing
                Try
                    userAxisPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisHorizontalPillars)(keyParam)
                Catch ex As Newtonsoft.Json.JsonException
                End Try
                Dim zazor As Double = 0
                If IsNothing(userAxisPillar) = False Then
                    If leftBeam = True Then
                        zazor = userAxisPillar.leftClearence
                    Else
                        zazor = userAxisPillar.rightClearence
                    End If
                End If
                pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, endAxisPillar.StartPoint.Pos, endAxisPillar.EndPoint.Pos)
                If pointIntersectCollection.Count > 0 Then
                    startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                    Dim startDistPrBeam As Double = -1
                    Dim off As Double = -1
                    Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                    If boolDist = True Then
                        startDistPrBeam = startDistPrBeam - userBeam.b - zazor / 2
                        'перевычисляем точку
                        startIntersectPoint2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, 0)
                        Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
                        Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
                        'вычисляем верх балки
                        Dim startAlignPointLeftElevation As Double = 0
                        Dim startAlignPointRightElevation As Double = 0
                        Try
                            startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                            startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                            If startAlignPointLeftElevation > startAlignPointRightElevation Then
                                startAlignPointLeftElevation = startAlignPointRightElevation
                            End If
                        Catch ex As System.NullReferenceException
                        End Try
                        endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                    End If
                Else
                    'нет пересечения пытаемся удлиннить ось
                    Dim newAxisPillar As DwgLine = New DwgLine
                    newAxisPillar.StartPoint = endAxisBeamsPillar.StartPoint
                    newAxisPillar.EndPoint = endAxisBeamsPillar.EndPoint
                    Dim boolExtendLine As Boolean = FuncBridge.FuncExtendBearm(newAxisPillar, newAxisPillar.Length, newAxisPillar.Length)
                    pointIntersectCollection = PolylineExtentions.GetIntersections(axisPolyline3D, newAxisPillar.StartPoint.Pos, newAxisPillar.EndPoint.Pos)
                    If pointIntersectCollection.Count > 0 Then
                        startIntersectPoint2D = pointIntersectCollection.ElementAt(0) 'осевая точка
                        'вычисляем верх балки
                        Dim startDistPrBeam As Double = -1
                        Dim off As Double = -1
                        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startIntersectPoint2D, startDistPrBeam, off)
                        Dim ptTopLeft As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, -1 * userBeam.widthTop / 2)
                        Dim ptTopRight As Vector2D = PolylineExtentions.StaOffsetToPos(axisPolyline3D, startDistPrBeam, userBeam.widthTop / 2)
                        Dim startAlignPointLeftElevation As Double = 0
                        Dim startAlignPointRightElevation As Double = 0
                        Try
                            startAlignPointLeftElevation = surf.GetElevation(ptTopLeft) - userBeam.dEarth 'высота верха лево
                            startAlignPointRightElevation = surf.GetElevation(ptTopRight) - userBeam.dEarth 'высота верха право
                            If startAlignPointLeftElevation > startAlignPointRightElevation Then
                                startAlignPointLeftElevation = startAlignPointRightElevation
                            End If
                        Catch ex As System.NullReferenceException
                        End Try
                        endPointPrBeam = New Vector3D(startIntersectPoint2D, startAlignPointLeftElevation - userBeam.height) 'положение оси балки(начальная точка)
                    End If
                End If
                If boolFindIntersectPoint = False Then
                    Return False
                End If
            End If
        End If
        If IsNothing(newLineShortBeam) = True Then
            newLineShortBeam = New DwgLine()
            ActivDocument.ActiveSpace.Entities.Add(newLineShortBeam)
        End If
        newLineShortBeam.StartPoint = startPointPrBeam
        newLineShortBeam.EndPoint = endPointPrBeam
        Return True
    End Function
    'функция ищет вторую точку опирания балки, без проверки допуска по зазору
    Public Shared Function FuncFindEndPointPrBeam(ByVal axisPolyline3D As Polyline3D, ByVal startPointPrBeam As Vector3D, ByVal lineShortBeam As Double, ByVal surf As Surface, ByVal userBeam As Beams) As Vector3D
        FuncFindEndPointPrBeam = New Vector3D(-1, -1, -1)
        Dim startDistPrBeam As Double = -1
        Dim off As Double = -1
        'находим расстояние на полилинии
        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPolyline3D, startPointPrBeam.Pos, startDistPrBeam, off)
        If boolDist = True Then
            Dim newDistEndPointPrBeam As Double = startDistPrBeam + lineShortBeam
            For i As Integer = 0 To 1000
                Dim centerPointBeam As Vector2D = Nothing
                Dim leftPointBeam As Vector2D = Nothing
                Dim rightPointBeam As Vector2D = Nothing
                Try
                    centerPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, 0)
                    leftPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, -1 * userBeam.widthTop / 2)
                    rightPointBeam = PolylineExtentions.StaOffsetToPos(axisPolyline3D, newDistEndPointPrBeam, userBeam.widthTop / 2)
                Catch ex As System.ArgumentOutOfRangeException
                    Return New Vector3D(-1, -1, -1)
                End Try
                Dim endAlignPointLeftElevation As Double = 0
                Dim endAlignPointRightElevation As Double = 0
                Try
                    endAlignPointLeftElevation = surf.GetElevation(leftPointBeam) - userBeam.dEarth 'высота верха лево
                    endAlignPointRightElevation = surf.GetElevation(rightPointBeam) - userBeam.dEarth 'высота верха право
                    If endAlignPointLeftElevation > endAlignPointRightElevation Then
                        endAlignPointLeftElevation = endAlignPointRightElevation
                    End If
                Catch ex As System.NullReferenceException
                End Try
                Dim tempPointNew As Vector3D = New Vector3D(centerPointBeam, endAlignPointLeftElevation)
                Dim tempDist As Double = (startPointPrBeam - tempPointNew).Length
                Dim dLenght As Double = lineShortBeam - tempDist
                If Math.Abs(dLenght) < 0.001 Then
                    Return tempPointNew
                Else
                    newDistEndPointPrBeam = newDistEndPointPrBeam + dLenght
                End If
            Next
        Else
            Return New Vector3D(-1, -1, -1)
        End If
    End Function
    'функция восстанавливает балку по ее параметрам возвращает координаты 8 точек (без рисования вспомогательных линий)
    Public Shared Function FuncRestoreElementsBeam(ByVal lineShortBearm As DwgLine, ByVal userBeam As Beams, ByRef startSectionBeam As List(Of Vector3D), ByRef endSectionBeam As List(Of Vector3D), Optional boolMonolitSites As Boolean = True, Optional boolBeamPs As Boolean = False, Optional boolFullExtend As Boolean = False) As Boolean
        FuncRestoreElementsBeam = False
        If IsNothing(lineShortBearm) = True Then Exit Function
        '1. Удлинняем балку на величину точек опирания
        Dim axisLineDownBeam As DwgLine = New DwgLine()
        axisLineDownBeam.StartPoint = lineShortBearm.StartPoint
        axisLineDownBeam.EndPoint = lineShortBearm.EndPoint
        'удлинняем балку на величину участков опирания балок 
        If boolFullExtend = True Or boolMonolitSites = True Then
            Dim boolExtBearm As Boolean = FuncExtendBearm(axisLineDownBeam, userBeam.a, userBeam.b)
        End If
        '1.делаем смещение балки вверх
        Dim b As Double = axisLineDownBeam.EndPoint.Z - axisLineDownBeam.StartPoint.Z
        Dim c As Double = axisLineDownBeam.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, userBeam.height)
        'получаем новые координаты верха балки
        Dim startPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(axisLineDownBeam.StartPoint.X, axisLineDownBeam.StartPoint.Y), axisLineDownBeam.Rotation, deltaXZ.X)
        Dim endPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(axisLineDownBeam.EndPoint.X, axisLineDownBeam.EndPoint.Y), axisLineDownBeam.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPosTopBeam3d As Vector3D = New Vector3D(startPosUpBeam2d.X, startPosUpBeam2d.Y, axisLineDownBeam.StartPoint.Z + deltaXZ.Y)
        Dim endPosTopBeam3d As Vector3D = New Vector3D(endPosUpBeam2d.X, endPosUpBeam2d.Y, axisLineDownBeam.EndPoint.Z + deltaXZ.Y)
        Dim axisLineTopBeam As DwgLine = New DwgLine()
        axisLineTopBeam.StartPoint = startPosTopBeam3d
        axisLineTopBeam.EndPoint = endPosTopBeam3d
        '====================================================================================================================
        Dim pointStart1 As Vector3D = Nothing
        Dim pointStart2 As Vector3D = Nothing
        Dim pointStart3 As Vector3D = Nothing
        Dim pointStart4 As Vector3D = Nothing
        Dim pointEnd1 As Vector3D = Nothing
        Dim pointEnd2 As Vector3D = Nothing
        Dim pointEnd3 As Vector3D = Nothing
        Dim pointEnd4 As Vector3D = Nothing
        'смещение низ право
        Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineDownBeam.Offset(dbCollection1, userBeam.widthBottom / 2)
        If dbCollection1.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection1.Item(0)
            pointStart1 = tempLine.StartPoint
            pointEnd1 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение низ лево
        Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineDownBeam.Offset(dbCollection2, -1 * userBeam.widthBottom / 2)
        If dbCollection2.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection2.Item(0)
            pointStart2 = tempLine.StartPoint
            pointEnd2 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение верх право
        Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection3, userBeam.widthTop / 2)
        If dbCollection3.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection3.Item(0)
            pointStart3 = tempLine.StartPoint
            pointEnd3 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение верх лево
        Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLineTopBeam.Offset(dbCollection4, -1 * userBeam.widthTop / 2)
        If dbCollection4.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection4.Item(0)
            pointStart4 = tempLine.StartPoint
            pointEnd4 = tempLine.EndPoint
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'исключаем участки омоноличивания
        If boolMonolitSites = True Then
            Dim boolExtendTopLeft As Boolean = MathFunction.FuncExtendPos(pointStart4, pointEnd4, -1 * userBeam.startLenghtMonolith, -1 * userBeam.endLenghtMonolith)
            Dim boolExtendTopRight As Boolean = MathFunction.FuncExtendPos(pointStart3, pointEnd3, -1 * userBeam.startLenghtMonolith, -1 * userBeam.endLenghtMonolith)
        End If
        'If boolMonolitSites = True Then
        '    If boolBeamPs = True Then
        '        'участки опирания уже отрезаны, убираем их из участков омоноличивания
        '        Dim startSiteMonolit As Double = userBeam.startLenghtMonolith
        '        If startSiteMonolit > userBeam.a Then
        '            startSiteMonolit -= userBeam.a
        '        End If
        '        Dim endSiteMonolit As Double = userBeam.endLenghtMonolith
        '        If endSiteMonolit > userBeam.b Then
        '            endSiteMonolit -= userBeam.b
        '        End If
        '        Dim boolExtendTopLeft As Boolean = MathFunction.FuncExtendPos(pointStart4, pointEnd4, -1 * startSiteMonolit, -1 * endSiteMonolit)
        '        Dim boolExtendTopRight As Boolean = MathFunction.FuncExtendPos(pointStart3, pointEnd3, -1 * startSiteMonolit, -1 * endSiteMonolit)
        '    Else
        '        Dim boolExtendTopLeft As Boolean = MathFunction.FuncExtendPos(pointStart4, pointEnd4, -1 * userBeam.startLenghtMonolith + userBeam.a, -1 * userBeam.endLenghtMonolith + userBeam.b)
        '        Dim boolExtendTopRight As Boolean = MathFunction.FuncExtendPos(pointStart3, pointEnd3, -1 * userBeam.startLenghtMonolith + userBeam.a, -1 * userBeam.endLenghtMonolith + userBeam.b)
        '    End If
        'End If
        'начальное сечение балки
        startSectionBeam = New List(Of Vector3D)
        startSectionBeam.Add(pointStart4) 'верх лево
        startSectionBeam.Add(pointStart3) 'верх право
        startSectionBeam.Add(pointStart2) 'низ лево
        startSectionBeam.Add(pointStart1) 'низ право
        'конечное сечение балки
        endSectionBeam = New List(Of Vector3D)
        endSectionBeam.Add(pointEnd4) 'верх лево
        endSectionBeam.Add(pointEnd3) 'верх право
        endSectionBeam.Add(pointEnd2) 'низ лево
        endSectionBeam.Add(pointEnd1) 'низ право
        Return True
    End Function
    'Функция возвращает 4 точки зазора от предыдущей балки
    Public Shared Function FuncCorrectPosBeamToZazor(ByVal prevLineShortBeam As DwgLine, ByVal lineShortBeam As DwgLine, ByRef listZazor As List(Of Double), Optional userBeam As Beams = Nothing) As Double
        FuncCorrectPosBeamToZazor = 0
        If IsNothing(lineShortBeam) = True Then Exit Function
        If IsNothing(prevLineShortBeam) = True Then Exit Function
        Dim userPrevBeam As Beams = Nothing
        Dim keyParamPrevBeam As String = FuncXRecords.FuncReadValueXData(prevLineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
        If keyParamPrevBeam.Trim.Length > 0 Then
            userPrevBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamPrevBeam)
        End If
        If IsNothing(userBeam) = True Then
            Dim keyParamBeam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParamBeam.Trim.Length > 0 Then
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
            End If
        End If
        If IsNothing(userBeam) = True Then Return 0
        If IsNothing(userPrevBeam) = True Then Return 0
        'получаем сечения предыдущей балки
        Dim startSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionPrevBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionPrevBeam As Boolean = FuncRestoreElementsBeam(prevLineShortBeam, userPrevBeam, startSectionPrevBeam, endSectionPrevBeam, False, False, True)
        'получаем сечения балки
        Dim startSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionBeam As Boolean = FuncRestoreElementsBeam(lineShortBeam, userBeam, startSectionBeam, endSectionBeam, False, False, True)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'приводим балки к одному уровню по низу
        'верх балки предыдущей
        Dim topLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(0) 'лево
        Dim topRightPrevBeam As Vector3D = endSectionPrevBeam.Item(1) 'право
        'низ балки предыдущей
        Dim downLeftPrevBeam As Vector3D = endSectionPrevBeam.Item(2) 'лево
        Dim downRightPrevBeam As Vector3D = endSectionPrevBeam.Item(3) 'право
        'высота по низу предыдущей балки
        Dim elevPrevBeam As Double = (downLeftPrevBeam.Z + downRightPrevBeam.Z) / 2
        'верх определяемая балка
        Dim topLeftBeam As Vector3D = startSectionBeam.Item(0) 'лево
        Dim topRightBeam As Vector3D = startSectionBeam.Item(1) 'право
        'низ балки определяемой
        Dim downLeftBeam As Vector3D = startSectionBeam.Item(2) 'лево
        Dim downRightBeam As Vector3D = startSectionBeam.Item(3) 'право
        'высота по низу предыдущей балки
        Dim elevBeam As Double = (downLeftBeam.Z + downRightBeam.Z) / 2
        'дельта по высоте
        Dim deltaH As Double = elevPrevBeam - elevBeam
        If Math.Abs(deltaH) > 0.01 Then
            Dim oldLenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length
            If deltaH < 0 Then
                'предудущая балка ниже чем расчетная (приводим ее уровню расчетной бплки
                downLeftPrevBeam = MathFunction.FuncCalcPointInLine(downLeftPrevBeam, topLeftPrevBeam, Math.Abs(deltaH), 3)
                downRightPrevBeam = MathFunction.FuncCalcPointInLine(downRightPrevBeam, topRightPrevBeam, Math.Abs(deltaH), 3)
                Dim lenghtDownWidthBeam As Double = (downLeftPrevBeam - downRightPrevBeam).Length - oldLenghtDownWidthBeam
                If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                    Dim dLen As Double = lenghtDownWidthBeam / 2
                    Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftPrevBeam, downRightPrevBeam, -1 * dLen, -1 * dLen, 3)
                End If
            Else 'расчетная балка ниже чем предыдущая (подтягиваем расчетную балку к предыдущей)
                downLeftBeam = MathFunction.FuncCalcPointInLine(downLeftBeam, topLeftBeam, deltaH)
                downRightBeam = MathFunction.FuncCalcPointInLine(downRightBeam, topRightBeam, deltaH)
                Dim lenghtDownWidthBeam As Double = (downLeftBeam - downRightBeam).Length - oldLenghtDownWidthBeam
                If Math.Abs(lenghtDownWidthBeam) > 0.001 Then
                    Dim dLen As Double = lenghtDownWidthBeam / 2
                    Dim boolExt As Boolean = MathFunction.FuncExtendPos(downLeftBeam, downRightBeam, -1 * dLen, -1 * dLen, 3)
                End If
            End If
        End If

        '====================================================================================================================
        'анализируем зазор по верху
        '====================================================================================================================
        'проверяем пересечение верха балки с горизонтальной линией предыдущей балки
        'верх лево
        Dim topLeftZazor As Double = 99999999
        Dim intersectPoint As Vector2D = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(0).Pos, topLeftBeam.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - topLeftBeam.Pos).Length
            If tempZazor < topLeftZazor Then
                topLeftZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            End If
        End If
        'пересечения нет, проверяем в обратную сторону
        If boolIntersect = False Then
            boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topLeftPrevBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(0).Pos, topLeftPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topLeftPrevBeam.Pos).Length
                    If tempZazor < topLeftZazor Then
                        topLeftZazor = tempZazor
                    End If
                End If
            End If
        End If
        'зазор не найден, удлинняем принудительно верх предыдущей балки и смотрим пересечение
        If topLeftZazor = 99999999 Then
            Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
            Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

            Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(0).Pos, topLeftBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - topLeftBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'верх право
        Dim topRightZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(topLeftPrevBeam.Pos, topRightPrevBeam.Pos, endSectionBeam.Item(1).Pos, topRightBeam.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - topRightBeam.Pos).Length
            If tempZazor < topRightZazor Then
                topRightZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, topLeftPrevBeam.Pos, topRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                If tempZazor < topRightZazor Then
                    topRightZazor = tempZazor
                End If
            End If
        End If
        'пересечения нет, проверяем в обратную сторону
        If boolIntersect = False Then
            boolIntersect = MathFunction.FuncIntersectPolyline(topLeftBeam.Pos, topRightBeam.Pos, startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - topRightPrevBeam.Pos).Length
                If tempZazor < topRightZazor Then
                    topRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(1).Pos, topRightPrevBeam.Pos, topLeftBeam.Pos, topRightBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - topRightPrevBeam.Pos).Length
                    If tempZazor < topRightZazor Then
                        topRightZazor = tempZazor
                    End If
                End If
            End If
        End If
        'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
        If topRightZazor = 99999999 Then
            Dim newTopLeftPrevBeam As Vector2D = topLeftPrevBeam.Pos
            Dim newTopRightPrevBeam As Vector2D = topRightPrevBeam.Pos

            Dim boolExt As Boolean = MathFunction.FuncExtendPos(newTopLeftPrevBeam, newTopRightPrevBeam, 0.5, 0.5)
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(1).Pos, topRightBeam.Pos, newTopLeftPrevBeam, newTopRightPrevBeam, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - topRightBeam.Pos).Length
                If tempZazor < topLeftZazor Then
                    topRightZazor = tempZazor
                End If
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'низ лево
        Dim downLeftZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(2).Pos, downLeftBeam.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - downLeftBeam.Pos).Length
            If tempZazor < downLeftZazor Then
                downLeftZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            End If
        End If
        'пересечения нет, проверяем в обратную сторону
        If boolIntersect = False Then
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downLeftPrevBeam.Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(2).Pos, downLeftPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downLeftPrevBeam.Pos).Length
                    If tempZazor < downLeftZazor Then
                        downLeftZazor = tempZazor
                    End If
                End If
            End If
        End If
        'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
        If downLeftZazor = 99999999 Then
            Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
            Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

            Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(2).Pos, downLeftBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - downLeftBeam.Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            End If
        End If

        'низ право
        Dim downRightZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(downLeftPrevBeam.Pos, downRightPrevBeam.Pos, endSectionBeam.Item(3).Pos, downRightBeam.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - downRightBeam.Pos).Length
            If tempZazor < downRightZazor Then
                downRightZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, downLeftPrevBeam.Pos, downRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            End If
        End If
        'пересечения нет, проверяем в обратную сторону
        If boolIntersect = False Then
            boolIntersect = MathFunction.FuncIntersectPolyline(downLeftBeam.Pos, downRightBeam.Pos, startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = -1 * (intersectPoint - downRightPrevBeam.Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            Else
                'пересечение виртуального пересечения
                boolIntersect = MathFunction.FuncIntersectionTwoRay(startSectionPrevBeam.Item(3).Pos, downRightPrevBeam.Pos, downLeftBeam.Pos, downRightBeam.Pos, intersectPoint)
                If boolIntersect = True Then
                    Dim tempZazor As Double = (intersectPoint - downRightPrevBeam.Pos).Length
                    If tempZazor < downRightZazor Then
                        downRightZazor = tempZazor
                    End If
                End If
            End If
        End If
        'зазор не найден, удлинняем принудительно верх предудущей балки и смотрим пересечение
        If downRightZazor = 99999999 Then
            Dim newDownLeftPrevBeam As Vector2D = downLeftPrevBeam.Pos
            Dim newDownRightPrevBeam As Vector2D = downRightPrevBeam.Pos

            Dim boolExt As Boolean = MathFunction.FuncExtendPos(newDownLeftPrevBeam, newDownRightPrevBeam, 0.5, 0.5)
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(endSectionBeam.Item(3).Pos, downRightBeam.Pos, newDownLeftPrevBeam, newDownRightPrevBeam, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - downRightBeam.Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            End If
        End If
        listZazor.Add(topLeftZazor)
        listZazor.Add(topRightZazor)
        listZazor.Add(downLeftZazor)
        listZazor.Add(downRightZazor)
        Return listZazor.Min
    End Function
    'Функция возвращает 4 точки зазора от сечения
    Public Shared Function FuncReturnPointZazorOfSection(ByVal lineShortBeam As DwgLine, ByVal lineSection As DwgLine, ByRef listZazor As List(Of Double), Optional userBeam As Beams = Nothing) As Double
        FuncReturnPointZazorOfSection = -9999
        If IsNothing(lineShortBeam) = True Then Exit Function
        If lineShortBeam.Length = 0 Then Exit Function
        If IsNothing(userBeam) = True Then
            Dim keyParamBeam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParamBeam.Trim.Length > 0 Then
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
            End If
        End If
        If IsNothing(userBeam) = True Then Exit Function
        If IsNothing(lineSection) = True Then Exit Function
        If lineSection.Length = 0 Then Exit Function
        'получаем сечения балки
        Dim startSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim endSectionBeam As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolSectionBeam As Boolean = FuncRestoreElementsBeam(lineShortBeam, userBeam, startSectionBeam, endSectionBeam, False, False, True)
        '====================================================================================================================
        'анализируем зазор по верху
        '====================================================================================================================
        'проверяем пересечение верха балки с горизонтальной линией предыдущей балки
        'верх лево
        Dim topLeftZazor As Double = 99999999
        Dim intersectPoint As Vector2D = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(0).Pos, startSectionBeam.Item(0).Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - startSectionBeam.Item(0).Pos).Length
            If tempZazor < topLeftZazor Then
                topLeftZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(0).Pos, startSectionBeam.Item(0).Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - startSectionBeam.Item(0).Pos).Length
                If tempZazor < topLeftZazor Then
                    topLeftZazor = tempZazor
                End If
            End If
        End If
        'верх право
        Dim topRightZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(1).Pos, startSectionBeam.Item(1).Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - startSectionBeam.Item(1).Pos).Length
            If tempZazor < topRightZazor Then
                topRightZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(1).Pos, startSectionBeam.Item(1).Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - startSectionBeam.Item(1).Pos).Length
                If tempZazor < topRightZazor Then
                    topRightZazor = tempZazor
                End If
            End If
        End If
        'низ лево
        Dim downLeftZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(2).Pos, startSectionBeam.Item(2).Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - startSectionBeam.Item(2).Pos).Length
            If tempZazor < downLeftZazor Then
                downLeftZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(2).Pos, startSectionBeam.Item(2).Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - startSectionBeam.Item(2).Pos).Length
                If tempZazor < downLeftZazor Then
                    downLeftZazor = tempZazor
                End If
            End If
        End If
        'низ право
        Dim downRightZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(3).Pos, startSectionBeam.Item(3).Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - startSectionBeam.Item(3).Pos).Length
            If tempZazor < downRightZazor Then
                downRightZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endSectionBeam.Item(3).Pos, startSectionBeam.Item(3).Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - startSectionBeam.Item(3).Pos).Length
                If tempZazor < downRightZazor Then
                    downRightZazor = tempZazor
                End If
            End If
        End If
        listZazor.Add(topLeftZazor)
        listZazor.Add(topRightZazor)
        listZazor.Add(downLeftZazor)
        listZazor.Add(downRightZazor)
        Return listZazor.Min
    End Function
    'функция возвращает 2 точки зазора по оси балки
    Public Shared Function FuncReturnPointCentreZazorOfSection(ByVal lineShortBeam As DwgLine, ByVal lineSection As DwgLine, ByRef listZazor As List(Of Double), Optional userBeam As Beams = Nothing) As Double
        FuncReturnPointCentreZazorOfSection = -9999
        If IsNothing(lineShortBeam) = True Then Exit Function
        If lineShortBeam.Length = 0 Then Exit Function
        If IsNothing(userBeam) = True Then
            Dim keyParamBeam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParamBeam.Trim.Length > 0 Then
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
            End If
        End If
        If IsNothing(userBeam) = True Then Exit Function
        If IsNothing(lineSection) = True Then Exit Function
        If lineSection.Length = 0 Then Exit Function
        'получаем сечения балки
        '1. Удлинняем балку на величину точек опирания
        Dim boolExtBearm As Boolean = FuncExtendBearm(lineShortBeam, userBeam.a, userBeam.b)
        '1.делаем смещение балки вверх
        Dim b As Double = lineShortBeam.EndPoint.Z - lineShortBeam.StartPoint.Z
        Dim c As Double = lineShortBeam.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, userBeam.height)
        'получаем новые координаты верха балки
        Dim startPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineShortBeam.StartPoint.X, lineShortBeam.StartPoint.Y), lineShortBeam.Rotation, deltaXZ.X)
        Dim endPosUpBeam2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineShortBeam.EndPoint.X, lineShortBeam.EndPoint.Y), lineShortBeam.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPosTopBeam3d As Vector3D = New Vector3D(startPosUpBeam2d.X, startPosUpBeam2d.Y, lineShortBeam.StartPoint.Z + deltaXZ.Y)
        Dim endPosTopBeam3d As Vector3D = New Vector3D(endPosUpBeam2d.X, endPosUpBeam2d.Y, lineShortBeam.EndPoint.Z + deltaXZ.Y)
        '====================================================================================================================
        'анализируем зазор по верху
        '====================================================================================================================
        'проверяем пересечение верха балки с горизонтальной линией предыдущей балки
        'верх лево
        Dim topZazor As Double = 99999999
        Dim intersectPoint As Vector2D = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endPosTopBeam3d.Pos, startPosTopBeam3d.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - startPosTopBeam3d.Pos).Length
            If tempZazor < topZazor Then
                topZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, endPosTopBeam3d.Pos, startPosTopBeam3d.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - startPosTopBeam3d.Pos).Length
                If tempZazor < topZazor Then
                    topZazor = tempZazor
                End If
            End If
        End If
        'низ лево
        Dim downZazor As Double = 99999999
        intersectPoint = New Vector2D(-1, -1)
        'проверяем на предмет явного пересечения
        boolIntersect = MathFunction.FuncIntersectPolyline(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, lineShortBeam.EndPoint.Pos, lineShortBeam.StartPoint.Pos, intersectPoint)
        If boolIntersect = True Then
            Dim tempZazor As Double = -1 * (intersectPoint - lineShortBeam.StartPoint.Pos).Length
            If tempZazor < downZazor Then
                downZazor = tempZazor
            End If
        Else
            'пересечение виртуального пересечения
            boolIntersect = MathFunction.FuncIntersectionTwoRay(lineSection.StartPoint.Pos, lineSection.EndPoint.Pos, lineShortBeam.EndPoint.Pos, lineShortBeam.StartPoint.Pos, intersectPoint)
            If boolIntersect = True Then
                Dim tempZazor As Double = (intersectPoint - lineShortBeam.StartPoint.Pos).Length
                If tempZazor < downZazor Then
                    downZazor = tempZazor
                End If
            End If
        End If
        listZazor.Add(topZazor)
        listZazor.Add(downZazor)
        Return listZazor.Min
    End Function
    'функция корректирует балку по высоте
    Public Shared Function FuncCorrectPosBeamToElevation(ByVal ActivDocument As Topomatic.Dwg.Drawing, ByRef lineShortBeam As DwgLine, ByVal axisPline3D As Polyline3D, ByVal surf As Surface, Optional userBeam As Beams = Nothing) As Boolean
        FuncCorrectPosBeamToElevation = False
        If IsNothing(lineShortBeam) = True Then
            Return False
        ElseIf lineShortBeam.Length = 0 Then
            Return False
        End If
        If IsNothing(surf) = True Then Return False

        If IsNothing(userBeam) = True Then
            Dim keyParam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParam.Trim.Length = 0 Then
                Return False
            End If
            Try
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParam)
            Catch ex As Newtonsoft.Json.JsonException
                Return False
            End Try
        End If

        If IsNothing(userBeam) = False Then
            Dim axisOffset As Double = userBeam.axisStartOffset
            Dim deltaElevation As Double = userBeam.dEarth
            Dim startPointBeam As Vector3D = lineShortBeam.StartPoint
            Dim endPointBeam As Vector3D = lineShortBeam.EndPoint
            For i As Integer = 0 To 100
                'восстанавливаем балку
                Dim startSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam As Boolean = FuncBridge.FuncRestoreElementsBeam(lineShortBeam, userBeam, startSectionPoint3d, endSectionPoint3d, False, True, False)
                Try
                    'корректируем балку по высоте
                    Dim elev1 As Double = surf.GetElevation(startSectionPoint3d.Item(0).Pos)
                    Dim elevF1 As Double = startSectionPoint3d.Item(0).Z
                    Dim elevTopStart1 As Double = elev1 - (deltaElevation + startSectionPoint3d.Item(0).Z)

                    Dim elev2 As Double = surf.GetElevation(startSectionPoint3d.Item(1).Pos)
                    Dim elevF2 As Double = startSectionPoint3d.Item(1).Z
                    Dim elevTopStart2 As Double = elev2 - (deltaElevation + startSectionPoint3d.Item(1).Z)

                    Dim elev3 As Double = surf.GetElevation(endSectionPoint3d.Item(0).Pos)
                    Dim elevF3 As Double = endSectionPoint3d.Item(0).Z
                    Dim elevTopStart3 As Double = elev3 - (deltaElevation + endSectionPoint3d.Item(0).Z)

                    Dim elevF4 As Double = endSectionPoint3d.Item(1).Z
                    Dim elev4 As Double = surf.GetElevation(endSectionPoint3d.Item(1).Pos)
                    Dim elevTopStart4 As Double = elev4 - (deltaElevation + endSectionPoint3d.Item(1).Z)

                    If elevTopStart1 < 0 And elevTopStart2 < 0 Then
                        If Math.Abs(elevTopStart1) > Math.Abs(elevTopStart2) Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        End If
                    ElseIf elevTopStart1 > 0 And elevTopStart2 > 0 Then
                        If elevTopStart1 > elevTopStart2 Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        End If
                    ElseIf elevTopStart1 < 0 And elevTopStart2 > 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                    ElseIf elevTopStart1 > 0 And elevTopStart2 < 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                    End If

                    If elevTopStart3 < 0 And elevTopStart4 < 0 Then
                        If Math.Abs(elevTopStart3) > Math.Abs(elevTopStart4) Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        End If
                    ElseIf elevTopStart3 > 0 And elevTopStart4 > 0 Then
                        If elevTopStart3 > elevTopStart4 Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        End If
                    ElseIf elevTopStart3 < 0 And elevTopStart4 > 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                    ElseIf elevTopStart3 > 0 And elevTopStart4 < 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                    End If
                    lineShortBeam.StartPoint = startPointBeam
                    lineShortBeam.EndPoint = endPointBeam
                    Return True
                    'делаем контрольную проверку
                    'Dim lContr As Double = (startPointBeam - endPointBeam).Length
                    'Dim deltaLenght As Double = lineShortBeam.Length - lContr
                    'If Math.Abs(deltaLenght) >= 0.001 Then
                    '    'корректируем балку по длине
                    '    Dim boolExtendBeam As Boolean = FuncExtendBearm(lineShortBeam, 0, deltaLenght, 3)
                    '    'садим балку точно на ось
                    '    Dim pk As Double = 0
                    '    Dim off As Double = 0
                    '    Dim boolFindPk As Boolean = axisPline3D.PosToStaOffset(endPointBeam.Pos, pk, off)
                    '    If boolFindPk = True Then
                    '        Dim newEndPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPline3D, pk, 0)
                    '        endPointBeam = New Vector3D(newEndPoint, endPointBeam.Z)
                    '        lineShortBeam.EndPoint = endPointBeam
                    '    Else
                    '        Exit For
                    '    End If
                    'Else
                    '    Exit For
                    'End If
                Catch ex As System.NullReferenceException
                    Return False
                End Try
            Next i
            'Dim lContrFinal As Double = (startPointBeam - endPointBeam).Length
            'Dim deltaLenghtFinal As Double = lineShortBeam.Length - lContrFinal
            'If Math.Abs(deltaLenghtFinal) <= 0.001 Then
            '    Return True
            'End If
        End If
    End Function
    'функция корректирует балку по длине
    Public Shared Function FuncCorrectionLenghtBeam(ByVal axisPline3d As Polyline3D, ByRef lineShortBeam As DwgLine, Optional userBeam As Beams = Nothing, Optional lenghtBeam As Double = 0) As Boolean
        FuncCorrectionLenghtBeam = False
        If IsNothing(axisPline3d) = True Then Exit Function
        If IsNothing(lineShortBeam) = True Then Exit Function
        If lineShortBeam.Length = 0 Then Exit Function
        If IsNothing(userBeam) = True Then
            Dim keyParamBeam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParamBeam.Trim.Length > 0 Then
                Try
                    userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
                    lenghtBeam = userBeam.fullLenght - userBeam.a - userBeam.b
                Catch ex As Newtonsoft.Json.JsonException
                End Try
            End If
        End If
        lenghtBeam = userBeam.fullLenght - userBeam.a - userBeam.b
        If lenghtBeam > 0 Then
            For i As Integer = 0 To 10
                Dim startPK As Double = 0
                Dim off As Double = 0
                Dim startPoint3d As Vector3D = lineShortBeam.StartPoint
                Dim boolStartPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPline3d, lineShortBeam.StartPoint.Pos, startPK, off)
                If boolStartPoint = True Then
                    Dim tempPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPline3d, startPK, 0)
                    startPoint3d = New Vector3D(tempPoint, startPoint3d.Z)
                End If
                Dim endPK As Double = 0
                Dim endPoint3d As Vector3D = lineShortBeam.EndPoint
                Dim boolEndPoint As Boolean = PolylineExtentions.PosToStaOffset(axisPline3d, lineShortBeam.EndPoint.Pos, endPK, off)
                If boolEndPoint = True Then
                    Dim tempPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPline3d, endPK, 0)
                    endPoint3d = New Vector3D(tempPoint, endPoint3d.Z)
                End If
                lineShortBeam.StartPoint = startPoint3d
                lineShortBeam.EndPoint = endPoint3d
                Dim newLenght As Double = Math.Round((endPoint3d - startPoint3d).Length, 3)
                Dim deltaLenght As Double = Math.Round(lenghtBeam - newLenght, 3)
                If deltaLenght >= 0.001 Then
                    Dim boolExtendBeam As Boolean = FuncExtendBearm(lineShortBeam, 0, deltaLenght, 3)
                Else
                    Return True
                End If
            Next
        End If
    End Function
    'функция вычисляет середину между смежными балками
    Public Shared Function FuncCalculateMiddlePointBeams(ByVal prevLineShortBeam As DwgLine, ByVal lineShortBeam As DwgLine, Optional userBeam As Beams = Nothing) As Vector3D
        FuncCalculateMiddlePointBeams = Nothing
        If IsNothing(lineShortBeam) = True Then Exit Function
        If IsNothing(prevLineShortBeam) = True Then Exit Function
        Dim userPrevBeam As Beams = Nothing
        Dim keyParamPrevBeam As String = FuncXRecords.FuncReadValueXData(prevLineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
        If keyParamPrevBeam.Trim.Length > 0 Then
            userPrevBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamPrevBeam)
        End If
        If IsNothing(userBeam) = True Then
            Dim keyParamBeam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
            If keyParamBeam.Trim.Length > 0 Then
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
            End If
        End If
        If IsNothing(userBeam) = True Then Return Nothing
        If IsNothing(userPrevBeam) = True Then Return Nothing
        Dim startPointPrevBeam As Vector3D = New Vector3D()
        Dim endPointPrevBeam As Vector3D = New Vector3D()
        Dim boolFindPoint As Boolean = FuncVirtualExtendBearm(prevLineShortBeam, 0, userPrevBeam.b, startPointPrevBeam, endPointPrevBeam)

        Dim startPointBeam As Vector3D = New Vector3D()
        Dim endPointBeam As Vector3D = New Vector3D()
        boolFindPoint = FuncVirtualExtendBearm(lineShortBeam, userBeam.a, 0, startPointBeam, endPointBeam)

        Dim middlePoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(endPointPrevBeam, startPointBeam)
        Return middlePoint
    End Function
    'функция проверяет и корректирует направление осей опоры и осей опирания (начало - лево, конец - право)
    Public Shared Function FuncСorrectDirectionAxisPillar(ByRef acAxisPillar As DwgLine, ByVal align As Alignment) As Boolean
        FuncСorrectDirectionAxisPillar = False
        If IsNothing(acAxisPillar) = True Then Exit Function
        If IsNothing(align) = True Then Exit Function
        Try
            Dim startPK As Double = -1
            Dim startOffset As Double = -1
            Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(acAxisPillar.StartPoint, startPK, startOffset)
            If boolPk = True Then
                If startOffset > 0 Then
                    Dim tempStartPoint As Vector3D = acAxisPillar.StartPoint
                    Dim tempEndPoint As Vector3D = acAxisPillar.EndPoint
                    acAxisPillar.StartPoint = tempEndPoint
                    acAxisPillar.EndPoint = tempStartPoint
                    Return True
                End If
            End If
        Catch ex As System.Exception
        End Try
        Try
            Dim endPK As Double = -1
            Dim endOffset As Double = -1
            Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(acAxisPillar.EndPoint, endPK, endOffset)
            If boolPk = True Then
                If endOffset < 0 Then
                    Dim tempStartPoint As Vector3D = acAxisPillar.StartPoint
                    Dim tempEndPoint As Vector3D = acAxisPillar.EndPoint
                    acAxisPillar.StartPoint = tempEndPoint
                    acAxisPillar.EndPoint = tempStartPoint
                    Return True
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function
    'функция возвращает 4 точки области перекрытия двух линий
    Public Shared Function FuncFindRectangleByLines(ByRef acLine1 As DwgLine, ByVal acLine2 As DwgLine, ByRef arrayPoint As Vector3D()) As Boolean
        If IsNothing(acLine1) = True Then
            Return False
        End If
        If IsNothing(acLine2) = True Then
            Return False
        End If
        If acLine1.Length = 0 Then
            Return False
        End If
        If acLine2.Length = 0 Then
            Return False
        End If
        ReDim Preserve arrayPoint(3)
        Dim angle1 As Double = acLine1.Rotation + Math.PI / 2
        If angle1 >= Math.PI * 2 Then
            angle1 -= Math.PI * 2
        End If
        Dim reverseAngle1 As Double = acLine1.Rotation - Math.PI / 2
        If reverseAngle1 < 0 Then
            reverseAngle1 += Math.PI * 2
        End If
        Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine1.StartPoint.Pos, angle1, 500)
        Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine1.StartPoint.Pos, reverseAngle1, 500)
        Dim tempLine1 As DwgLine = New DwgLine()
        tempLine1.StartPoint = New Vector3D(pt1, 0)
        tempLine1.EndPoint = New Vector3D(pt2, 0)

        Dim pt3 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine1.EndPoint.Pos, angle1, 500)
        Dim pt4 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine1.EndPoint.Pos, reverseAngle1, 500)
        Dim tempLine2 As DwgLine = New DwgLine()
        tempLine2.StartPoint = New Vector3D(pt3, 0)
        tempLine2.EndPoint = New Vector3D(pt4, 0)

        Dim angle2 As Double = acLine2.Rotation + Math.PI / 2
        If angle2 >= Math.PI * 2 Then
            angle2 -= Math.PI * 2
        End If
        Dim reverseAngle2 As Double = acLine2.Rotation - Math.PI / 2
        If reverseAngle2 < 0 Then
            reverseAngle2 += Math.PI * 2
        End If
        Dim pt5 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine2.StartPoint.Pos, angle2, 500)
        Dim pt6 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine2.StartPoint.Pos, reverseAngle2, 500)
        Dim tempLine3 As DwgLine = New DwgLine()
        tempLine3.StartPoint = New Vector3D(pt5, 0)
        tempLine3.EndPoint = New Vector3D(pt6, 0)

        Dim pt7 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine2.EndPoint.Pos, angle2, 500)
        Dim pt8 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine2.EndPoint.Pos, reverseAngle2, 500)
        Dim tempLine4 As DwgLine = New DwgLine()
        tempLine4.StartPoint = New Vector3D(pt7, 0)
        tempLine4.EndPoint = New Vector3D(pt8, 0)

        Dim pointIntersect As Vector2D = New Vector2D(0, 0)
        'первое пересечение Линия 2 длинее линии1
        Dim boolIntersect As Boolean = MathFunction.FuncIntersectionTwoSegments(tempLine1.StartPoint.Pos, tempLine1.EndPoint.Pos, acLine2.StartPoint.Pos, acLine2.EndPoint.Pos, pointIntersect)
        If boolIntersect = True Then
            arrayPoint(0) = acLine1.StartPoint
            Dim l As Double = (acLine1.StartPoint.Pos - pointIntersect).Length
            Dim pointIntersect3D As Vector3D = MathFunction.FuncCalcPointInLine(acLine2.StartPoint, acLine2.EndPoint, l)
            arrayPoint(1) = pointIntersect3D
        End If
        'первое пересечение Линия 1 длинее линии 2
        If boolIntersect = False Then
            'tempPt1 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine2.StartPoint.Pos, angle2, -10)
            pointIntersect = New Vector2D(0, 0)
            boolIntersect = MathFunction.FuncIntersectionTwoSegments(tempLine3.StartPoint.Pos, tempLine3.EndPoint.Pos, acLine1.StartPoint.Pos, acLine1.EndPoint.Pos, pointIntersect)
            If boolIntersect = True Then
                Dim l As Double = (acLine1.StartPoint.Pos - pointIntersect).Length
                Dim pointIntersect3D As Vector3D = MathFunction.FuncCalcPointInLine(acLine1.StartPoint, acLine1.EndPoint, l)
                arrayPoint(0) = pointIntersect3D
                arrayPoint(1) = acLine2.StartPoint
            End If
        End If

        pointIntersect = New Vector2D(0, 0)
        'первое пересечение Линия 2 длинее линии1
        Dim boolIntersect2 As Boolean = MathFunction.FuncIntersectionTwoSegments(tempLine4.StartPoint.Pos, tempLine4.EndPoint.Pos, acLine1.StartPoint.Pos, acLine1.EndPoint.Pos, pointIntersect)
        If boolIntersect2 = True Then
            arrayPoint(2) = acLine2.EndPoint
            Dim l As Double = (acLine1.StartPoint.Pos - pointIntersect).Length
            Dim pointIntersect3D As Vector3D = MathFunction.FuncCalcPointInLine(acLine1.StartPoint, acLine1.EndPoint, l)
            arrayPoint(3) = pointIntersect3D
        End If
        'первое пересечение Линия 2 длинее линии1
        If boolIntersect2 = False Then
            'tempPt2 = MathFunction.funcCalcCoordinatesByInsPointAndAngle(acLine1.EndPoint.Pos, angle1, -10)
            pointIntersect = New Vector2D(0, 0)
            boolIntersect2 = MathFunction.FuncIntersectionTwoSegments(tempLine2.StartPoint.Pos, tempLine2.EndPoint.Pos, acLine2.StartPoint.Pos, acLine2.EndPoint.Pos, pointIntersect)
            If boolIntersect = True Then
                Dim l As Double = (acLine2.StartPoint.Pos - pointIntersect).Length
                Dim pointIntersect3D As Vector3D = MathFunction.FuncCalcPointInLine(acLine2.StartPoint, acLine2.EndPoint, l)
                arrayPoint(2) = pointIntersect3D
                arrayPoint(3) = acLine1.EndPoint
            End If
        End If
        If boolIntersect = True And boolIntersect2 = True Then
            Return True
        End If
    End Function
    'функция находит координаты точки на полилинии зная длину приращения
    Public Shared Function FuncFindPositionPointPolylineByDist(ByVal polyline3d As Polyline3D, ByVal startPK As Double, ByVal lenght As Double, ByRef pt As Vector2D) As Boolean
        FuncFindPositionPointPolylineByDist = False
        Dim off As Double = 0
        Dim startPoint As Vector2D = polyline3d.StaOffsetToPos(startPK, off)
        For i As Integer = 0 To 100
            Dim newDist As Double = startPK + lenght
            Dim startNewPoint As Vector2D = polyline3d.StaOffsetToPos(newDist, off)
            Dim distContr As Double = (startNewPoint - startPoint).Length
            If Math.Abs(distContr - lenght) > 0.001 Then
                lenght = distContr
            Else
                pt = startNewPoint
                Return True
            End If
        Next
    End Function
    'коррекция горизонтального расстояния за угол наклона линии
    Public Shared Function FuncCorrectionLenghtBeamToElevation(ByVal axisPline As DwgPolyline, ByVal startPoint As Vector3D, ByVal radius As Double, ByVal surf As Surface, ByVal heightBearm As Double, ByVal dEarth As Double) As Vector3D
        Dim axisPline3D As IPolyline3D = New Polyline3D()
        axisPline.GetPolyline(axisPline3D)
        Dim startdist As Double = -1
        Dim off As Double = -1
        'находим расстояние на полилинии
        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startPoint.Pos, startdist, off)
        If boolDist = True Then
            Dim findDist As Double = radius
            For i As Integer = 0 To 1000
                Dim delta As Double = startdist + findDist
                Dim tempPoint As Vector2D = Nothing
                Try
                    tempPoint = PolylineExtentions.StaOffsetToPos(axisPline3D, delta, 0)
                Catch ex As System.ArgumentOutOfRangeException
                    Return New Vector3D(-1, -1, -1)
                End Try
                Dim elev As Double = startPoint.Z
                Try
                    elev = surf.GetElevation(tempPoint)
                Catch ex As System.NullReferenceException
                    Return New Vector3D(-1, -1, -1)
                End Try
                elev = elev - heightBearm - dEarth
                Dim tempPointNew As Vector3D = New Vector3D(tempPoint.X, tempPoint.Y, elev)
                Dim tempDist As Double = (startPoint - tempPointNew).Length
                Dim dLenght As Double = radius - tempDist
                If Math.Abs(dLenght) < 0.001 Then
                    Return tempPointNew
                Else
                    findDist = findDist + dLenght
                End If
            Next
        Else
            Return New Vector3D(-1, -1, -1)
        End If
    End Function
    'функция проверяет и корректирует балку еcли она против направления пикетажа
    Public Shared Function FuncCorrectionAxisDirectionBeam(ByRef lineBearm As DwgLine, ByRef align As Alignment) As Boolean
        FuncCorrectionAxisDirectionBeam = False
        If IsNothing(lineBearm) = True Then
            Return False
        End If
        If IsNothing(align) = True Then
            Return False
        End If
        If lineBearm.Length = 0 Then
            Return False
        End If
        'находим пикеты начала и конца предудущей балки
        Dim pkStart As Double = 0
        Dim offStart As Double = 0
        Dim boolFindPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(lineBearm.StartPoint.Pos, pkStart, offStart)
        Dim pkEnd As Double = 0
        Dim offend As Double = 0
        Dim boolFindPk1 As Boolean = align.Plan.CompoundLine.PosToStaOffset(lineBearm.EndPoint.Pos, pkEnd, offend)
        If boolFindPk = True And boolFindPk1 = True Then
            If pkStart > pkEnd Then
                Dim tempStartPoint As Vector3D = lineBearm.StartPoint
                Dim tempEndPoint As Vector3D = lineBearm.EndPoint
                lineBearm.StartPoint = tempEndPoint
                lineBearm.EndPoint = tempStartPoint
                Return True
            End If
        End If
    End Function


    'функция перемещает ось опоры вдоль трассы на указанную величину
    Public Shared Function FuncMoveAxisPillarToAlignment(ByVal alignPline3D As Polyline3D, ByRef axisPillar As DwgLine, ByVal offset As Double) As Boolean
        FuncMoveAxisPillarToAlignment = False
        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(alignPline3D, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
        If pointIntersectCollection.Count > 0 Then
            Dim tempStartPoint As Vector2D = pointIntersectCollection(0)
            Dim angleStart As Double = (axisPillar.StartPoint.Pos - tempStartPoint).Angle
            Dim distStart As Double = (axisPillar.StartPoint.Pos - tempStartPoint).Length

            Dim angleEnd As Double = (axisPillar.EndPoint.Pos - tempStartPoint).Angle
            Dim distEnd As Double = (axisPillar.EndPoint.Pos - tempStartPoint).Length

            Dim pkTemp As Double = 0
            Dim pkoff As Double = 0
            Dim boolPk As Boolean = alignPline3D.PosToStaOffset(tempStartPoint, pkTemp, pkoff)
            If boolPk = True Then
                Dim newPiketAxisPillar As Double = pkTemp + offset
                Dim newTempStartPoint As Vector2D = New Vector2D()
                Try
                    newTempStartPoint = alignPline3D.StaOffsetToPos(newPiketAxisPillar, 0)
                    'newTempStartPoint = alignPline3D.StaOffsetToPos(pkTemp, 0)
                    If newTempStartPoint.X <> 0 And newTempStartPoint.Y <> 0 Then
                        'создаем произвольный вектор 
                        Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angleStart, distStart)
                        Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angleEnd, distEnd)
                        axisPillar.StartPoint = pt1
                        axisPillar.EndPoint = pt2
                        Return True
                    End If
                Catch ex As System.ArgumentOutOfRangeException
                End Try
            End If
        End If
    End Function
    'функция перемещает ось опоры вдоль трассы на заданный пикет
    Public Shared Function FuncMoveAxisPillarToAlignmentPK(ByVal align As Alignment, ByRef axisPillar As DwgLine, ByVal station As Double) As Boolean
        FuncMoveAxisPillarToAlignmentPK = False
        Dim newTempStartPoint As Vector2D = New Vector2D()
        Dim boolPk As Boolean = align.Plan.CompoundLine.StaOffsetToPos(station, 0, newTempStartPoint)
        If boolPk = True Then
            Dim angle As Double = axisPillar.Rotation
            Dim reverseAngle As Double = angle + Math.PI
            If reverseAngle >= Math.PI * 2 Then
                reverseAngle -= Math.PI * 2
            End If
            'создаем произвольный вектор 
            Dim distStart As Double = (newTempStartPoint - axisPillar.StartPoint.Pos).Length
            Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angle, distStart)
            Dim distend As Double = (newTempStartPoint - axisPillar.EndPoint.Pos).Length
            Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, reverseAngle, distend)
            axisPillar.StartPoint = New Vector3D(pt1, axisPillar.StartPoint.Z)
            axisPillar.EndPoint = New Vector3D(pt2, axisPillar.EndPoint.Z)
            Return True
        End If
    End Function
    'функция возвращает все балки мостового сооружения для выбранного пролета
    Public Shared Function FuncFindAllBeamsInProlet(ByVal numberProlet As Integer, ByVal dictionaryObject As Dictionary(Of String, String(,)), ByRef arrayBeams As String(,)) As Integer
        FuncFindAllBeamsInProlet = False
        Dim countBeams As Integer = 0
        If IsNothing(dictionaryObject) = True Then Return False
        If dictionaryObject.Count = 0 Then Return False
        If dictionaryObject.ContainsKey("Ось балки") = True Then
            Dim arrayFindBeams As String(,) = dictionaryObject.Item("Ось балки")
            If IsArray(arrayFindBeams) = True Then
                For i As Integer = 0 To arrayFindBeams.GetUpperBound(1)
                    Dim keyParamBeam As String = arrayFindBeams(1, i)
                    If IsNothing(keyParamBeam) = True Then Continue For
                    If keyParamBeam.Trim.Length > 0 Then
                        Dim userBeam As Beams = Nothing
                        Try
                            userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                        If IsNothing(userBeam) = False Then
                            If userBeam.numberProlet = numberProlet Then
                                ReDim Preserve arrayBeams(4, countBeams)
                                arrayBeams(0, countBeams) = userBeam.numberRow
                                arrayBeams(1, countBeams) = keyParamBeam
                                arrayBeams(2, countBeams) = arrayFindBeams(2, i)
                                arrayBeams(3, countBeams) = arrayFindBeams(3, i)
                                arrayBeams(4, countBeams) = arrayFindBeams(4, i)
                                countBeams += 1
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return countBeams + 1
    End Function
    'функция возвращает все балки мостового сооружения для выбранной опоры
    Public Shared Function FuncFindAllBeamsInPillar(ByVal numberPillar As Integer, ByVal dictionaryObject As Dictionary(Of String, String(,)), ByRef arrayBeams As String(,), Optional ByVal leftOrRight As String = "") As Integer
        FuncFindAllBeamsInPillar = False
        Dim countBeams As Integer = 0
        If IsNothing(dictionaryObject) = True Then Return False
        If dictionaryObject.Count = 0 Then Return False
        If dictionaryObject.ContainsKey("Ось балки") = True Then
            Dim arrayFindBeams As String(,) = dictionaryObject.Item("Ось балки")
            If IsArray(arrayFindBeams) = True Then
                For i As Integer = 0 To arrayFindBeams.GetUpperBound(1)
                    Dim keyParamBeam As String = arrayFindBeams(1, i)
                    Dim userBeam As Beams = Nothing
                    If IsNothing(keyParamBeam) = True Then Continue For
                    If FuncFormatZn.FuncValidateJson(keyParamBeam) = True Then
                        Try
                            userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                        If IsNothing(userBeam) = False Then
                            Dim numColl As Integer = userBeam.numberProlet
                            Dim numRow As Integer = userBeam.numberRow
                            'если номер пролета равен номеру опоры или номер пролета на 1 меньше номера опоры
                            If numColl = numberPillar Or numColl = numberPillar - 1 Then
                                If leftOrRight.Trim.Length > 0 Then
                                    If leftOrRight Like "left" Then
                                        If numRow <= 0 Then
                                            ReDim Preserve arrayBeams(3, countBeams)
                                            arrayBeams(0, countBeams) = numColl
                                            arrayBeams(1, countBeams) = numRow
                                            arrayBeams(2, countBeams) = keyParamBeam
                                            arrayBeams(3, countBeams) = arrayFindBeams(4, i)
                                            countBeams += 1
                                        End If
                                    ElseIf leftOrRight Like "right" Then
                                        If numRow > 0 Then
                                            ReDim Preserve arrayBeams(3, countBeams)
                                            arrayBeams(0, countBeams) = numColl
                                            arrayBeams(1, countBeams) = numRow
                                            arrayBeams(2, countBeams) = keyParamBeam
                                            arrayBeams(3, countBeams) = arrayFindBeams(4, i)
                                            countBeams += 1
                                        End If
                                    End If
                                Else
                                    ReDim Preserve arrayBeams(3, countBeams)
                                    arrayBeams(0, countBeams) = numColl
                                    arrayBeams(1, countBeams) = numRow
                                    arrayBeams(2, countBeams) = keyParamBeam
                                    arrayBeams(3, countBeams) = arrayFindBeams(4, i)
                                    countBeams += 1
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return countBeams + 1
    End Function

    'функция ищет крайние балки мостового сооружения в выбранном пролете
    Public Shared Function FuncFindExtrmBeamsToBridge(ByVal numberProlet As Integer, ByVal dictionaryObject As Dictionary(Of String, String(,)), ByRef arrayBeams As String(,)) As Boolean
        FuncFindExtrmBeamsToBridge = False
        If IsNothing(dictionaryObject) = True Then Return False
        If dictionaryObject.Count = 0 Then Return False
        If dictionaryObject.ContainsKey("Ось балки") = True Then
            Dim arrayFindBeams As String(,) = dictionaryObject.Item("Ось балки")
            If IsArray(arrayFindBeams) = True Then
                Dim maxOffset As Double = -99999
                Dim minOffset As Double = 9999
                ReDim arrayBeams(4, 1)
                For i As Integer = 0 To arrayFindBeams.GetUpperBound(1)
                    Dim keyParamBeam As String = arrayFindBeams(1, i)
                    If IsNothing(keyParamBeam) = True Then Continue For
                    If keyParamBeam.Trim.Length > 0 Then
                        Dim userBeam As Beams = Nothing
                        Try
                            userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                        If IsNothing(userBeam) = False Then
                            If userBeam.numberProlet = numberProlet Then
                                If userBeam.axisStartOffset > maxOffset Then
                                    maxOffset = userBeam.axisStartOffset
                                    arrayBeams(0, 1) = userBeam.numberRow
                                    arrayBeams(1, 1) = keyParamBeam
                                    arrayBeams(2, 1) = arrayFindBeams(2, i)
                                    arrayBeams(3, 1) = arrayFindBeams(3, i)
                                    arrayBeams(4, 1) = arrayFindBeams(4, i)
                                End If
                                If userBeam.axisStartOffset < minOffset Then
                                    minOffset = userBeam.axisStartOffset
                                    arrayBeams(0, 0) = userBeam.numberRow
                                    arrayBeams(1, 0) = keyParamBeam
                                    arrayBeams(2, 0) = arrayFindBeams(2, i)
                                    arrayBeams(3, 0) = arrayFindBeams(3, i)
                                    arrayBeams(4, 0) = arrayFindBeams(4, i)
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End Function
    'функция ищет крайние балки мостового сооружения в выбранном пролете
    Public Shared Function FuncFindExtrmBeamsToBridge(ByRef activDocument As Topomatic.Dwg.Drawing, ByVal align As Alignment, ByVal arrayBeams As String(,), ByRef listLineBeams As List(Of DwgLine)) As Boolean
        FuncFindExtrmBeamsToBridge = False
        If IsNothing(align) = True Then Return False
        If arrayBeams.Length = 0 Then Return False
        Dim maxOffset As Double = -99999
        Dim minOffset As Double = 9999
        Dim leftBeam As DwgLine = Nothing
        Dim rightBeam As DwgLine = Nothing
        For i As Integer = 0 To arrayBeams.GetUpperBound(1)
            Dim hgObject As UInteger = arrayBeams(3, i)
            If hgObject > 0 Then
                Dim lineBeam As DwgLine = Nothing
                Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgObject, lineBeam)
                If boolFindObject = True Then
                    Dim centerLine As Vector2D = New Vector2D((lineBeam.StartPoint.X + lineBeam.EndPoint.X) / 2, (lineBeam.StartPoint.Y + lineBeam.EndPoint.Y) / 2)
                    Dim pk As Double = 0
                    Dim off As Double = 0
                    Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(centerLine, pk, off)
                    If off < 0 Then
                        If off < minOffset Then
                            minOffset = off
                            leftBeam = lineBeam
                        End If
                    ElseIf off > 0 Then
                        If off > maxOffset Then
                            maxOffset = off
                            rightBeam = lineBeam
                        End If
                    End If
                End If
            End If
        Next i
        If IsNothing(leftBeam) = False Then
            Dim boolRevers As Boolean = FuncBridge.FuncCorrectionAxisDirectionBeam(leftBeam, align)
        End If
        If IsNothing(rightBeam) = False Then
            Dim boolRevers As Boolean = FuncBridge.FuncCorrectionAxisDirectionBeam(rightBeam, align)
        End If
        listLineBeams.Add(leftBeam)
        listLineBeams.Add(rightBeam)
        Return True
    End Function
    'функция по номеру пролета возвращает две оси опирания
    Public Shared Function FuncReturnObjectByHgArray(ByVal activDocument As Drawing, ByVal arrayData As String(,), ByRef acEnt As DwgEntity, Optional ByRef keyParam As String = "", Optional idElement As String = "", Optional ByVal indArray As Integer = 0) As Boolean
        If IsNothing(activDocument) = True Then Return False
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(0) > 3 Then
                Dim hg As String = arrayData(4, indArray)
                keyParam = arrayData(1, indArray)
                idElement = arrayData(2, indArray)
                If IsNothing(hg) = False Then
                    If hg.Trim.Length > 0 Then
                        Dim numHg As UInteger = CUInt(hg)
                        Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(numHg, acEnt)
                        If boolFindObject = True Then
                            Return True
                        End If
                    End If
                End If
            End If
        End If
    End Function
    'функция по номеру опоры возврящает опору
    Public Shared Function FuncReturnAxisPillarByNumber(ByVal activDocument As Drawing, ByVal numberPillar As Integer, ByVal dictionaryAllObject As Dictionary(Of String, String(,)), ByRef axisPillar As DwgLine, Optional ByRef userPillar As AxisHorizontalPillars = Nothing) As Boolean
        FuncReturnAxisPillarByNumber = False
        If IsNothing(dictionaryAllObject) = True Then Return False
        If dictionaryAllObject.Count = 0 Then Return False
        If dictionaryAllObject.ContainsKey("Ось опоры") = True Then
            Dim arrayPillars As String(,) = dictionaryAllObject.Item("Ось опоры")
            If IsArray(arrayPillars) = True Then
                For i As Integer = 0 To arrayPillars.GetUpperBound(1)
                    Dim keyParamBeam As String = arrayPillars(1, i)
                    If IsNothing(keyParamBeam) = True Then Continue For
                    If keyParamBeam.Trim.Length > 0 Then
                        Dim userTempPillar As AxisHorizontalPillars = Nothing
                        Try
                            userTempPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisHorizontalPillars)(keyParamBeam)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                        If IsNothing(userTempPillar) = False Then
                            If userTempPillar.number = numberPillar Then
                                Dim hgLine As UInteger = CUInt(arrayPillars(4, i))
                                Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgLine, axisPillar)
                                If boolFindObject = True Then
                                    userPillar = userTempPillar
                                    Return True
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End Function
    'функция возвращает положение шкафной стенки
    Public Shared Function FuncReturnLineCabinetWall(ByVal activDocument As Drawing, ByVal numberPillar As Integer, ByVal dictionaryAllObject As Dictionary(Of String, String(,)), ByVal align As Alignment, ByVal clearence As Double, ByVal outletLeftBeam As Double, ByVal outletRightBeam As Double, ByVal userNozzle As Nozzle, ByRef lineStartCabinetWall As DwgLine) As Boolean
        If IsNothing(dictionaryAllObject) = True Then Return False
        If dictionaryAllObject.Count = 0 Then Return False
        Dim axisLinePillar As DwgLine = Nothing
        Dim boolFindPillar As Boolean = FuncBridge.FuncReturnAxisPillarByNumber(activDocument, numberPillar, dictionaryAllObject, axisLinePillar)
        If boolFindPillar = False Then
            Return False
        End If
        Dim arrayBeams As String(,) = {}
        If numberPillar = 1 Then
            Dim countBeam As Integer = FuncBridge.FuncFindAllBeamsInProlet(1, dictionaryAllObject, arrayBeams)
        Else
            Dim countBeam As Integer = FuncBridge.FuncFindAllBeamsInProlet(numberPillar - 1, dictionaryAllObject, arrayBeams)
        End If
        If arrayBeams.Length = 0 Then
            Return False
        End If
        If arrayBeams.GetUpperBound(1) > 0 Then
            Dim boolSort As Boolean = MathFunction.FuncSortStrArray(arrayBeams, 0)
        End If

        Dim minClearance As Double = -999999
        Dim maxOff As Double = -999999
        Dim minOff As Double = 999999
        Dim widthLeftBeam As Double = 0
        Dim widthRightBeam As Double = 0
        Dim axisLeftBeam As DwgLine = Nothing
        Dim userLeftBeam As Beams = Nothing
        Dim axisRightBeam As DwgLine = Nothing
        Dim userRightBeam As Beams = Nothing
        For i As Integer = 0 To arrayBeams.GetUpperBound(1)
            Dim hgBeam As UInteger = CUInt(arrayBeams(4, i))
            Dim axisLineBeam As DwgLine = Nothing
            Dim boolFindLine1 As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgBeam, axisLineBeam)
            If boolFindLine1 = False Then
                Continue For
            End If
            Dim keyParamBeam1 As String = arrayBeams(1, 0)
            Dim userBeam As Beams = Nothing
            Try
                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam1)
            Catch ex As Newtonsoft.Json.JsonException
            End Try
            'восстанавливаем первую балку
            Dim startListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
            Dim endListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
            Dim boolRestoreBeam As Boolean = FuncBridge.FuncRestoreElementsBeam(axisLineBeam, userBeam, startListPointBeam, endListPointBeam, False, False, True)
            'ищем пересечение оси опирания балки и элемента балки
            Dim boolRez1 As Boolean = False
            Dim ptIntersect1 As Vector2D = MathFunction.FuncFindLineIntersection(endListPointBeam(0).Pos, startListPointBeam(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez1)
            Dim boolRez2 As Boolean = False
            Dim ptIntersect2 As Vector2D = MathFunction.FuncFindLineIntersection(endListPointBeam(1).Pos, startListPointBeam(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez2)
            Dim boolRez3 As Boolean = False
            Dim ptIntersect3 As Vector2D = MathFunction.FuncFindLineIntersection(endListPointBeam(2).Pos, startListPointBeam(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez3)
            Dim boolRez4 As Boolean = False
            Dim ptIntersect4 As Vector2D = MathFunction.FuncFindLineIntersection(endListPointBeam(3).Pos, startListPointBeam(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez4)
            If numberPillar = 1 Then
                Dim distZazor1 As Double = (ptIntersect1 - startListPointBeam(0).Pos).Length
                Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(startListPointBeam(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                Dim intSt As Integer = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect1)
                If intSt = -1 Then
                    distZazor1 = -1 * distZazor1
                End If
                If distZazor1 > minClearance Then
                    minClearance = distZazor1
                End If

                Dim distZazor2 As Double = (ptIntersect2 - startListPointBeam(1).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startListPointBeam(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect2)
                If intSt = -1 Then
                    distZazor2 = -1 * distZazor2
                End If
                If distZazor2 > minClearance Then
                    minClearance = distZazor2
                End If

                Dim distZazor3 As Double = (ptIntersect3 - startListPointBeam(2).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startListPointBeam(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect3)
                If intSt = -1 Then
                    distZazor3 = -1 * distZazor3
                End If
                If distZazor3 > minClearance Then
                    minClearance = distZazor3
                End If

                Dim distZazor4 As Double = (ptIntersect4 - startListPointBeam(3).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startListPointBeam(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect4)
                If intSt = -1 Then
                    distZazor4 = -1 * distZazor4
                End If
                If distZazor4 > minClearance Then
                    minClearance = distZazor4
                End If
            Else
                Dim distZazor1 As Double = (ptIntersect1 - endListPointBeam(0).Pos).Length
                Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(endListPointBeam(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                Dim intSt As Integer = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect1)
                If intSt = 1 Then
                    distZazor1 = -1 * distZazor1
                End If
                If distZazor1 > minClearance Then
                    minClearance = distZazor1
                End If
                Dim distZazor2 As Double = (ptIntersect2 - endListPointBeam(1).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endListPointBeam(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect2)
                If intSt = 1 Then
                    distZazor2 = -1 * distZazor2
                End If
                If distZazor2 > minClearance Then
                    minClearance = distZazor2
                End If
                Dim distZazor3 As Double = (ptIntersect3 - endListPointBeam(2).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endListPointBeam(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect3)
                If intSt = 1 Then
                    distZazor3 = -1 * distZazor3
                End If
                If distZazor3 > minClearance Then
                    minClearance = distZazor3
                End If
                Dim distZazor4 As Double = (ptIntersect4 - endListPointBeam(3).Pos).Length
                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endListPointBeam(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                intSt = MathFunction.funcLeftOrRightPointToLinearObject(axisLinePillar, ptIntersect4)
                If intSt = 1 Then
                    distZazor4 = -1 * distZazor4
                End If
                If distZazor4 > minClearance Then
                    minClearance = distZazor4
                End If
            End If
            Dim middlePoint As Vector2D = MathFunction.funcCalcMiddleCoordByToPoints2d(axisLineBeam.StartPoint.Pos, axisLineBeam.EndPoint.Pos)
            Dim pk As Double = 0
            Dim offPK As Double = 0
            Dim boolPkOff As Boolean = align.Plan.CompoundLine.PosToStaOffset(middlePoint, pk, offPK)
            If boolPkOff = True Then
                If offPK > maxOff Then
                    maxOff = offPK
                    axisRightBeam = axisLineBeam
                    widthRightBeam = userBeam.widthTop
                    userRightBeam = userBeam
                End If
                If offPK < minOff Then
                    minOff = offPK
                    axisLeftBeam = axisLineBeam
                    widthLeftBeam = userBeam.widthTop
                    userLeftBeam = userBeam
                End If
            End If
        Next i

        minClearance += clearence
        '=========================================================================================================================
        'восстанавливаем линию шкафной стенки
        Dim tempLineCabinetWall As DwgLine = New DwgLine()
        If numberPillar = 1 Then
            Dim coolEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            axisLinePillar.Offset(coolEnt, minClearance)
            If coolEnt.Count > 0 Then
                tempLineCabinetWall = coolEnt(0)
            End If
        Else
            Dim coolEnt As List(Of DwgEntity) = New List(Of DwgEntity)
            axisLinePillar.Offset(coolEnt, -1 * minClearance)
            If coolEnt.Count > 0 Then
                tempLineCabinetWall = coolEnt(0)
            End If
        End If
        'находим пересечение крайних балок и линии шкафной стенки
        Dim tempLeftBeam As DwgLine = New DwgLine()
        Dim coolEntBeam As List(Of DwgEntity) = New List(Of DwgEntity)
        axisLeftBeam.Offset(coolEntBeam, -1 * (widthLeftBeam / 2 + outletLeftBeam))
        If coolEntBeam.Count > 0 Then
            tempLeftBeam = coolEntBeam(0)
        End If

        Dim tempRightBeam As DwgLine = New DwgLine()
        coolEntBeam = New List(Of DwgEntity)
        axisRightBeam.Offset(coolEntBeam, widthRightBeam / 2 + outletRightBeam)
        If coolEntBeam.Count > 0 Then
            tempRightBeam = coolEntBeam(0)
        End If

        Dim boolRez As Boolean = False
        Dim ptIntersectLeft As Vector2D = MathFunction.FuncFindLineIntersection(tempLeftBeam.StartPoint.Pos, tempLeftBeam.EndPoint.Pos, tempLineCabinetWall.StartPoint.Pos, tempLineCabinetWall.EndPoint.Pos, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать линию Шкафной стенки.")
            Return False
        End If
        boolRez = False
        Dim ptIntersectRight As Vector2D = MathFunction.FuncFindLineIntersection(tempRightBeam.StartPoint.Pos, tempRightBeam.EndPoint.Pos, tempLineCabinetWall.StartPoint.Pos, tempLineCabinetWall.EndPoint.Pos, boolRez)
        If boolRez = False Then
            MsgBox("Не удалось рассчитать линию Шкафной стенки.")
            Return False
        End If
        If IsNothing(userNozzle) = False Then
            lineStartCabinetWall.StartPoint = New Vector3D(ptIntersectLeft, userNozzle.topElevation)
            lineStartCabinetWall.EndPoint = New Vector3D(ptIntersectRight, userNozzle.topElevation)
        Else
            If numberPillar = 1 Then
                Dim startListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
                Dim endListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
                Dim boolRestoreBeam As Boolean = FuncBridge.FuncRestoreElementsBeam(axisLeftBeam, userLeftBeam, startListPointBeam, endListPointBeam, False, False, True)
                lineStartCabinetWall.StartPoint = New Vector3D(ptIntersectLeft, startListPointBeam(0).Z)
                startListPointBeam = New List(Of Vector3D)
                endListPointBeam = New List(Of Vector3D)
                boolRestoreBeam = FuncBridge.FuncRestoreElementsBeam(axisRightBeam, userRightBeam, startListPointBeam, endListPointBeam, False, False, True)
                lineStartCabinetWall.EndPoint = New Vector3D(ptIntersectRight, startListPointBeam(1).Z)
            Else
                Dim startListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
                Dim endListPointBeam As List(Of Vector3D) = New List(Of Vector3D)
                Dim boolRestoreBeam As Boolean = FuncBridge.FuncRestoreElementsBeam(axisLeftBeam, userLeftBeam, startListPointBeam, endListPointBeam, False, False, True)
                lineStartCabinetWall.StartPoint = New Vector3D(ptIntersectLeft, endListPointBeam(0).Z)
                startListPointBeam = New List(Of Vector3D)
                endListPointBeam = New List(Of Vector3D)
                boolRestoreBeam = FuncBridge.FuncRestoreElementsBeam(axisRightBeam, userRightBeam, startListPointBeam, endListPointBeam, False, False, True)
                lineStartCabinetWall.EndPoint = New Vector3D(ptIntersectRight, endListPointBeam(1).Z)
            End If
        End If
        Return True
    End Function
    'функция ищет в словаре Структурную линию
    Public Shared Function FuncFindStructureLineByIDElement(ByVal acSurface As Surface, ByVal codeLine As Integer, ByVal IDElement As String) As StructureLine
        FuncFindStructureLineByIDElement = Nothing
        If IsNothing(acSurface) = True Then
            Return Nothing
        End If
        Dim structureLineCollection As StructureLines = acSurface.StructureLines

        If structureLineCollection.Count > 0 Then
            For i As Integer = 0 To structureLineCollection.Count - 1
                Dim structureLine As StructureLine = structureLineCollection.Item(i)
                If structureLine.LinearCode = codeLine Then
                    Dim semanticData As SemanticDataSet = structureLine.LinearSemantic
                    If semanticData.Count > 0 Then
                        For j As Integer = 0 To semanticData.Count - 1
                            Dim nameData As SemanticStringNode = semanticData.Root(j)
                            If nameData.Tag Like "ElementID" Then
                                If semanticData.Values(j) Like IDElement Then
                                    Return structureLine
                                End If
                            End If
                        Next j
                    End If
                End If
            Next i
        End If
    End Function
    'функция читает свойства балки
    Public Shared Function FuncReturnPropBeam(ByVal axisLineBearm As DwgLine) As Beams
        FuncReturnPropBeam = Nothing
        Dim keyParam As String = FuncXRecords.FuncReadValueXData(axisLineBearm, "KeyParameters")
        If IsNothing(keyParam) = False Then
            If keyParam.Trim.Length > 4 Then
                Dim tempUserBeam As Beams = Nothing
                Try
                    tempUserBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParam)
                Catch ex As Newtonsoft.Json.JsonException
                End Try
                If IsNothing(tempUserBeam) = False Then
                    Return tempUserBeam
                End If
            End If
        End If
    End Function
    'функция находит балки с минимальной высотой для одного пролета
    Public Shared Function FuncCalculateMinElevationBeams(ByVal activDocument As Drawing, ByVal align As Alignment, ByVal arrayBeams As String(,), ByVal numberPillar As Integer, Optional roundZn As Integer = 3) As Double
        Dim minElevationBottomBeam As Double = 999999 'отметка самой нижней балки в пролете
        If IsArray(arrayBeams) = True Then
            For i As Integer = 0 To arrayBeams.GetUpperBound(1)
                Dim hgAxisBeam As UInteger = CUInt(arrayBeams(3, i))
                If hgAxisBeam > 0 Then
                    Dim axisLineBeam As DwgLine = New DwgLine()
                    Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgAxisBeam, axisLineBeam)
                    If boolFindObject = True Then
                        Dim boolCorr As Boolean = FuncBridge.FuncCorrectionAxisDirectionBeam(axisLineBeam, align)
                        Dim numCollBeam As Integer = CInt(arrayBeams(0, i))
                        Dim elevStartBeam As Double = axisLineBeam.StartPoint.Z
                        Dim elevEndBeam As Double = axisLineBeam.EndPoint.Z
                        'если номер пролета равен номеру опоры, то берем начало балки, иначе конец балки
                        If numberPillar = numCollBeam Then
                            If elevStartBeam < minElevationBottomBeam Then
                                minElevationBottomBeam = elevStartBeam
                            End If
                        Else
                            If elevEndBeam < minElevationBottomBeam Then
                                minElevationBottomBeam = elevEndBeam
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        If minElevationBottomBeam <> 999999 Then
            minElevationBottomBeam = Math.Round(minElevationBottomBeam, roundZn)
            Return minElevationBottomBeam
        End If
    End Function
    'функция возвращает координаты точек опирания балок для выбранной опоры
    Public Shared Function FuncCalculatePrPointBeams(ByVal activDocument As Drawing, ByVal align As Alignment, ByVal arrayBeams As String(,), ByVal numberPillar As Integer) As Dictionary(Of String, Vector3D)
        Dim dictPF As Dictionary(Of String, Vector3D) = New Dictionary(Of String, Vector3D)
        If IsArray(arrayBeams) = True Then
            For i As Integer = 0 To arrayBeams.GetUpperBound(1)
                Dim hgAxisBeam As UInteger = CUInt(arrayBeams(3, i))
                If hgAxisBeam > 0 Then
                    Dim axisLineBeam As DwgLine = New DwgLine()
                    Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgAxisBeam, axisLineBeam)
                    If boolFindObject = True Then
                        Dim boolCorr As Boolean = FuncBridge.FuncCorrectionAxisDirectionBeam(axisLineBeam, align)
                        Dim numCollBeam As Integer = CInt(arrayBeams(0, i))
                        Dim numRowBeam As Integer = CInt(arrayBeams(1, i))
                        Dim elevStartBeam As Double = axisLineBeam.StartPoint.Z
                        Dim elevEndBeam As Double = axisLineBeam.EndPoint.Z
                        'если номер пролета равен номеру опоры, то берем начало балки, иначе конец балки
                        If numberPillar = numCollBeam Then
                            If dictPF.Count = 0 Then
                                dictPF.Add(numRowBeam & "-" & numCollBeam, axisLineBeam.StartPoint)
                            Else
                                If dictPF.ContainsKey(numRowBeam & "-" & numCollBeam) = False Then
                                    dictPF.Add(numRowBeam & "-" & numCollBeam, axisLineBeam.StartPoint)
                                End If
                            End If
                        Else
                            If dictPF.Count = 0 Then
                                dictPF.Add(numRowBeam & "-" & numCollBeam, axisLineBeam.EndPoint)
                            Else
                                If dictPF.ContainsKey(numRowBeam & "-" & numCollBeam) = False Then
                                    dictPF.Add(numRowBeam & "-" & numCollBeam, axisLineBeam.EndPoint)
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
        Return dictPF
    End Function

    'функция находит расстояние от оси опирания балок до линии шкафной стенки
    Public Shared Function FuncCalculatePositionCabinetWall(ByVal activDocument As Drawing, ByVal axisLinePillar As DwgLine, ByVal arrayBeams As String(,), ByVal boolStartBeams As Boolean) As Double
        'находим минимальный зазор между балками и шкафной стенкой и строим линию шкафной стенки и находит точки по контуру насадки
        Dim minClearance As Double = -999999
        If IsArray(arrayBeams) = False Then Return 0
        If arrayBeams.Length = 0 Then Return 0
        For i As Integer = 0 To arrayBeams.GetUpperBound(1)
            Dim hgAxisBeam As UInteger = CUInt(arrayBeams(3, i))
            If hgAxisBeam > 0 Then
                Dim axisLineBeam As DwgLine = New DwgLine()
                Dim boolFindObject As Boolean = activDocument.ActiveSpace.Entities.TryGetObject(hgAxisBeam, axisLineBeam)
                If boolFindObject = True Then
                    Dim userBeam As Beams = Nothing
                    Dim keyParam As String = arrayBeams(2, i)
                    If IsNothing(keyParam) = True Then
                        Continue For
                    End If
                    If keyParam.Trim.Length < 4 Then
                        Continue For
                    End If
                    Try
                        userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParam)
                    Catch ex As Newtonsoft.Json.JsonException
                    End Try
                    If IsNothing(userBeam) = False Then
                        Dim startPointElements As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
                        Dim endPointElements As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
                        Dim boolRestoreElements As Boolean = FuncBridge.FuncRestoreElementsBeam(axisLineBeam, userBeam, startPointElements, endPointElements, False, False, True)
                        'ищем пересечения с осью первой или последней опоры
                        If boolRestoreElements = True Then
                            Dim boolRez1 As Boolean = False
                            Dim ptIntersect1 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(0).Pos, startPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez1)
                            Dim boolRez2 As Boolean = False
                            Dim ptIntersect2 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(1).Pos, startPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez2)
                            Dim boolRez3 As Boolean = False
                            Dim ptIntersect3 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(2).Pos, startPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez3)
                            Dim boolRez4 As Boolean = False
                            Dim ptIntersect4 As Vector2D = MathFunction.FuncFindLineIntersection(endPointElements(3).Pos, startPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, boolRez4)
                            If boolStartBeams = True Then
                                Dim distZazor1 As Double = (ptIntersect1 - startPointElements(0).Pos).Length
                                Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(startPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                                If distZazor1 > minClearance Then
                                    minClearance = distZazor1
                                End If
                                Dim distZazor2 As Double = (ptIntersect2 - startPointElements(1).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                                If distZazor2 > minClearance Then
                                    minClearance = distZazor2
                                End If
                                Dim distZazor3 As Double = (ptIntersect3 - startPointElements(2).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                                If distZazor3 > minClearance Then
                                    minClearance = distZazor3
                                End If
                                Dim distZazor4 As Double = (ptIntersect4 - startPointElements(3).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(startPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                                If distZazor4 > minClearance Then
                                    minClearance = distZazor4
                                End If
                            Else
                                Dim distZazor1 As Double = (ptIntersect1 - endPointElements(0).Pos).Length
                                Dim pointIntersect As Vector2D = MathFunction.FuncFindPerpendicularPoint(endPointElements(0).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor1)
                                If distZazor1 > minClearance Then
                                    minClearance = distZazor1
                                End If
                                Dim distZazor2 As Double = (ptIntersect2 - endPointElements(1).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(1).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor2)
                                If distZazor2 > minClearance Then
                                    minClearance = distZazor2
                                End If
                                Dim distZazor3 As Double = (ptIntersect3 - endPointElements(2).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(2).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor3)
                                If distZazor3 > minClearance Then
                                    minClearance = distZazor3
                                End If
                                Dim distZazor4 As Double = (ptIntersect4 - endPointElements(3).Pos).Length
                                pointIntersect = MathFunction.FuncFindPerpendicularPoint(endPointElements(3).Pos, axisLinePillar.StartPoint.Pos, axisLinePillar.EndPoint.Pos, distZazor4)
                                If distZazor4 > minClearance Then
                                    minClearance = distZazor4
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        Return minClearance
    End Function


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функции оформления
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncDrawBoundaryBridge(ByRef activDocument As Topomatic.Dwg.Drawing, ByVal dimleftBorderStructure As Double, ByVal dimRightBorderStructure As Double, ByVal startAxisPillar As DwgLine, ByVal endAxisPillar As DwgLine, ByVal elementsBeams As Dictionary(Of Integer, DwgLine()), ByVal align As Alignment, Optional ByRef boundStructure As DwgPolyline = Nothing, Optional ByRef alignStructure As DwgPolyline = Nothing) As Boolean
        FuncDrawBoundaryBridge = False
        If dimleftBorderStructure = 0 Then Return False
        If dimRightBorderStructure = 0 Then Return False
        If IsNothing(startAxisPillar) = True Then Return False
        If IsNothing(endAxisPillar) = True Then Return False
        '===============================================================================================================================
        Dim acPoly3dAlign As Polyline3D = New Polyline3D
        align.Plan.CompoundLine.ToPolyLine(acPoly3dAlign)

        Dim axisPline0 As DwgPolyline = Nothing 'полилиния из трассы  в прямом направлении
        Dim ArrayCoord As Double(,) = Nothing
        Dim countarrayPos As Integer = 0
        For i As Integer = 0 To acPoly3dAlign.Count - 1
            Dim pos As BugleVector3D = acPoly3dAlign.Item(i)
            ReDim Preserve ArrayCoord(3, countarrayPos)
            ArrayCoord(0, countarrayPos) = pos.Vertex.X
            ArrayCoord(1, countarrayPos) = pos.Vertex.Y
            ArrayCoord(2, countarrayPos) = pos.Bugle
            countarrayPos += 1
        Next
        If IsArray(ArrayCoord) = True Then
            If ArrayCoord.GetUpperBound(1) > 0 Then
                axisPline0 = RoburFunc.FuncDrawPolylineToArrayCoord(activDocument, ArrayCoord)
            End If
        End If
        'находим края сооружения (распаралеливаем ось на величину границы сооружения
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.FuncOffsetAlignment(activDocument, align, -1 * dimleftBorderStructure)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)

        'обратное направление
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.FuncOffsetAlignment(activDocument, align, -1 * dimRightBorderStructure, True)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)

        Dim pkStartDirect As Double = -999999999
        Dim pkEndDirect As Double = -999999999
        Dim pkStartReverse As Double = -999999999
        Dim pkEndReverse As Double = -999999999

        Dim angle As Double = startAxisPillar.Rotation
        Dim reverseAngle As Double = angle + Math.PI
        If reverseAngle >= Math.PI * 2 Then
            reverseAngle -= Math.PI * 2
        End If

        Dim arrayElements As DwgLine() = elementsBeams.Item(1)
        Dim pkStart As Double = 9999999999
        Dim off As Double = 0
        For Each acLine As DwgLine In arrayElements
            Dim firstPoint As Vector2D = acLine.StartPoint.Pos
            Dim secondPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(firstPoint, angle, startAxisPillar.Length, 3)
            Dim pointIntersectCollection0 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, firstPoint, secondPoint)
            If pointIntersectCollection0.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection0.ElementAt(0) 'осевая точка
                Dim pk As Double = 0
                Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(intersectPoint2D, pk, off)
                If boolPk = True Then
                    If pk < pkStart Then
                        pkStart = pk
                    End If
                End If
            End If
        Next
        If pkStart = 9999999999 Then
            Exit Function
        End If
        Dim startAlignPoint As Vector2D = acPoly3dAlign.StaOffsetToPos(pkStart, 0)
        Dim ptStartLeft As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(startAlignPoint, reverseAngle, dimleftBorderStructure * 2)
        Dim ptStartRight As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(startAlignPoint, angle, dimRightBorderStructure * 2)
        Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, startAlignPoint, ptStartLeft)
        If pointIntersectCollection1.Count > 0 Then
            Dim intersectPoint2D As Vector2D = pointIntersectCollection1.ElementAt(0) '
            Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkStartDirect, off)
        End If
        Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, startAlignPoint, ptStartRight)
        If pointIntersectCollection2.Count > 0 Then
            Dim intersectPoint2D As Vector2D = pointIntersectCollection2.ElementAt(0) '
            Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkStartReverse, off)
        End If

        angle = endAxisPillar.Rotation
        reverseAngle = angle + Math.PI
        If reverseAngle >= Math.PI * 2 Then
            reverseAngle -= Math.PI * 2
        End If

        If elementsBeams.Count = 1 Then
            arrayElements = elementsBeams.Item(1)
        Else
            arrayElements = elementsBeams.Item(2)
        End If
        Dim pkEnd As Double = -9999999999
        off = 0
        For Each acLine As DwgLine In arrayElements
            Dim firstPoint As Vector2D = acLine.EndPoint.Pos
            Dim secondPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(firstPoint, angle, endAxisPillar.Length, 3)
            Dim pointIntersectCollection0 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, firstPoint, secondPoint)
            If pointIntersectCollection0.Count > 0 Then
                Dim intersectPoint2D As Vector2D = pointIntersectCollection0.ElementAt(0) 'осевая точка
                Dim pk As Double = 0
                Dim boolPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(intersectPoint2D, pk, off)
                If boolPk = True Then
                    If pk > pkEnd Then
                        pkEnd = pk
                    End If
                End If
            End If
        Next
        If pkEnd = -9999999999 Then
            Exit Function
        End If
        Dim endAlignPoint As Vector2D = acPoly3dAlign.StaOffsetToPos(pkEnd, 0)
        Dim ptEndLeft As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(endAlignPoint, reverseAngle, dimleftBorderStructure * 2)
        Dim ptEndRight As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(endAlignPoint, angle, dimRightBorderStructure * 2)
        Dim pointIntersectCollection3 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, endAlignPoint, ptEndLeft)
        If pointIntersectCollection3.Count > 0 Then
            Dim intersectPoint2D As Vector2D = pointIntersectCollection3.ElementAt(0) '
            Dim boolPKD As Boolean = axisPline3DDirect.PosToStaOffset(intersectPoint2D, pkEndDirect, off)
        End If
        Dim pointIntersectCollection4 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, endAlignPoint, ptEndRight)
        If pointIntersectCollection4.Count > 0 Then
            Dim intersectPoint2D As Vector2D = pointIntersectCollection4.ElementAt(0) '
            Dim boolPKD As Boolean = axisPline3DReverse.PosToStaOffset(intersectPoint2D, pkEndReverse, off)
        End If

        If pkStartDirect <> -999999999 And pkStartReverse <> -999999999 And pkEndDirect <> -999999999 And pkEndReverse <> -999999999 Then
            Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
            For i As Integer = 0 To axisPlineDirect.Count - 1
                pline2dCurv.Add(axisPlineDirect.Item(i))
            Next i
            Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(pkStartDirect, pkEndDirect)
            Dim pline2dCurv1 As Polyline2DCurve = New Polyline2DCurve()
            If IsArray(arrayPlineCurve) = True Then
                For i As Integer = 0 To arrayPlineCurve.Length - 1
                    Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                    If Math.Abs(plineCurv.Length - (pkEndDirect - pkStartDirect)) <= 0.01 Then
                        pline2dCurv1 = plineCurv
                    End If
                Next i
            End If
            pline2dCurv = New Polyline2DCurve()
            For i As Integer = 0 To axisPlineReverse.Count - 1
                pline2dCurv.Add(axisPlineReverse.Item(i))
            Next i
            arrayPlineCurve = pline2dCurv.Break(pkEndReverse, pkStartReverse)
            Dim pline2dCurv2 As Polyline2DCurve = New Polyline2DCurve()
            If IsArray(arrayPlineCurve) = True Then
                For i As Integer = 0 To arrayPlineCurve.Length - 1
                    Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                    If Math.Abs(plineCurv.Length - (pkStartReverse - pkEndReverse)) <= 0.01 Then
                        pline2dCurv2 = plineCurv
                    End If
                Next i
            End If

            If pline2dCurv2.Length > 0 And pline2dCurv1.Length > 0 Then
                If IsNothing(boundStructure) = True Then
                    boundStructure = New DwgPolyline()
                    activDocument.ActiveSpace.Entities.Add(boundStructure)
                ElseIf boundStructure.Length = 0 Then
                    activDocument.ActiveSpace.Entities.Add(boundStructure)
                Else
                    boundStructure.Clear()
                End If

                For i As Integer = 0 To pline2dCurv1.Count - 1
                    boundStructure.Add(pline2dCurv1.Item(i))
                Next i
                Dim count As Integer = pline2dCurv2.Count - 1
                For i As Integer = 0 To pline2dCurv2.Count - 1
                    Dim bulg As Double = 0
                    If i <> 0 Then
                        bulg = pline2dCurv2.Item(i - 1).Bugle
                    End If
                    boundStructure.Add(pline2dCurv2.Item(i))
                Next i
                boundStructure.Closed = True
            End If
            'ось трассы
            pline2dCurv = New Polyline2DCurve()
            For i As Integer = 0 To axisPline0.Count - 1
                pline2dCurv.Add(axisPline0.Item(i))
            Next i
            arrayPlineCurve = pline2dCurv.Break(pkStart, pkEnd)
            Dim pline2dCurv3 As Polyline2DCurve = New Polyline2DCurve()
            If IsArray(arrayPlineCurve) = True Then
                For i As Integer = 0 To arrayPlineCurve.Length - 1
                    Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                    If Math.Abs(plineCurv.Length - (pkEnd - pkStart)) <= 0.01 Then
                        pline2dCurv3 = plineCurv
                    End If
                Next i
            End If
            If pline2dCurv3.Length > 0 Then
                If IsNothing(alignStructure) = True Then
                    alignStructure = New DwgPolyline()
                    activDocument.ActiveSpace.Entities.Add(alignStructure)
                ElseIf alignStructure.Length = 0 Then
                    If activDocument.ActiveSpace.Entities.Contains(alignStructure) = False Then
                        activDocument.ActiveSpace.Entities.Add(alignStructure)
                    End If
                Else
                    alignStructure.Clear()
                End If

                For i As Integer = 0 To pline2dCurv3.Count - 1
                    alignStructure.Add(pline2dCurv3.Item(i))
                Next i
            End If

            If IsNothing(axisPlineDirect) = False Then
                activDocument.ActiveSpace.Entities.Remove(axisPlineDirect)
            End If
            If IsNothing(axisPlineReverse) = False Then
                activDocument.ActiveSpace.Entities.Remove(axisPlineReverse)
            End If
            If IsNothing(axisPline0) = False Then
                activDocument.ActiveSpace.Entities.Remove(axisPline0)
            End If
            Return True
        Else
            If activDocument.ActiveSpace.Entities.Contains(axisPlineDirect) = True Then
                activDocument.ActiveSpace.Entities.Remove(axisPlineDirect)
            End If
            If activDocument.ActiveSpace.Entities.Contains(axisPlineReverse) = True Then
                activDocument.ActiveSpace.Entities.Remove(axisPlineReverse)
            End If
            For i As Integer = 0 To axisPline0.Count - 1
                alignStructure.Add(axisPline0.Item(i))
            Next i
            If activDocument.ActiveSpace.Entities.Contains(axisPline0) = True Then
                activDocument.ActiveSpace.Entities.Remove(axisPlineDirect)
            End If
            Return False
        End If
    End Function
    'коррекция положения балки на допуск по зазору (функция работает по всем балкам из диалогового окна расстановки балок)
    Public Shared Function FuncCorrBalka(ByRef lineShortBearm As DwgLine, ByRef pointEndSectionColl As List(Of Vector3D), ByVal heightBearm As Double, ByVal widthDownBearm As Double, ByVal widthUpBearm As Double, ByVal startDistToPointPr As Double, ByVal endDistToPointPr As Double, ByVal zazor As Double, ByRef pointStartSectionColl As List(Of Vector3D), ByVal axisIPline3D As IPolyline3D, ByVal surf As Surface, ByVal dEarth As Double, ByVal userAlign As Alignment) As Boolean
        FuncCorrBalka = False
        If IsNothing(lineShortBearm) = True Then Exit Function
        Try
            'проверка балки на направление по ходу пикетажа
            Dim PkStartLineShortBearm As Double = 0
            Dim PkEndLineShortBearm As Double = 0
            Dim off As Double = 0
            Dim boolpk1start As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(lineShortBearm.StartPoint.Pos, PkStartLineShortBearm, off)
            Dim boolpk1end As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(lineShortBearm.EndPoint.Pos, PkEndLineShortBearm, off)
            '1. Удлинняем балку (временно, на величину участков опирания)
            Dim lenghtShortBearm As Double = lineShortBearm.Length
            Dim lineBearm As DwgLine = lineShortBearm.Clone() 'работаем с копией
            Dim boolExtBearm As Boolean = False

            If boolpk1end = True And boolpk1start = True Then
                If PkStartLineShortBearm > PkEndLineShortBearm Then
                    boolExtBearm = FuncExtendBearm(lineBearm, endDistToPointPr, startDistToPointPr)
                Else
                    boolExtBearm = FuncExtendBearm(lineBearm, startDistToPointPr, endDistToPointPr)
                End If
            Else
                boolExtBearm = FuncExtendBearm(lineBearm, startDistToPointPr, endDistToPointPr)
            End If
            If boolExtBearm = False Then Exit Function
            '2 делаем смещение балки вверх
            '2.1 находим угол наклона балки
            Dim b As Double = lineBearm.EndPoint.Z - lineBearm.StartPoint.Z
            Dim c As Double = lineBearm.Length
            Dim i As Double = Math.Asin(b / c)
            Dim wAngle As Double = (i * 180) / Math.PI
            Dim insCoord As Vector2D = New Vector2D(0, 0)
            Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, heightBearm)
            'получаем новые координаты верха балки
            Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.StartPoint.X, lineBearm.StartPoint.Y), lineBearm.Rotation, deltaXZ.X)
            Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.EndPoint.X, lineBearm.EndPoint.Y), lineBearm.Rotation, deltaXZ.X)
            Dim lContr As Double = MathFunction.funcCalcDistanceByToPoints2d(startCoordUpBearm, endCoordUpBearm) 'первая проверка
            'создаем верх балки
            Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineBearm.StartPoint.Z + deltaXZ.Y)
            Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineBearm.EndPoint.Z + deltaXZ.Y)
            Dim lContr2 As Double = (startPointUpBalkaPr - endPointUpBalkaPr).Length 'вторая проверка

            Dim lineUpBalka As DwgLine = New DwgLine()
            lineUpBalka.StartPoint = startPointUpBalkaPr
            lineUpBalka.EndPoint = endPointUpBalkaPr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '3. приводим балки к одному уровню (временно)
            If pointEndSectionColl.Count = 4 Then
                'верх балки предыдущей
                Dim upLeftPrev As Vector3D = pointEndSectionColl.Item(2)
                Dim upRightPrev As Vector3D = pointEndSectionColl.Item(3)
                'низ балки предыдущей
                Dim downLeftPrev As Vector3D = pointEndSectionColl.Item(0)
                Dim downRightPrev As Vector3D = pointEndSectionColl.Item(1)
                'высота по низу предыдущей балки
                Dim Hprev As Double = (downLeftPrev.Z + downRightPrev.Z) / 2
                'высота по низу балки определяемой
                Dim Hbm As Double = lineBearm.StartPoint.Z
                'дельта по высоте
                Dim deltaH As Double = Hprev - Hbm
                'временные координаты для предыдущей балки если она ниже чем расчетная
                Dim downLeftPrev2 As Vector3D = pointEndSectionColl.Item(0)
                Dim downRightPrev2 As Vector3D = pointEndSectionColl.Item(1)
                If deltaH < 0 Then
                    'предудущая балка ниже чем расчетная (приводим ее уровню расчетной бплки
                    downLeftPrev2 = MathFunction.FuncCalcPointInLine(downLeftPrev, upLeftPrev, Math.Abs(deltaH))
                    downRightPrev2 = MathFunction.FuncCalcPointInLine(downRightPrev, upRightPrev, Math.Abs(deltaH))
                Else 'расчетная балка ниже чем предыдущая (подтягиваем расчетную балку к предыдущей)
                    Dim tempPointB1 As Vector3D = MathFunction.FuncCalcPointInLine(lineBearm.StartPoint, lineUpBalka.StartPoint, deltaH)
                    Dim tempPointB2 As Vector3D = MathFunction.FuncCalcPointInLine(lineBearm.EndPoint, lineUpBalka.EndPoint, deltaH)
                    lineBearm.StartPoint = tempPointB1
                    lineBearm.EndPoint = tempPointB2
                End If
                'делаем повторнуб проверку
                'высота по низу предыдущей балки
                Hprev = (downLeftPrev2.Z + downRightPrev2.Z) / 2
                'высота по низу балки определяемой
                Hbm = lineBearm.StartPoint.Z
                'дельта по высоте
                deltaH = Hprev - Hbm
                If Math.Abs(deltaH) > 0.05 Then
                    MsgBox("Не удалось привести балки к одному уровню!!!")
                End If
            End If

            '4. распаралеливаем линии расчетной балки
            Dim deltaZazor As Double = 0
            Dim pointStart1 As Vector3D = Nothing
            Dim pointStart2 As Vector3D = Nothing
            Dim pointStart3 As Vector3D = Nothing
            Dim pointStart4 As Vector3D = Nothing
            Dim pointEnd1 As Vector3D = Nothing
            Dim pointEnd2 As Vector3D = Nothing
            Dim pointEnd3 As Vector3D = Nothing
            Dim pointEnd4 As Vector3D = Nothing
            Dim H As Double = lineBearm.EndPoint.Z - lineBearm.StartPoint.Z
            'считаем зазор по низу балки
            '====================================================================================================================
            'смещение низ право
            Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
            Dim offsetElement As Double = widthDownBearm / 2
Line1:
            lineBearm.Offset(dbCollection1, offsetElement)
            If dbCollection1.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection1.Item(0)
                'удлинняем линию для нахождения точки
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count > 3 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(1).Pos, pointEndSectionColl.Item(0).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(1).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(0).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection1 = New List(Of DwgEntity)
                                GoTo Line1
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection1 = New List(Of DwgEntity)
                                GoTo Line1
                            End If
                        End If
                        'ищем угол
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        If dAngle < 2 Then
                            Dim tempd As Double = dist + zazor
                            If tempd > deltaZazor Then
                                deltaZazor = tempd
                            End If
                        Else
                            If dist < zazor Then
                                Dim tempd As Double = zazor - dist
                                If tempd > deltaZazor Then
                                    deltaZazor = tempd
                                End If
                            End If
                        End If
                    End If
                End If
                pointStart1 = tempLine.StartPoint
                pointEnd1 = tempLine.EndPoint
            End If
            '====================================================================================================================
            'низ лево'
            Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = -1 * widthDownBearm / 2
Line2:
            lineBearm.Offset(dbCollection2, offsetElement)
            If dbCollection2.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection2.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(0).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(1).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection2 = New List(Of DwgEntity)
                                GoTo Line2
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection2 = New List(Of DwgEntity)
                                GoTo Line2
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        If dAngle < 2 Then
                            Dim tempd As Double = dist + zazor
                            If tempd > deltaZazor Then
                                deltaZazor = tempd
                            End If
                        Else
                            If dist < zazor Then
                                Dim tempd As Double = zazor - dist
                                If tempd > deltaZazor Then
                                    deltaZazor = tempd
                                End If
                            End If
                        End If
                    End If
                End If
                pointStart2 = tempLine.StartPoint
                pointEnd2 = tempLine.EndPoint
            End If
            'считаем зазор по верху балки
            '====================================================================================================================
            'верх право
            Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = widthUpBearm / 2
line3:
            lineUpBalka.Offset(dbCollection3, offsetElement)
            If dbCollection3.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection3.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(3).Pos, pointEndSectionColl.Item(2).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(3).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(2).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection3 = New List(Of DwgEntity)
                                GoTo line3
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection3 = New List(Of DwgEntity)
                                GoTo line3
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        If dAngle < 2 Then
                            Dim tempd As Double = dist + zazor
                            If tempd > deltaZazor Then
                                deltaZazor = tempd
                            End If
                        Else
                            If dist < zazor Then
                                Dim tempd As Double = zazor - dist
                                If tempd > deltaZazor Then
                                    deltaZazor = tempd
                                End If
                            End If
                        End If
                    End If
                End If
                pointStart3 = tempLine.StartPoint
                pointEnd3 = tempLine.EndPoint
            End If
            '====================================================================================================================
            'верх лево
            Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = -1 * widthUpBearm / 2
line4:
            lineUpBalka.Offset(dbCollection4, offsetElement)
            If dbCollection4.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection4.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(2).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(3).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection4 = New List(Of DwgEntity)
                                GoTo line4
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection4 = New List(Of DwgEntity)
                                GoTo line4
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        If dAngle < 2 Then
                            Dim tempd As Double = dist + zazor
                            If tempd > deltaZazor Then
                                deltaZazor = tempd
                            End If
                        Else
                            If dist < zazor Then
                                Dim tempd As Double = zazor - dist
                                If tempd > deltaZazor Then
                                    deltaZazor = tempd
                                End If
                            End If
                        End If
                    End If
                End If
                pointStart4 = tempLine.StartPoint
                pointEnd4 = tempLine.EndPoint
            End If
            'If Math.Abs(deltaZazor) > 0.001 Then
            '5. делаем коррекцию положения расчетной балки
            'поправка в линию
            If Math.Abs(deltaZazor) > 0.001 Then
                Dim boolMoveBearm As Boolean = FuncMoveBearm(lineShortBearm, deltaZazor)
            End If
            'садим точно на ось
            'первую точку
            Try
                Dim dist1 As Double = -1
                Dim offDist As Double = -1
                Dim boolSt1 As Boolean = PolylineExtentions.PosToStaOffset(axisIPline3D, lineShortBearm.StartPoint.Pos, dist1, offDist)
                If boolSt1 = True Then
                    Dim newPtStart As Vector2D = PolylineExtentions.StaOffsetToPos(axisIPline3D, dist1, 0)
                    Dim newElev1 As Double = surf.GetElevation(newPtStart) - heightBearm - dEarth
                    lineShortBearm.StartPoint = New Vector3D(newPtStart.X, newPtStart.Y, newElev1)
                End If
                'последнюю точку
                Dim boolSt2 As Boolean = PolylineExtentions.PosToStaOffset(axisIPline3D, lineShortBearm.EndPoint.Pos, dist1, offDist)
                If boolSt2 = True Then
                    Dim newPtEnd As Vector2D = PolylineExtentions.StaOffsetToPos(axisIPline3D, dist1, 0)
                    Dim newElev2 As Double = surf.GetElevation(newPtEnd) - heightBearm - dEarth
                    lineShortBearm.EndPoint = New Vector3D(newPtEnd.X, newPtEnd.Y, newElev2)
                End If

                Dim dhLeftStart As Double = 0
                Dim dhRightStart As Double = 0
                Dim dhLeftEnd As Double = 0
                Dim dhRightEnd As Double = 0
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'параллелим ось балки влево
                Dim dbCollection5 As List(Of DwgEntity) = New List(Of DwgEntity)
                lineShortBearm.Offset(dbCollection5, -1 * widthUpBearm / 2)
                If dbCollection5.Count = 1 Then
                    Dim tempLine As DwgLine = dbCollection5.Item(0)
                    'определяем высоты ее концов
                    Try
                        Dim newElev1 As Double = surf.GetElevation(tempLine.StartPoint.Pos)
                        Dim newElev2 As Double = surf.GetElevation(tempLine.EndPoint.Pos)
                        dhLeftStart = tempLine.StartPoint.Z + heightBearm + dEarth - newElev1
                        dhLeftEnd = tempLine.EndPoint.Z + heightBearm + dEarth - newElev2
                    Catch ex As System.Exception
                    End Try
                End If

                Dim dbCollection6 As List(Of DwgEntity) = New List(Of DwgEntity)
                lineShortBearm.Offset(dbCollection6, widthUpBearm / 2)
                If dbCollection6.Count = 1 Then
                    Dim tempLine As DwgLine = dbCollection6.Item(0)
                    Dim newElev1 As Double = surf.GetElevation(tempLine.StartPoint.Pos)
                    Dim newElev2 As Double = surf.GetElevation(tempLine.EndPoint.Pos)
                    dhRightStart = tempLine.StartPoint.Z + heightBearm + dEarth - newElev1
                    dhRightEnd = tempLine.EndPoint.Z + heightBearm + dEarth - newElev2
                End If

                If dhLeftStart > 0 And dhRightStart <= 0 Then
                    lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z - dhLeftStart)
                ElseIf dhLeftStart <= 0 And dhRightStart > 0 Then
                    lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z - dhRightStart)
                ElseIf dhLeftStart > 0 And dhRightStart > 0 Then
                    If dhLeftStart > dhRightStart Then
                        lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z - dhLeftStart)
                    Else
                        lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z - dhRightStart)
                    End If
                ElseIf dhLeftStart < 0 And dhRightStart < 0 Then
                    If dhLeftStart > dhRightStart Then
                        lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z + dhRightStart)
                    Else
                        lineShortBearm.StartPoint = New Vector3D(lineShortBearm.StartPoint.Pos, lineShortBearm.StartPoint.Z + dhLeftStart)
                    End If
                End If

                If dhLeftEnd > 0 And dhRightEnd <= 0 Then
                    lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z - dhLeftEnd)
                ElseIf dhLeftEnd <= 0 And dhRightEnd > 0 Then
                    lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z - dhRightEnd)
                ElseIf dhLeftEnd > 0 And dhRightEnd > 0 Then
                    If dhLeftEnd > dhRightEnd Then
                        lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z - dhLeftEnd)
                    Else
                        lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z - dhRightEnd)
                    End If
                ElseIf dhLeftEnd < 0 And dhRightEnd < 0 Then
                    If dhLeftEnd > dhRightEnd Then
                        lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z + dhRightEnd)
                    Else
                        lineShortBearm.EndPoint = New Vector3D(lineShortBearm.EndPoint.Pos, lineShortBearm.EndPoint.Z + dhLeftEnd)
                    End If
                End If
            Catch ex As System.Exception
                Return False
            End Try
            'смотрим длину
            Dim tempL As Double = lineShortBearm.Length
            Dim dLenght As Double = tempL - lenghtShortBearm
            If Math.Abs(dLenght) >= 0.001 Then
                'делаем коррекцию длины
                Dim boolExtend As Boolean = FuncExtendBearm(lineShortBearm, 0, dLenght)
                If boolExtend = False Then
                    Return False
                End If
            Else
                pointEndSectionColl = New List(Of Vector3D)
                pointStartSectionColl = New List(Of Vector3D)
                Dim boolRestoreBram As Boolean = FuncRestoreCoordinatesBalka(lineShortBearm, heightBearm, widthDownBearm, widthUpBearm, startDistToPointPr, endDistToPointPr, pointStartSectionColl, pointEndSectionColl)
                Return True
            End If
        Catch ex As System.Exception
        End Try
    End Function
    'функция корректирует балку при изменении ее высоты
    Public Shared Function FuncCorrBalkaElevation(ByVal ActivDocument As Topomatic.Dwg.Drawing, ByRef lineShortBeam As DwgLine, ByVal deltaElevation As Double, ByVal align As Alignment, ByVal surf As Surface) As Boolean
        FuncCorrBalkaElevation = False
        If IsNothing(lineShortBeam) = True Then
            Return False
        ElseIf lineShortBeam.Length = 0 Then
            Return False
        End If
        If IsNothing(surf) = True Then Return False
        Dim keyParam As String = FuncXRecords.FuncReadValueXData(lineShortBeam, "KeyParameters", "PROJECT_BRIDGE")
        If keyParam.Trim.Length = 0 Then
            Return False
        End If
        Try
            Dim userBeam As Beams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParam)
            Dim fullLenghtBeam As Double = userBeam.fullLenght
            Dim a As Double = userBeam.a
            Dim b As Double = userBeam.b
            Dim shortLenghtBeam As Double = fullLenghtBeam - a - b
            Dim axisOffset As Double = userBeam.axisStartOffset
            Dim axisPline As DwgPolyline = FuncAlignment.FuncOffsetAlignment(ActivDocument, align, axisOffset)
            Dim axisPline3D As IPolyline3D = New Polyline3D()
            axisPline.GetPolyline(axisPline3D)
            Dim startPointBeam As Vector3D = lineShortBeam.StartPoint
            Dim endPointBeam As Vector3D = lineShortBeam.EndPoint
            For i As Integer = 0 To 100
                'восстанавливаем балку
                Dim startSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim endSectionPoint3d As List(Of Topomatic.Cad.Foundation.Vector3D) = New List(Of Topomatic.Cad.Foundation.Vector3D)
                Dim restoreBeam As Boolean = FuncBridge.FuncRestoreCoordinatesBalka(lineShortBeam, userBeam.height, userBeam.widthBottom, userBeam.widthTop, userBeam.a, userBeam.b, startSectionPoint3d, endSectionPoint3d)
                Try
                    'корректируем балку по высоте
                    Dim elev1 As Double = surf.GetElevation(startSectionPoint3d.Item(2).Pos)
                    Dim elevTopStart1 As Double = elev1 - (deltaElevation + startSectionPoint3d.Item(2).Z)

                    Dim elev2 As Double = surf.GetElevation(startSectionPoint3d.Item(3).Pos)
                    Dim elevTopStart2 As Double = elev2 - (deltaElevation + startSectionPoint3d.Item(3).Z)

                    Dim elev3 As Double = surf.GetElevation(endSectionPoint3d.Item(2).Pos)
                    Dim elevTopStart3 As Double = elev3 - (deltaElevation + endSectionPoint3d.Item(2).Z)

                    Dim elev4 As Double = surf.GetElevation(endSectionPoint3d.Item(3).Pos)
                    Dim elevTopStart4 As Double = elev4 - (deltaElevation + endSectionPoint3d.Item(3).Z)

                    If elevTopStart1 < 0 And elevTopStart2 < 0 Then
                        If Math.Abs(elevTopStart1) > Math.Abs(elevTopStart2) Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        End If
                    ElseIf elevTopStart1 > 0 And elevTopStart2 > 0 Then
                        If elevTopStart1 > elevTopStart2 Then
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                        Else
                            startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                        End If
                    ElseIf elevTopStart1 < 0 And elevTopStart2 > 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart1)
                    ElseIf elevTopStart1 > 0 And elevTopStart2 < 0 Then
                        startPointBeam = New Topomatic.Cad.Foundation.Vector3D(startPointBeam.X, startPointBeam.Y, startPointBeam.Z + elevTopStart2)
                    End If

                    If elevTopStart3 < 0 And elevTopStart4 < 0 Then
                        If Math.Abs(elevTopStart3) > Math.Abs(elevTopStart4) Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        End If
                    ElseIf elevTopStart3 > 0 And elevTopStart4 > 0 Then
                        If elevTopStart3 > elevTopStart4 Then
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                        Else
                            endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                        End If
                    ElseIf elevTopStart3 < 0 And elevTopStart4 > 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart3)
                    ElseIf elevTopStart3 > 0 And elevTopStart4 < 0 Then
                        endPointBeam = New Topomatic.Cad.Foundation.Vector3D(endPointBeam.X, endPointBeam.Y, endPointBeam.Z + elevTopStart4)
                    End If
                    lineShortBeam.StartPoint = startPointBeam
                    lineShortBeam.EndPoint = endPointBeam

                    'делаем контрольную проверку
                    Dim lContr As Double = (startPointBeam - endPointBeam).Length
                    Dim deltaLenght As Double = shortLenghtBeam - lContr
                    If Math.Abs(deltaLenght) >= 0.001 Then
                        'корректируем балку по длине
                        Dim boolExtendBeam As Boolean = FuncExtendBearm(lineShortBeam, 0, deltaLenght, 3)
                        'садим балку точно на ось
                        Dim pk As Double = 0
                        Dim off As Double = 0
                        Dim boolFindPk As Boolean = axisPline3D.PosToStaOffset(endPointBeam.Pos, pk, off)
                        If boolFindPk = True Then
                            Dim newEndPoint As Vector2D = PolylineExtentions.StaOffsetToPos(axisPline3D, pk, 0)
                            endPointBeam = New Vector3D(newEndPoint, endPointBeam.Z)
                            lineShortBeam.EndPoint = endPointBeam
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Catch ex As System.Exception
                    Return False
                End Try
            Next i
            If IsNothing(axisPline) = False Then
                ActivDocument.ActiveSpace.Entities.Remove(axisPline)
            End If
            Dim lContrFinal As Double = (startPointBeam - endPointBeam).Length
            Dim deltaLenghtFinal As Double = shortLenghtBeam - lContrFinal
            If Math.Abs(deltaLenghtFinal) <= 0.001 Then
                Return True
            End If
        Catch ex As Newtonsoft.Json.JsonException
            Return False
        End Try
    End Function
    'функция возвращает минимальный зазор
    Public Shared Function FuncReturnZazorBeams(ByRef lineShortBearm As DwgLine, ByRef pointEndSectionColl As List(Of Vector3D), ByVal heightBearm As Double, ByVal widthDownBearm As Double, ByVal widthUpBearm As Double, ByVal startDistToPointPr As Double, ByVal endDistToPointPr As Double, ByVal userAlign As Alignment) As Double
        FuncReturnZazorBeams = 0
        If IsNothing(lineShortBearm) = True Then Exit Function
        Try
            Dim PkStartLineShortBearm As Double = 0
            Dim PkEndLineShortBearm As Double = 0
            Dim off As Double = 0
            Dim boolpk1start As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(lineShortBearm.StartPoint.Pos, PkStartLineShortBearm, off)
            Dim boolpk1end As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(lineShortBearm.EndPoint.Pos, PkEndLineShortBearm, off)
            If boolpk1end = True And boolpk1start = True Then
                If PkStartLineShortBearm > PkEndLineShortBearm Then
                    Dim tempPoint As Vector3D = lineShortBearm.StartPoint
                    lineShortBearm.StartPoint = lineShortBearm.EndPoint
                    lineShortBearm.EndPoint = tempPoint
                End If
            End If
            '1. Удлинняем балку (временно, на величину участков опирания)
            Dim lenghtShortBearm As Double = lineShortBearm.Length
            Dim lineBearm As DwgLine = lineShortBearm.Clone() 'работаем с копией
            Dim boolExtBearm As Boolean = FuncExtendBearm(lineBearm, startDistToPointPr, endDistToPointPr)
            If boolExtBearm = False Then Exit Function
            '2 делаем смещение балки вверх
            '2.1 находим угол наклона балки
            Dim b As Double = lineBearm.EndPoint.Z - lineBearm.StartPoint.Z
            Dim c As Double = lineBearm.Length
            Dim i As Double = Math.Asin(b / c)
            Dim wAngle As Double = (i * 180) / Math.PI
            Dim insCoord As Vector2D = New Vector2D(0, 0)
            Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, heightBearm)
            'получаем новые координаты верха балки
            Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.StartPoint.X, lineBearm.StartPoint.Y), lineBearm.Rotation, deltaXZ.X)
            Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.EndPoint.X, lineBearm.EndPoint.Y), lineBearm.Rotation, deltaXZ.X)
            Dim lContr As Double = MathFunction.funcCalcDistanceByToPoints2d(startCoordUpBearm, endCoordUpBearm) 'первая проверка
            'создаем верх балки
            Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineBearm.StartPoint.Z + deltaXZ.Y)
            Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineBearm.EndPoint.Z + deltaXZ.Y)
            Dim lContr2 As Double = (startPointUpBalkaPr - endPointUpBalkaPr).Length 'вторая проверка

            Dim lineUpBalka As DwgLine = New DwgLine()
            lineUpBalka.StartPoint = startPointUpBalkaPr
            lineUpBalka.EndPoint = endPointUpBalkaPr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '3. приводим балки к одному уровню (временно)
            If pointEndSectionColl.Count = 4 Then
                'верх балки предыдущей
                Dim upLeftPrev As Vector3D = pointEndSectionColl.Item(2)
                Dim upRightPrev As Vector3D = pointEndSectionColl.Item(3)
                'низ балки предыдущей
                Dim downLeftPrev As Vector3D = pointEndSectionColl.Item(0)
                Dim downRightPrev As Vector3D = pointEndSectionColl.Item(1)
                'высота по низу предыдущей балки
                Dim Hprev As Double = (downLeftPrev.Z + downRightPrev.Z) / 2
                'высота по низу балки определяемой
                Dim Hbm As Double = lineBearm.StartPoint.Z
                'дельта по высоте
                Dim deltaH As Double = Hprev - Hbm
                'временные координаты для предыдущей балки если она ниже чем расчетная
                Dim downLeftPrev2 As Vector3D = pointEndSectionColl.Item(0)
                Dim downRightPrev2 As Vector3D = pointEndSectionColl.Item(1)
                If deltaH < 0 Then
                    'предудущая балка ниже чем расчетная (приводим ее уровню расчетной бплки
                    downLeftPrev2 = MathFunction.FuncCalcPointInLine(downLeftPrev, upLeftPrev, Math.Abs(deltaH))
                    downRightPrev2 = MathFunction.FuncCalcPointInLine(downRightPrev, upRightPrev, Math.Abs(deltaH))
                Else 'расчетная балка ниже чем предыдущая (подтягиваем расчетную балку к предыдущей)
                    Dim tempPointB1 As Vector3D = MathFunction.FuncCalcPointInLine(lineBearm.StartPoint, lineUpBalka.StartPoint, deltaH)
                    Dim tempPointB2 As Vector3D = MathFunction.FuncCalcPointInLine(lineBearm.EndPoint, lineUpBalka.EndPoint, deltaH)
                    lineBearm.StartPoint = tempPointB1
                    lineBearm.EndPoint = tempPointB2
                End If
                'делаем повторнуб проверку
                'высота по низу предыдущей балки
                Hprev = (downLeftPrev2.Z + downRightPrev2.Z) / 2
                'высота по низу балки определяемой
                Hbm = lineBearm.StartPoint.Z
                'дельта по высоте
                deltaH = Hprev - Hbm
                If Math.Abs(deltaH) > 0.05 Then
                    MsgBox("Не удалось привести балки к одному уровню!!!")
                End If
            End If
            '4. распаралеливаем линии расчетной балки
            Dim minZazor As Double = 99999999999
            Dim maxZazor As Double = 0
            'считаем зазор по низу балки
            '====================================================================================================================
            'смещение низ право
            Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
            Dim offsetElement As Double = widthDownBearm / 2
Line1:
            lineBearm.Offset(dbCollection1, offsetElement)
            If dbCollection1.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection1.Item(0)
                'удлинняем линию для нахождения точки
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count > 3 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(1).Pos, pointEndSectionColl.Item(0).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(1).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(0).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection1 = New List(Of DwgEntity)
                                GoTo Line1
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection1 = New List(Of DwgEntity)
                                GoTo Line1
                            End If
                        End If
                        'ищем угол
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        'балки сонаправлены
                        If dAngle < 2 Then
                            If dist < minZazor Then
                                minZazor = dist
                            End If
                        Else
                            If dist > maxZazor Then
                                maxZazor = dist
                            End If
                        End If
                    End If
                End If
            End If
            '====================================================================================================================
            'низ лево'
            Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = -1 * widthDownBearm / 2
Line2:
            lineBearm.Offset(dbCollection2, offsetElement)
            If dbCollection2.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection2.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(0).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(1).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(0).Pos, pointEndSectionColl.Item(1).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection2 = New List(Of DwgEntity)
                                GoTo Line2
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection2 = New List(Of DwgEntity)
                                GoTo Line2
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        'балки сонаправлены
                        If dAngle < 2 Then
                            If dist < minZazor Then
                                minZazor = dist
                            End If
                        Else
                            If dist > maxZazor Then
                                maxZazor = dist
                            End If
                        End If
                    End If
                End If
            End If
            'считаем зазор по верху балки
            '====================================================================================================================
            'верх право
            Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = widthUpBearm / 2
line3:
            lineUpBalka.Offset(dbCollection3, offsetElement)
            If dbCollection3.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection3.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(3).Pos, pointEndSectionColl.Item(2).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(3).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(2).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection3 = New List(Of DwgEntity)
                                GoTo line3
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement - deltaDist
                                dbCollection3 = New List(Of DwgEntity)
                                GoTo line3
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        'балки сонаправлены
                        If dAngle < 2 Then
                            If dist < minZazor Then
                                minZazor = dist
                            End If
                        Else
                            If dist > maxZazor Then
                                maxZazor = dist
                            End If
                        End If
                    End If
                End If
            End If
            '====================================================================================================================
            'верх лево
            Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
            offsetElement = -1 * widthUpBearm / 2
line4:
            lineUpBalka.Offset(dbCollection4, offsetElement)
            If dbCollection4.Count = 1 Then
                Dim tempLine As DwgLine = dbCollection4.Item(0)
                Dim pointIntersect As Vector2D = Nothing
                If pointEndSectionColl.Count = 4 Then
                    Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(tempLine.EndPoint.Pos, tempLine.StartPoint.Pos, pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos, pointIntersect)
                    If boolPointIntersect = True Then
                        'если расчетная балка шире чем предыдущая делаем коррекцию
                        Dim tempDist1 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(2).Pos)
                        Dim tempDist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, pointEndSectionColl.Item(3).Pos)
                        Dim tempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointEndSectionColl.Item(2).Pos, pointEndSectionColl.Item(3).Pos)
                        If tempDist1 > tempDist Then
                            Dim deltaDist As Double = tempDist1 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection4 = New List(Of DwgEntity)
                                GoTo line4
                            End If
                        ElseIf tempDist2 > tempDist Then
                            Dim deltaDist As Double = tempDist2 - tempDist
                            If deltaDist > 0.02 Then
                                offsetElement = offsetElement + deltaDist
                                dbCollection4 = New List(Of DwgEntity)
                                GoTo line4
                            End If
                        End If
                        Dim angle As Double = (pointIntersect - tempLine.StartPoint.Pos).Angle
                        Dim dist As Double = (pointIntersect - tempLine.StartPoint.Pos).Length
                        Dim dAngle As Double = Math.Abs(angle - tempLine.Rotation)
                        'балки сонаправлены
                        If dAngle < 2 Then
                            If dist < minZazor Then
                                minZazor = dist
                            End If
                        Else
                            If dist > maxZazor Then
                                maxZazor = dist
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function
    'функция возвращает зазор между определенным элементом и плоскостью предыдущей балки
    Public Shared Function FuncReturnZazorByElrmentAndBeam(ByRef lineElements As DwgLine, ByRef pointSectionColl As List(Of Vector3D)) As Double
        FuncReturnZazorByElrmentAndBeam = 0
        If IsNothing(lineElements) = True Then Exit Function
        If lineElements.Length = 0 Then Exit Function
        Try
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '3. приводим балки к одному уровню (временно)
            If pointSectionColl.Count = 4 Then
                'верх балки предыдущей
                Dim upLeftPrev As Vector3D = pointSectionColl.Item(2)
                Dim upRightPrev As Vector3D = pointSectionColl.Item(3)
                'низ балки предыдущей
                Dim downLeftPrev As Vector3D = pointSectionColl.Item(0)
                Dim downRightPrev As Vector3D = pointSectionColl.Item(1)
                'высота по низу предыдущей балки
                Dim downPrev As Double = (downLeftPrev.Z + downRightPrev.Z) / 2
                Dim upPrev As Double = (downLeftPrev.Z + downRightPrev.Z) / 2
                'высота по низу балки определяемой
                Dim Hbm As Double = lineElements.StartPoint.Z
                'корректируем балку по высоте
                If Hbm > downPrev And Hbm < upPrev Then

                End If
            End If
        Catch ex As Exception

        End Try





    End Function
    'функция читает свойства балки
    Public Shared Function FuncReturnPropertiesBeam(ByVal axisLineBearm As DwgLine, ByRef name As String, ByRef userBeam As Beams, ByRef IDBeam As String, ByRef idBridge As String, ByRef note As String) As Boolean
        FuncReturnPropertiesBeam = False
        '1.считываем расширенные данные с предыдущей балки
        Dim arrayDataBeam As String(,) = Nothing
        Try
            Dim boolReadXDataPrevBeam As Boolean = FuncXRecords.FuncReadXData(axisLineBearm, "PROJECT_BRIDGE", arrayDataBeam)
            If IsArray(arrayDataBeam) = True Then
                name = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "Name", 0, 1)
                IDBeam = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "ElementID", 0, 1)
                idBridge = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "BrigeID", 0, 1)
                note = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "NOTE", 0, 1)
                If name Like "Ось балки" Then
                    Dim strGSONPrevBeam As String = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "KeyParameters", 0, 1)
                    userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(strGSONPrevBeam)
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'функция читает свойства мостового сооружения
    Public Shared Function FuncReturnPropertiesBridge(ByVal axisLineBridge As DwgPolyline, ByRef name As String, ByRef userBridge As Bridge, ByRef IDBeam As String, ByRef idBridge As String, ByRef note As String) As Boolean
        FuncReturnPropertiesBridge = False
        '1.считываем расширенные данные с предыдущей балки
        Dim arrayDataBeam As String(,) = Nothing
        Try
            Dim boolReadXDataPrevBeam As Boolean = FuncXRecords.FuncReadXData(axisLineBridge, "PROJECT_BRIDGE", arrayDataBeam)
            If IsArray(arrayDataBeam) = True Then
                name = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "Name", 0, 1)
                IDBeam = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "ElementID", 0, 1)
                idBridge = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "BrigeID", 0, 1)
                note = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "NOTE", 0, 1)
                If name Like "Главная ось сооружения" Or name Like "главная ось сооружения" Then
                    Dim strGSONPrevBeam As String = MathFunction.FuncFindValueToArray2d(arrayDataBeam, "KeyParameters", 0, 1)
                    userBridge = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Bridge)(strGSONPrevBeam)
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'функция восстанавливает трассу для выбранной балки в виде 3д полилинии
    Public Shared Function FuncReturnAlignmentPolyline(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByVal nameAlign As String, ByRef align As Alignment, ByRef surf As Surface, Optional offsetAxis As Double = 0) As DwgPolyline
        FuncReturnAlignmentPolyline = Nothing
        '3. читаем трассу
        Dim axisIPline3D As DwgPolyline = New DwgPolyline()
        Dim modelRoads As IProjectModel() = PluginCoreOps.FindModels(New String() {"road"})
        For Each child As IProjectModel In modelRoads
            Dim modelUri As Topomatic.FoundationClasses.URI = child.Uri
            Dim fileNameSfcx As String = IO.Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
            If fileNameSfcx.Trim Like nameAlign Then
                Dim userRoadModel As RoadModel = child.Model
                If IsNothing(userRoadModel) = False Then
                    Dim axisPline0 As DwgPolyline = Nothing 'полилиния из трассы  в прямом направлении
                    align = userRoadModel.Alignment
                    surf = userRoadModel.Surface
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    axisIPline3D = FuncAlignment.FuncOffsetAlignment(ActivDocument, align, offsetAxis)
                    Return axisIPline3D
                End If
            End If
        Next
    End Function

    'функция записывает в массив ObjectID вспомогательных построений 
    Public Shared Function FuncFindElementsBeam(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByVal beamsElementDictionary As Dictionary(Of String, String(,)), ByVal idBeam As String, ByRef arrayElements As DwgLine(), ByRef tlcObject As DwgModel3DElement) As Boolean
        FuncFindElementsBeam = False
        'находим пикеты начала и конца предудущей балки
        If IsNothing(idBeam) = True Then Exit Function
        If idBeam.Trim.Length < 2 Then Exit Function
        If beamsElementDictionary.Count = 0 Then Exit Function
        ReDim arrayElements(3)
        For i As Integer = 0 To beamsElementDictionary.Count - 1
            Dim userKeysPairs As KeyValuePair(Of String, String(,)) = beamsElementDictionary.ElementAt(i)
            Dim arrayTempElements As String(,) = userKeysPairs.Value
            If IsArray(arrayTempElements) = True Then
                For j As Integer = 0 To arrayTempElements.GetUpperBound(1)
                    Dim idTempBeams As String = arrayTempElements(2, j)
                    If idBeam Like idTempBeams Then
                        Dim nameObject As String = arrayTempElements(0, j)
                        Dim keyParams As Double = Val(arrayTempElements(1, j))
                        Dim objectID As UInteger = Val(arrayTempElements(4, j))
                        Dim acTempObject As DwgObject = Nothing
                        Dim boolFindObject As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(objectID, acTempObject)
                        If boolFindObject = True Then
                            If TypeOf acTempObject Is DwgLine Then
                                If nameObject Like "Низ ребра балки" Then
                                    If keyParams < 0 Then
                                        arrayElements(0) = acTempObject 'низ лево
                                    Else
                                        arrayElements(1) = acTempObject 'низ право
                                    End If
                                ElseIf nameObject Like "Верх грани плиты балки" Then
                                    If keyParams < 0 Then
                                        arrayElements(2) = acTempObject 'верх лево
                                    Else
                                        arrayElements(3) = acTempObject 'верх право
                                    End If
                                End If
                            ElseIf TypeOf acTempObject Is DwgModel3DElement Then
                                If nameObject Like "Балка" Then
                                    tlcObject = acTempObject
                                End If
                            End If
                        End If
                    End If
                Next j
            End If
        Next
    End Function
    'функция корректирует массив ObjectID вспомогательных построений 
    Public Shared Function FuncCorrectElementsBeam(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByRef arrayElements As DwgLine(), ByRef pointStartSectionColl As List(Of Vector3D), ByRef pointEndSectionColl As List(Of Vector3D), ByRef tlcObject As DwgModel3DElement, Optional templateXml As String = "") As Boolean
        FuncCorrectElementsBeam = False
        'находим пикеты начала и конца предудущей балки
        Dim boolCorrectBottomRightLine As Boolean = False
        Dim boolCorrectBottomLeftLine As Boolean = False
        Dim boolCorrectTopRightLine As Boolean = False
        Dim boolCorrectTopLeftLine As Boolean = False
        If IsArray(arrayElements) = True Then
            For i As Integer = 0 To arrayElements.Length - 1
                If i = 0 Then
                    Dim acLine1 As DwgLine = arrayElements(0)
                    acLine1.StartPoint = pointStartSectionColl(1)
                    acLine1.EndPoint = pointEndSectionColl(1)
                ElseIf i = 1 Then
                    Dim acLine1 As DwgLine = arrayElements(1)
                    acLine1.StartPoint = pointStartSectionColl(0)
                    acLine1.EndPoint = pointEndSectionColl(0)
                ElseIf i = 2 Then
                    Dim acLine1 As DwgLine = arrayElements(2)
                    acLine1.StartPoint = pointStartSectionColl(3)
                    acLine1.EndPoint = pointEndSectionColl(3)
                ElseIf i = 3 Then
                    Dim acLine1 As DwgLine = arrayElements(3)
                    acLine1.StartPoint = pointStartSectionColl(2)
                    acLine1.EndPoint = pointEndSectionColl(2)
                End If
            Next i
        End If
    End Function
    'функция восстанавливает балку по ее параметрам (без рисования вспомогательных линий)
    Public Shared Function FuncRestoreCoordinatesBalka(ByRef lineShortBearm As DwgLine, ByVal heightBearm As Double, ByVal widthDownBearm As Double, ByVal widthUpBearm As Double, ByVal startDistToPointPr As Double, ByVal endDistToPointPr As Double, ByRef pointStartSectionColl As List(Of Vector3D), ByRef pointEndSectionColl As List(Of Vector3D)) As Boolean
        FuncRestoreCoordinatesBalka = False
        If IsNothing(lineShortBearm) = True Then Exit Function
        '1. Удлинняем балку (временно, на величину участков опирания)
        Dim lenghtShortBearm As Double = lineShortBearm.Length
        Dim lineBearm As DwgLine = lineShortBearm.Clone()
        Dim boolExtBearm As Boolean = FuncExtendBearm(lineBearm, startDistToPointPr, endDistToPointPr)
        '1.делаем смещение балки вверх
        Dim b As Double = lineShortBearm.EndPoint.Z - lineShortBearm.StartPoint.Z
        Dim c As Double = lineShortBearm.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, heightBearm)
        'получаем новые координаты верха балки
        Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.StartPoint.X, lineBearm.StartPoint.Y), lineBearm.Rotation, deltaXZ.X)
        Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineBearm.EndPoint.X, lineBearm.EndPoint.Y), lineBearm.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineBearm.StartPoint.Z + deltaXZ.Y)
        Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineBearm.EndPoint.Z + deltaXZ.Y)
        Dim lineUpBalka As DwgLine = New DwgLine()
        lineUpBalka.StartPoint = startPointUpBalkaPr
        lineUpBalka.EndPoint = endPointUpBalkaPr
        '====================================================================================================================
        Dim pointStart1 As Vector3D = Nothing
        Dim pointStart2 As Vector3D = Nothing
        Dim pointStart3 As Vector3D = Nothing
        Dim pointStart4 As Vector3D = Nothing
        Dim pointEnd1 As Vector3D = Nothing
        Dim pointEnd2 As Vector3D = Nothing
        Dim pointEnd3 As Vector3D = Nothing
        Dim pointEnd4 As Vector3D = Nothing
        'смещение низ право
        Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineBearm.Offset(dbCollection1, widthDownBearm / 2)
        If dbCollection1.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection1.Item(0)
            pointStart1 = tempLine.StartPoint
            pointEnd1 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение низ лево
        Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineBearm.Offset(dbCollection2, -1 * widthDownBearm / 2)
        If dbCollection2.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection2.Item(0)
            pointStart2 = tempLine.StartPoint
            pointEnd2 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение верх право
        Dim dbCollection3 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineUpBalka.Offset(dbCollection3, widthUpBearm / 2)
        If dbCollection3.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection3.Item(0)
            pointStart3 = tempLine.StartPoint
            pointEnd3 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение верх лево
        Dim dbCollection4 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineUpBalka.Offset(dbCollection4, -1 * widthUpBearm / 2)
        If dbCollection4.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection4.Item(0)
            pointStart4 = tempLine.StartPoint
            pointEnd4 = tempLine.EndPoint
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        pointEndSectionColl = New List(Of Vector3D)
        pointEndSectionColl.Add(pointEnd1)
        pointEndSectionColl.Add(pointEnd2)
        pointEndSectionColl.Add(pointEnd3)
        pointEndSectionColl.Add(pointEnd4)
        pointStartSectionColl = New List(Of Vector3D)
        pointStartSectionColl.Add(pointStart1)
        pointStartSectionColl.Add(pointStart2)
        pointStartSectionColl.Add(pointStart3)
        pointStartSectionColl.Add(pointStart4)
        Return True
    End Function




    'функция возвращает координаты верха балки (балка короткая)(1 коллекция начало право, лево, 2 коллекия конец, право, лево)
    Public Shared Function FuncReturnPositionsTopElementBeam(ByRef lineShortBearm As DwgLine, ByVal userBeam As Beams, ByRef pointStartSectionColl As List(Of Vector3D), ByRef pointEndSectionColl As List(Of Vector3D)) As Boolean
        FuncReturnPositionsTopElementBeam = False
        If IsNothing(lineShortBearm) = True Then Exit Function
        '1.делаем смещение балки вверх
        Dim b As Double = lineShortBearm.EndPoint.Z - lineShortBearm.StartPoint.Z
        Dim c As Double = lineShortBearm.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, userBeam.height)
        'получаем новые координаты верха балки
        Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineShortBearm.StartPoint.Pos, lineShortBearm.Rotation, deltaXZ.X)
        Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(lineShortBearm.EndPoint.Pos, lineShortBearm.Rotation, deltaXZ.X)

        'создаем верх балки
        Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineShortBearm.StartPoint.Z + deltaXZ.Y)
        Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineShortBearm.EndPoint.Z + deltaXZ.Y)

        'Dim startPointUpBalkaPr1 As Vector3D = New Vector3D(startCoordUpBearm1.X, startCoordUpBearm1.Y, lineShortBearm.StartPoint.Z + deltaXZ1.Y)
        'Dim endPointUpBalkaPr1 As Vector3D = New Vector3D(endCoordUpBearm1.X, endCoordUpBearm1.Y, lineShortBearm.EndPoint.Z + deltaXZ1.Y)

        Dim lineUpBalka As DwgLine = New DwgLine()
        lineUpBalka.StartPoint = startPointUpBalkaPr
        lineUpBalka.EndPoint = endPointUpBalkaPr
        '====================================================================================================================
        Dim pointStart1 As Vector3D = Nothing 'верх право
        Dim pointStart2 As Vector3D = Nothing 'верх лево
        Dim pointEnd1 As Vector3D = Nothing 'верх право
        Dim pointEnd2 As Vector3D = Nothing 'верх лево
        '====================================================================================================================
        'смещение верх право
        Dim dbCollection1 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineUpBalka.Offset(dbCollection1, userBeam.widthTop / 2)
        If dbCollection1.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection1.Item(0)
            'правая линия
            pointStart1 = tempLine.StartPoint
            pointEnd1 = tempLine.EndPoint
        End If
        '====================================================================================================================
        'смещение верх лево
        Dim dbCollection2 As List(Of DwgEntity) = New List(Of DwgEntity)
        lineUpBalka.Offset(dbCollection2, -1 * userBeam.widthTop / 2)
        If dbCollection2.Count = 1 Then
            Dim tempLine As DwgLine = dbCollection2.Item(0)
            'левая линия
            pointStart2 = tempLine.StartPoint
            pointEnd2 = tempLine.EndPoint
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        pointStartSectionColl = New List(Of Vector3D)
        pointStartSectionColl.Add(pointStart1)
        pointStartSectionColl.Add(pointStart2)

        pointEndSectionColl = New List(Of Vector3D)
        pointEndSectionColl.Add(pointEnd1)
        pointEndSectionColl.Add(pointEnd2)

        Return True
    End Function

    'функция считает зазор между балками
    Public Shared Function FuncCalculateZazor(ByVal userAlign As Alignment, ByVal pointStartSectionColl1 As List(Of Vector3D), ByVal pointEndSectionColl1 As List(Of Vector3D), ByVal pointStartSectionColl2 As List(Of Vector3D), ByVal pointEndSectionColl2 As List(Of Vector3D), ByRef arrayZazor As Double()) As Boolean
        FuncCalculateZazor = False
        ReDim arrayZazor(3)
        If IsNothing(userAlign) = True Then
            Return False
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'проверяемая балка
        Dim startBeam1PointTopLeft As Vector3D = pointStartSectionColl1.Item(0)
        Dim startBeam1PointTopRight As Vector3D = pointStartSectionColl1.Item(1)
        Dim startBeam1PointDownLeft As Vector3D = pointStartSectionColl1.Item(2)
        Dim startBeam1PointDownRight As Vector3D = pointStartSectionColl1.Item(3)

        Dim endBeam1PointTopLeft As Vector3D = pointEndSectionColl1.Item(0) 'используется в расчете
        Dim endBeam1PointTopRight As Vector3D = pointEndSectionColl1.Item(1)
        Dim endBeam1PointDownLeft As Vector3D = pointEndSectionColl1.Item(2)
        Dim endBeam1PointDownRight As Vector3D = pointEndSectionColl1.Item(3)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смежная с проверяемой балкой
        Dim startBeam2PointTopLeft As Vector3D = pointStartSectionColl2.Item(0) 'используется в расчете
        Dim startBeam2PointTopRight As Vector3D = pointStartSectionColl2.Item(1)
        Dim startBeam2PointDownLeft As Vector3D = pointStartSectionColl2.Item(2)
        Dim startBeam2PointDownRight As Vector3D = pointStartSectionColl2.Item(3)

        Dim endBeam2PointTopLeft As Vector3D = pointEndSectionColl2.Item(0)
        Dim endBeam2PointTopRight As Vector3D = pointEndSectionColl2.Item(1)
        Dim endBeam2PointDownLeft As Vector3D = pointEndSectionColl2.Item(2)
        Dim endBeam2PointDownRight As Vector3D = pointEndSectionColl2.Item(3)

        Dim pk1start As Double = -1
        Dim pk1End As Double = -1

        Dim pk2Start As Double = -1
        Dim pk2End As Double = -1
        Dim off1 As Double = -1
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'делаем реверс балки если она направлена против хода пикетажа
        Dim boolpk1start As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(startBeam1PointTopLeft, pk1start, off1)
        Dim boolpk1end As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(endBeam1PointTopLeft, pk1End, off1)
        If boolpk1start = True And boolpk1end = True Then
            If pk1start > pk1End Then
                endBeam1PointTopLeft = pointStartSectionColl1.Item(0)
                endBeam1PointTopRight = pointStartSectionColl1.Item(1)
                endBeam1PointDownLeft = pointStartSectionColl1.Item(2)
                endBeam1PointDownRight = pointStartSectionColl1.Item(3)

                startBeam1PointTopLeft = pointEndSectionColl1.Item(0)
                startBeam1PointTopRight = pointEndSectionColl1.Item(1)
                startBeam1PointDownLeft = pointEndSectionColl1.Item(2)
                startBeam1PointDownRight = pointEndSectionColl1.Item(3)
            End If
        End If
        Dim boolpk2start As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(startBeam2PointTopLeft, pk2Start, off1)
        Dim boolpk2end As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(endBeam2PointTopLeft, pk2End, off1)
        If boolpk2start = True And boolpk2end = True Then
            If pk2End < pk2Start Then
                startBeam2PointTopLeft = pointEndSectionColl2.Item(0)
                startBeam2PointTopRight = pointEndSectionColl2.Item(1)
                startBeam2PointDownLeft = pointEndSectionColl2.Item(2)
                startBeam2PointDownRight = pointEndSectionColl2.Item(3)

                endBeam2PointTopLeft = pointStartSectionColl2.Item(0)
                endBeam2PointTopRight = pointStartSectionColl2.Item(1)
                endBeam2PointDownLeft = pointStartSectionColl2.Item(2)
                endBeam2PointDownRight = pointStartSectionColl2.Item(3)
            End If
        End If
        'абсолютная высота по низу предудущей балки
        Dim Hprev As Double = (startBeam2PointDownLeft.Z + startBeam2PointDownRight.Z) / 2
        'абсолютная высота по низу расчетной балки
        Dim Hbm As Double = (endBeam1PointDownLeft.Z + endBeam1PointDownRight.Z) / 2
        'разность высот
        Dim deltaH As Double = Hprev - Hbm

        Dim downLeftPrev2 As Vector3D = startBeam2PointDownLeft
        Dim downRightPrev2 As Vector3D = startBeam2PointDownRight

        If Math.Abs(deltaH) > 0.01 Then
            If deltaH < 0 Then
                startBeam1PointDownLeft = MathFunction.FuncCalcPointInLine(downLeftPrev2, startBeam2PointTopLeft, Math.Abs(deltaH))
                startBeam1PointDownRight = MathFunction.FuncCalcPointInLine(downRightPrev2, startBeam2PointTopRight, Math.Abs(deltaH))
            Else
                startBeam2PointDownLeft = MathFunction.FuncCalcPointInLine(downLeftPrev2, startBeam2PointTopLeft, deltaH)
                startBeam2PointDownRight = MathFunction.FuncCalcPointInLine(downRightPrev2, startBeam2PointTopRight, deltaH)
            End If
        End If
        'абсолютная высота по низу предудущей балки
        Hprev = (startBeam2PointDownLeft.Z + startBeam2PointDownRight.Z) / 2
        'абсолютная высота по низу расчетной балки
        Hbm = (endBeam1PointDownLeft.Z + endBeam1PointDownRight.Z) / 2
        'разность высот
        deltaH = Hprev - Hbm
        If Math.Abs(deltaH) > 0.01 Then
            MsgBox("Ошибка в высоте." & deltaH)
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'считаем зазоры
        'верх лево
        Dim pointIntersect As Vector2D = Nothing
        Dim boolPointIntersect As Boolean = MathFunction.FuncIntersectionTwoRay(endBeam1PointTopRight.Pos, endBeam1PointTopLeft.Pos, endBeam2PointTopLeft.Pos, startBeam2PointTopLeft.Pos, pointIntersect)
        If boolPointIntersect = True Then
            'ищем угол
            Dim angle As Double = (pointIntersect - startBeam2PointTopLeft.Pos).Angle
            arrayZazor(0) = (pointIntersect - startBeam2PointTopLeft.Pos).Length
        Else
            boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointTopLeft.Pos, endBeam1PointTopRight.Pos, endBeam2PointTopLeft.Pos, startBeam2PointTopLeft.Pos, pointIntersect)
            If boolPointIntersect = True Then
                Dim angle As Double = (pointIntersect - startBeam2PointTopLeft.Pos).Angle
                arrayZazor(0) = (pointIntersect - startBeam2PointTopLeft.Pos).Length
            End If
        End If
        '====================================================================================================================
        'верх право
        pointIntersect = Nothing
        boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointTopLeft.Pos, endBeam1PointTopRight.Pos, endBeam2PointTopRight.Pos, startBeam2PointTopRight.Pos, pointIntersect)
        If boolPointIntersect = True Then
            Dim angle As Double = (pointIntersect - startBeam2PointTopRight.Pos).Angle
            arrayZazor(1) = (pointIntersect - startBeam2PointTopRight.Pos).Length
        Else
            boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointTopRight.Pos, endBeam1PointTopLeft.Pos, endBeam2PointTopRight.Pos, startBeam2PointTopRight.Pos, pointIntersect)
            If boolPointIntersect = True Then
                Dim angle As Double = (pointIntersect - startBeam2PointTopRight.Pos).Angle
                arrayZazor(1) = (pointIntersect - startBeam2PointTopRight.Pos).Length
            End If
        End If
        '====================================================================================================================
        'низ лево
        pointIntersect = Nothing
        boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointDownRight.Pos, endBeam1PointDownLeft.Pos, endBeam2PointDownLeft.Pos, startBeam2PointDownLeft.Pos, pointIntersect)
        If boolPointIntersect = True Then
            Dim angle As Double = (pointIntersect - endBeam2PointDownLeft.Pos).Angle
            arrayZazor(2) = (pointIntersect - startBeam2PointDownLeft.Pos).Length
        Else
            boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointDownLeft.Pos, endBeam1PointDownRight.Pos, endBeam2PointDownLeft.Pos, startBeam2PointDownLeft.Pos, pointIntersect)
            If boolPointIntersect = True Then
                Dim angle As Double = (pointIntersect - endBeam2PointDownLeft.Pos).Angle
                arrayZazor(2) = (pointIntersect - startBeam2PointDownLeft.Pos).Length
            End If
        End If
        '====================================================================================================================
        'низ право
        pointIntersect = Nothing
        boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointDownLeft.Pos, endBeam1PointDownRight.Pos, endBeam2PointDownRight.Pos, startBeam2PointDownRight.Pos, pointIntersect)
        If boolPointIntersect = True Then
            Dim angle As Double = (pointIntersect - startBeam2PointDownRight.Pos).Angle
            arrayZazor(3) = (pointIntersect - startBeam2PointDownRight.Pos).Length
        Else
            boolPointIntersect = MathFunction.FuncIntersectionTwoRay(endBeam1PointDownRight.Pos, endBeam1PointDownLeft.Pos, endBeam2PointDownRight.Pos, startBeam2PointDownRight.Pos, pointIntersect)
            If boolPointIntersect = True Then
                Dim angle As Double = (pointIntersect - startBeam2PointDownRight.Pos).Angle
                arrayZazor(3) = (pointIntersect - startBeam2PointDownRight.Pos).Length
            End If
        End If
        Return True
    End Function

    'коррекция горизонтального расстояния за угол наклона линии
    Public Shared Function FuncFindPointBearm(ByVal axisPline As DwgPolyline, ByVal startPoint As Vector3D, ByVal radius As Double, ByVal surf As Surface, ByVal heightBearm As Double, ByVal dEarth As Double) As Vector3D
        Dim axisPline3D As IPolyline3D = New Polyline3D()
        axisPline.GetPolyline(axisPline3D)
        Dim startdist As Double = -1
        Dim off As Double = -1
        'находим расстояние на полилинии
        Dim boolDist As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startPoint.Pos, startdist, off)
        If boolDist = True Then
            Dim findDist As Double = radius
            For i As Integer = 0 To 1000
                Dim delta As Double = startdist + findDist
                Dim tempPoint As Vector2D = Nothing
                Try
                    tempPoint = PolylineExtentions.StaOffsetToPos(axisPline3D, delta, 0)
                Catch ex As System.ArgumentOutOfRangeException
                    Return New Vector3D(-1, -1, -1)
                End Try
                Dim elev As Double = startPoint.Z
                Try
                    elev = surf.GetElevation(tempPoint)
                Catch ex As System.Exception
                    Return New Vector3D(-1, -1, -1)
                End Try
                elev = elev - heightBearm - dEarth
                Dim tempPointNew As Vector3D = New Vector3D(tempPoint.X, tempPoint.Y, elev)
                Dim tempDist As Double = (startPoint - tempPointNew).Length
                Dim dLenght As Double = radius - tempDist
                If Math.Abs(dLenght) < 0.001 Then
                    Return tempPointNew
                Else
                    findDist = findDist + dLenght
                End If
            Next
        Else
            Return New Vector3D(-1, -1, -1)
        End If
    End Function
    'функция удлинняет/укорачивает балку
    Public Shared Function FuncExtendBearm(ByRef lineBearm As DwgLine, ByVal startLenght As Double, ByVal endLenght As Double, Optional round As Integer = 3) As Boolean
        FuncExtendBearm = False
        If IsNothing(lineBearm) = True Then Return False
        Try
            Dim L As Double = lineBearm.Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = lineBearm.StartPoint.X + k1 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y1 As Double = lineBearm.StartPoint.Y + k1 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z1 As Double = lineBearm.StartPoint.Z + k1 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim startPoint As Vector3D = New Vector3D(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = lineBearm.EndPoint.X + k2 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y2 As Double = lineBearm.EndPoint.Y + k2 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z2 As Double = lineBearm.EndPoint.Z + k2 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim endPoint As Vector3D = New Vector3D(X2, Y2, z2)
            lineBearm.StartPoint = startPoint
            lineBearm.EndPoint = endPoint
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function

    'функция делает перенос оси балки
    Public Shared Function FuncMoveBearm(ByRef lineBearm As DwgLine, ByVal deltaLenght As Double, Optional round As Integer = 3) As Boolean
        FuncMoveBearm = False
        Try
            Dim L As Double = lineBearm.Length
            Dim k As Double = deltaLenght / L
            Dim X1 As Double = lineBearm.StartPoint.X + k * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y1 As Double = lineBearm.StartPoint.Y + k * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z1 As Double = lineBearm.StartPoint.Z + k * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim startPoint As Vector3D = New Vector3D(X1, Y1, z1)

            Dim X2 As Double = lineBearm.EndPoint.X + k * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y2 As Double = lineBearm.EndPoint.Y + k * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z2 As Double = lineBearm.EndPoint.Z + k * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            Dim endPoint As Vector3D = New Vector3D(X2, Y2, z2)

            lineBearm.StartPoint = startPoint
            lineBearm.EndPoint = endPoint
            Return True
        Catch ex As Exception
        End Try
    End Function

    'функция возвращает точки начала и конца виртуального удлиннения балок
    Public Shared Function FuncVirtualExtendBearm(ByRef lineBearm As DwgLine, ByVal startLenght As Double, ByVal endLenght As Double, ByRef startPt As Vector3D, ByRef endPt As Vector3D, Optional round As Integer = 3) As Boolean
        FuncVirtualExtendBearm = False
        Try
            Dim L As Double = lineBearm.Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = lineBearm.StartPoint.X + k1 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y1 As Double = lineBearm.StartPoint.Y + k1 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z1 As Double = lineBearm.StartPoint.Z + k1 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            startPt = New Vector3D(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = lineBearm.EndPoint.X + k2 * (lineBearm.EndPoint.X - lineBearm.StartPoint.X)
            Dim Y2 As Double = lineBearm.EndPoint.Y + k2 * (lineBearm.EndPoint.Y - lineBearm.StartPoint.Y)
            Dim z2 As Double = lineBearm.EndPoint.Z + k2 * (lineBearm.EndPoint.Z - lineBearm.StartPoint.Z)
            endPt = New Vector3D(X2, Y2, z2)
            Return True
        Catch ex As Exception
        End Try
    End Function

    'функция считает зазор по минимальному и максимальному значению
    Public Shared Function FuncCalculateZazorBeams(ByVal dimLeftLenghtStructure As Double, ByVal dimRightLenghtStructure As Double, ByVal offsetRowBeams As Double, ByVal leftZazor As Double, ByVal RightZazor As Double) As Double
        'делаем расчет зазора
        Dim fullDimBridge As Double = dimLeftLenghtStructure + dimRightLenghtStructure
        If RightZazor > leftZazor Then
            Dim deltaZazor As Double = RightZazor - leftZazor
            FuncCalculateZazorBeams = (deltaZazor * (dimLeftLenghtStructure + offsetRowBeams) / (fullDimBridge)) + leftZazor
        ElseIf RightZazor < leftZazor Then
            Dim deltaZazor As Double = leftZazor - RightZazor
            FuncCalculateZazorBeams = (deltaZazor * (dimRightLenghtStructure - offsetRowBeams) / (fullDimBridge)) + RightZazor
        Else
            FuncCalculateZazorBeams = leftZazor
        End If
    End Function

    'функция возвращает высоты (по верху) точек опирания балки
    Public Shared Function FuncReturnTopElevationPointPrBeam(ByVal lineDownBalka As DwgLine, ByVal height As Double, ByVal surf As Surface, ByRef arrayElev As Double()) As Boolean
        FuncReturnTopElevationPointPrBeam = False
        If IsNothing(lineDownBalka) = True Then Exit Function
        If height <= 0 Then Exit Function
        If IsNothing(surf) = True Then Exit Function

        Erase arrayElev
        '1.делаем смещение балки вверх
        Dim b As Double = lineDownBalka.EndPoint.Z - lineDownBalka.StartPoint.Z
        Dim c As Double = lineDownBalka.Length
        Dim i As Double = Math.Asin(b / c)
        Dim wAngle As Double = (i * 180) / Math.PI
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + Math.PI / 2, height)
        'получаем новые координаты верха балки
        Dim startCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineDownBalka.StartPoint.X, lineDownBalka.StartPoint.Y), lineDownBalka.Rotation, deltaXZ.X)
        Dim endCoordUpBearm As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(New Vector2D(lineDownBalka.EndPoint.X, lineDownBalka.EndPoint.Y), lineDownBalka.Rotation, deltaXZ.X)
        'создаем верх балки
        Dim startPointUpBalkaPr As Vector3D = New Vector3D(startCoordUpBearm.X, startCoordUpBearm.Y, lineDownBalka.StartPoint.Z + deltaXZ.Y)
        Dim endPointUpBalkaPr As Vector3D = New Vector3D(endCoordUpBearm.X, endCoordUpBearm.Y, lineDownBalka.EndPoint.Z + deltaXZ.Y)
        Try
            Dim startPointEarth As Double = surf.GetElevation(startCoordUpBearm)
            Dim endPointEarth As Double = surf.GetElevation(endCoordUpBearm)
            ReDim arrayElev(3)
            arrayElev(0) = startPointEarth
            arrayElev(1) = endPointEarth
            arrayElev(2) = startPointUpBalkaPr.Z
            arrayElev(3) = endPointUpBalkaPr.Z
            Return True
        Catch ex As System.Exception
        End Try
    End Function

    'функция возвращает ось балки ближайшей к выбранной точке
    Public Shared Function FuncReturnAxisBeamByPoint(ByRef activDoc As Topomatic.Dwg.Drawing, ByRef axisBeamDictionary As Dictionary(Of String, String(,)), ByVal userPoint As Vector3D, ByRef lineDownBalka As DwgLine, ByRef userPropeties As Beams, ByRef userVertex As Vector3D, ByVal idBridge As String, Optional ByVal maxLenght As Double = 0.5) As Double
        FuncReturnAxisBeamByPoint = 9999999
        If axisBeamDictionary.Count > 0 Then
            If axisBeamDictionary.ContainsKey(idBridge) = True Then
                Dim arrayBeams As String(,) = axisBeamDictionary.Item(idBridge)
                If IsArray(arrayBeams) = True Then
                    For i As Integer = 0 To arrayBeams.GetUpperBound(1)
                        Dim hgObject As UInteger = CUInt(arrayBeams(4, i))
                        Dim acLine As DwgLine = Nothing
                        Dim boolFindObject As Boolean = activDoc.ActiveSpace.Entities.TryGetObject(hgObject, acLine)
                        If boolFindObject = True Then
                            Dim l1 As Double = (userPoint.Pos - acLine.StartPoint.Pos).Length
                            Dim l2 As Double = (userPoint.Pos - acLine.EndPoint.Pos).Length
                            If l1 < maxLenght Or l2 < maxLenght Then
                                Dim arrayData As String(,) = Nothing
                                Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(acLine, "PROJECT_BRIDGE", arrayData)
                                If IsArray(arrayData) Then
                                    Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                                    If nameObject Like "Ось балки" Then
                                        Dim keyParamAxis As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1)
                                        Dim axisBeam As Beams = Nothing
                                        Try
                                            axisBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamAxis)
                                        Catch ex As Newtonsoft.Json.JsonException
                                        End Try
                                        If IsNothing(axisBeam) = False Then
                                            userPropeties = axisBeam
                                            lineDownBalka = acLine
                                            If l1 < maxLenght Then
                                                userVertex = acLine.StartPoint
                                                maxLenght = l1
                                            Else
                                                userVertex = acLine.EndPoint
                                                maxLenght = l2
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        Return maxLenght
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'работа с файлами
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает альбомы балок
    Public Shared Function FuncReadDirectoryAlbumsBeams(ByVal putchAlbum As String, ByRef arrayNameAlbumDirectory As String()) As Boolean
        FuncReadDirectoryAlbumsBeams = False
        If Directory.Exists(putchAlbum) = False Then
            Return False
        End If
        Dim countArrayNameAlbom As Integer = 0
        'проверяем наличие директории
        'получаем директории с альбомами
        Dim allfolders As String() = Directory.GetDirectories(putchAlbum)
        If IsArray(allfolders) = True Then
            For i As Integer = 0 To allfolders.Length - 1
                Dim tempFolder As String = allfolders(i)
                Dim nameAlbum As String = New DirectoryInfo(tempFolder).Name
                '============================================================================================
                '1 ищем файл xml
                Dim fullPatchFiles As String = putchAlbum & nameAlbum & "\"
                Dim xmlBeamsFile As String() = Directory.GetFiles(fullPatchFiles, "*.xml")
                If xmlBeamsFile.Length > 0 Then
                    ReDim Preserve arrayNameAlbumDirectory(countArrayNameAlbom)
                    arrayNameAlbumDirectory(countArrayNameAlbom) = nameAlbum
                    countArrayNameAlbom += 1
                End If
            Next
        End If
        If countArrayNameAlbom > 0 Then
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение альбома балок из файла msExcel
    Public Shared Function FuncReadAlbumBeamsFromExcelFiles(ByVal putchDirectory As String, ByRef dictionaryBearm As Dictionary(Of String, Beams()), ByRef arrayNameBeams As String()) As Boolean
        FuncReadAlbumBeamsFromExcelFiles = False
        Erase arrayNameBeams
        Dim arrayBeams As Beams() = Nothing
        Dim countArrayBeams As Integer = 0
        If IO.Directory.Exists(putchDirectory) = True Then
            Dim nameAlbum As String = New DirectoryInfo(putchDirectory).Name
            'находим файлы xls в директории альбома
            Dim allFilesXLS As String() = Directory.GetFiles(putchDirectory, "*.xlsx")
            If allFilesXLS.Count > 0 Then
                Dim xlApp As Excel.Application = New Excel.Application()
                For j As Integer = 0 To allFilesXLS.Length - 1
                    Dim fileXLS As String = allFilesXLS(j)
                    Dim xlWorkBook As Excel.Workbook = Nothing
                    Dim xlWorkSheet As Excel.Worksheet = Nothing
                    xlWorkBook = xlApp.Workbooks.Open(fileXLS, 0, True, 5, "", "", True, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", False, False, 0, True, 1, 0)
                    xlWorkSheet = xlWorkBook.Worksheets.Item(1)
                    If xlWorkSheet IsNot Nothing Then
                        For indRows As Integer = 2 To xlWorkSheet.Rows.Count - 1
                            Dim userBeam As Beams = New Beams
                            Dim model As String = xlWorkSheet.Cells(indRows, 1).value
                            If IsNothing(model) = True Then
                                Exit For
                            ElseIf model.Trim.Length = 0 Then
                                Exit For
                            Else
                                userBeam.model = model
                                userBeam.nameAlbum = nameAlbum
                                userBeam.fullLenght = xlWorkSheet.Cells(indRows, 2).value
                                userBeam.widthTop = xlWorkSheet.Cells(indRows, 3).value
                                userBeam.widthBottom = xlWorkSheet.Cells(indRows, 4).value
                                userBeam.height = xlWorkSheet.Cells(indRows, 5).value
                                userBeam.a = xlWorkSheet.Cells(indRows, 6).value
                                userBeam.b = xlWorkSheet.Cells(indRows, 7).value
                                userBeam.mass = xlWorkSheet.Cells(indRows, 8).value
                                userBeam.concrete = xlWorkSheet.Cells(indRows, 9).value
                                userBeam.volume = xlWorkSheet.Cells(indRows, 10).value
                                userBeam.steel = xlWorkSheet.Cells(indRows, 11).value
                                userBeam.modelTLS = xlWorkSheet.Cells(indRows, 12).value
                                ReDim Preserve arrayBeams(countArrayBeams)
                                arrayBeams(countArrayBeams) = userBeam
                                countArrayBeams += 1
                            End If
                        Next indRows
                        If IsArray(arrayBeams) = True Then
                            If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                                dictionaryBearm.Remove(nameAlbum)
                            End If
                            dictionaryBearm.Add(nameAlbum, arrayBeams)
                            For k As Integer = 0 To arrayBeams.Length - 1
                                Dim userBeam As Beams = arrayBeams(k)
                                ReDim Preserve arrayNameBeams(k)
                                arrayNameBeams(k) = userBeam.model
                            Next k
                        End If
                    End If
                    xlWorkBook.Close(True, Nothing, Nothing)
                Next j
                xlApp.Quit()
            End If
        End If
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение альбома балок из файла csv
    Public Shared Function FuncReadAlbumBeamsFromCSVFiles(ByVal putchDirectory As String, ByRef dictionaryBearm As Dictionary(Of String, Beams()), ByRef arrayNameBeams As String()) As Boolean
        FuncReadAlbumBeamsFromCSVFiles = False
        Erase arrayNameBeams
        Dim arrayBeams As Beams() = Nothing
        Dim countArrayBeams As Integer = 0
        If IO.Directory.Exists(putchDirectory) = True Then
            Dim nameAlbum As String = New DirectoryInfo(putchDirectory).Name
            'находим файлы xls в директории альбома
            Dim allFilesCSV As String() = Directory.GetFiles(putchDirectory, "*.csv")
            If allFilesCSV.Count > 0 Then
                For j As Integer = 0 To allFilesCSV.Length - 1
                    Dim fileCSV As String = allFilesCSV(j)
                    Dim input As StreamReader = New StreamReader(fileCSV, System.Text.Encoding.Default)
                    Dim count As Integer = 0
                    Do Until input.EndOfStream
                        Dim line1 As String = input.ReadLine() 'считываем строку
                        If count = 0 Then
                            count += 1
                            Continue Do
                        End If
                        line1 = line1.Trim
                        If line1.Trim.Length = 0 Then Continue Do
                        Dim arrayStr As String() = line1.Split(";")
                        If arrayStr.Length > 10 Then
                            Dim userBeam As Beams = New Beams
                            userBeam.model = arrayStr(0)
                            userBeam.nameAlbum = nameAlbum
                            userBeam.fullLenght = arrayStr(1)
                            userBeam.widthTop = arrayStr(2)
                            userBeam.widthBottom = arrayStr(3)
                            userBeam.height = arrayStr(4)
                            userBeam.a = arrayStr(5)
                            userBeam.b = arrayStr(6)
                            userBeam.mass = arrayStr(7)
                            userBeam.concrete = arrayStr(8)
                            userBeam.volume = arrayStr(9)
                            userBeam.steel = arrayStr(10)
                            userBeam.modelTLS = arrayStr(11)
                            ReDim Preserve arrayBeams(countArrayBeams)
                            arrayBeams(countArrayBeams) = userBeam
                            countArrayBeams += 1
                        End If
                        count += 1
                    Loop
                    If IsArray(arrayBeams) = True Then
                        If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                            dictionaryBearm.Remove(nameAlbum)
                        End If
                        dictionaryBearm.Add(nameAlbum, arrayBeams)
                        For k As Integer = 0 To arrayBeams.Length - 1
                            Dim userBeam As Beams = arrayBeams(k)
                            ReDim Preserve arrayNameBeams(k)
                            arrayNameBeams(k) = userBeam.model
                        Next k
                    End If
                    input.Close()
                Next j
            End If
        End If
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'конвертирование альбома балок из формата xls в формат xml
    Public Shared Function FuncCreateXMLBeamfByExcel(ByVal putchXLSFiles As String(), ByVal putchFileXML As String) As Boolean
        FuncCreateXMLBeamfByExcel = False
        Dim xws As XmlWriterSettings = New XmlWriterSettings()
        xws.Indent = True
        xws.NewLineOnAttributes = True
        Using xw As XmlWriter = XmlWriter.Create(putchFileXML, xws)
            xw.WriteStartDocument()
            xw.WriteStartElement("Beams")
            If IsArray(putchXLSFiles) = True Then
                Dim xlApp As Excel.Application = New Excel.Application()
                For i As Integer = 0 To putchXLSFiles.Length - 1
                    Dim fileXLS As String = putchXLSFiles(i)
                    If IO.File.Exists(fileXLS) = True Then
                        Dim xlWorkBook As Excel.Workbook = Nothing
                        Dim xlWorkSheet As Excel.Worksheet = Nothing
                        Try
                            xlWorkBook = xlApp.Workbooks.Open(fileXLS,, True)
                            xlWorkSheet = xlWorkBook.Worksheets.Item(1)
                            If xlWorkSheet IsNot Nothing Then
                                For indRows As Integer = 2 To xlWorkSheet.Rows.Count - 1
                                    For indColl As Integer = 1 To 100
                                        Dim nameField As String = xlWorkSheet.Cells(1, indColl).value
                                        Dim value As String = xlWorkSheet.Cells(indRows, indColl).value
                                        If IsNothing(nameField) = True Then
                                            Exit For
                                        End If
                                        If nameField.Trim.Length = 0 Then
                                            Exit For
                                        End If
                                        nameField = nameField.Replace(" ", "_")
                                        If IsNothing(value) = True Then
                                            value = ""
                                        End If
                                        If indColl = 1 And value.Trim.Length = 0 Then
                                            Exit For
                                        End If
                                        If indColl = 1 Then
                                            xw.WriteStartElement("Beam")
                                        End If
                                        xw.WriteAttributeString(nameField, value)
                                    Next indColl
                                    xw.WriteEndElement()
                                Next indRows
                            End If
                        Catch ex As Exception
                            If IsNothing(xlWorkBook) = False Then
                                xlWorkBook.Close(True)
                                xlWorkBook = Nothing
                            End If
                            If IsNothing(xlApp) = False Then
                                xlApp.Quit()
                            End If
                            Return False
                        Finally
                            If IsNothing(xlWorkBook) = False Then
                                xlWorkBook.Close(True)
                            End If
                        End Try
                    End If
                Next i
                If IsNothing(xlApp) = False Then
                    xlApp.Quit()
                End If
                xw.WriteEndElement()
                xw.WriteEndDocument()
                xw.Flush()
                xw.Close()
            End If
        End Using
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение альбома балок из файла xml (
    Public Shared Function FuncReadAlbumBeamsFromXMLFiles(ByVal fullPatchXML As String, ByRef dictionaryBearm As Dictionary(Of String, Beams()), ByRef arrayNameBeams As String(), Optional nameAlbum As String = "") As Boolean
        FuncReadAlbumBeamsFromXMLFiles = False
        Dim countArrayBeams As Integer = 0
        If IsArray(arrayNameBeams) = True Then
            If arrayNameBeams.Length > 0 Then
                countArrayBeams = arrayNameBeams.Length
            End If
        End If
        Dim countBeams As Integer = 0
        Dim arrayBeams As Beams() = Nothing
        If dictionaryBearm.Count > 0 Then
            If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                arrayBeams = dictionaryBearm.Item(nameAlbum)
                countBeams = arrayBeams.Length
            End If
        End If
        If IO.File.Exists(fullPatchXML) = True Then
            Dim xDoc As XmlDocument = New XmlDocument()
            Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Beam" Then
                            If reader.HasAttributes = True Then
                                Dim userBeam As Beams = New Beams
                                While (reader.MoveToNextAttribute())
                                    Dim name As String = reader.Name
                                    Dim valN As String = reader.Value
                                    If name Like "model" Then
                                        userBeam.model = valN
                                        ReDim Preserve arrayNameBeams(countArrayBeams)
                                        arrayNameBeams(countArrayBeams) = userBeam.model
                                        countArrayBeams += 1
                                    ElseIf name Like "fullLenght" Then
                                        userBeam.fullLenght = Val(valN)
                                    ElseIf name Like "widthTop" Then
                                        userBeam.widthTop = Val(valN)
                                    ElseIf name Like "widthBottom" Then
                                        userBeam.widthBottom = Val(valN)
                                    ElseIf name Like "height" Then
                                        userBeam.height = Val(valN)
                                    ElseIf name Like "a" Then
                                        userBeam.a = Val(valN)
                                    ElseIf name Like "b" Then
                                        userBeam.b = Val(valN)
                                    ElseIf name Like "mass" Then
                                        userBeam.mass = Val(valN)
                                    ElseIf name Like "concrete" Then
                                        userBeam.concrete = valN
                                    ElseIf name Like "volume" Then
                                        userBeam.volume = valN
                                    ElseIf name Like "steel" Then
                                        userBeam.steel = valN
                                    ElseIf name Like "modelTLS" Then
                                        userBeam.modelTLS = valN
                                    End If
                                End While
                                userBeam.nameAlbum = nameAlbum
                                ReDim Preserve arrayBeams(countBeams)
                                arrayBeams(countBeams) = userBeam
                                countBeams += 1
                            End If
                        End If
                End Select
            End While
            reader.Close()
            If IsArray(arrayBeams) = True Then
                If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                    dictionaryBearm.Remove(nameAlbum)
                End If
                dictionaryBearm.Add(nameAlbum, arrayBeams)
            End If
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение альбома балок из файла xml (
    Public Shared Function FuncReadNameBeamsFromXMLFiles(ByVal fullPatchXML As String, ByRef arrayNameBeams As String()) As Boolean
        FuncReadNameBeamsFromXMLFiles = False
        Dim countArrayNameBeams As Integer = 0
        If IsArray(arrayNameBeams) = True Then
            If arrayNameBeams.Length > 0 Then
                countArrayNameBeams = arrayNameBeams.Length
            End If
        End If
        If IO.File.Exists(fullPatchXML) = True Then
            Dim xDoc As XmlDocument = New XmlDocument()
            Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Beam" Then
                            If reader.HasAttributes = True Then
                                Dim userBeam As Beams = New Beams
                                While (reader.MoveToNextAttribute())
                                    Dim name As String = reader.Name
                                    Dim valN As String = reader.Value
                                    If name Like "model" Then
                                        userBeam.model = valN
                                        ReDim Preserve arrayNameBeams(countArrayNameBeams)
                                        arrayNameBeams(countArrayNameBeams) = userBeam.model
                                        countArrayNameBeams += 1
                                        Exit While
                                    End If
                                End While
                            End If
                        End If
                End Select
            End While
            reader.Close()
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция водвращает ось объекта по его 3д модели
    Public Shared Function FuncFindAxisLineBy3dElements(ByVal acModel3d As DwgModel3DElement) As DwgLine
        FuncFindAxisLineBy3dElements = Nothing
        If IsNothing(acModel3d) = False Then
            Dim arrayData As String(,) = {}
            Dim boolDataPs1 As Boolean = FuncXRecords.FuncReadXData(acModel3d, "PROJECT_BRIDGE", arrayData)
            If arrayData.Length > 0 Then
                Dim localId As String = MathFunction.FuncFindValueToArray2d(arrayData, "ElementID", 0, 1)
                Dim ActivDocument As Topomatic.Dwg.Drawing = acModel3d.Drawing
                If IsNothing(ActivDocument) = False Then
                    For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                        If TypeOf (acEnt) Is DwgLine Then
                            Dim acLine As DwgLine = acEnt
                            Dim idElement As String = FuncXRecords.FuncReadValueXData(acLine, "ElementID", "PROJECT_BRIDGE")
                            If idElement Like localId Then
                                Return acLine
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает имена объектов
    Public Shared Function FuncReturnBridgeObject(ByVal activeDoc As Topomatic.Dwg.Drawing, ByRef brigeGeneralAxisDictionary As Dictionary(Of String, String())) As Boolean
        brigeGeneralAxisDictionary = New Dictionary(Of String, String())
        FuncReturnBridgeObject = False
        If IsNothing(activeDoc) = False Then
            For Each acEnt As DwgEntity In activeDoc.ActiveSpace.Entities
                If TypeOf acEnt Is DwgPolyline Then
                    Dim acAxisBridge As DwgPolyline = acEnt
                    Dim arrayData As String(,) = {}
                    Dim boolReadXdata As Boolean = FuncXRecords.FuncReadXData(acAxisBridge, "PROJECT_BRIDGE", arrayData)
                    If arrayData.Length > 0 Then
                        Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                        If nameObject Like "Главная ось сооружения" Or nameObject Like "главная ось сооружения" Then
                            Dim idBridge As String = MathFunction.FuncFindValueToArray2d(arrayData, "BrigeID", 0, 1)
                            Dim keyParam As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1)
                            Dim idElement As String = MathFunction.FuncFindValueToArray2d(arrayData, "ElementID", 0, 1)
                            If brigeGeneralAxisDictionary.Count > 0 Then
                                If brigeGeneralAxisDictionary.ContainsKey(idBridge) = False Then
                                    Dim arrayWrite As String() = Nothing
                                    ReDim arrayWrite(4)
                                    arrayWrite(0) = nameObject
                                    arrayWrite(1) = keyParam
                                    arrayWrite(2) = idElement
                                    arrayWrite(3) = idBridge
                                    arrayWrite(4) = acAxisBridge.ObjectID
                                    brigeGeneralAxisDictionary.Add(idBridge, arrayWrite)
                                End If
                            Else
                                Dim arrayWrite As String() = Nothing
                                ReDim arrayWrite(4)
                                arrayWrite(0) = nameObject
                                arrayWrite(1) = keyParam
                                arrayWrite(2) = idElement
                                arrayWrite(3) = idBridge
                                arrayWrite(4) = acAxisBridge.ObjectID
                                brigeGeneralAxisDictionary.Add(idBridge, arrayWrite)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If brigeGeneralAxisDictionary.Count > 0 Then
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение характеристик определенной балки из файла xml
    Public Shared Function FuncReadBeamsFromXMLFiles(ByVal fullPatchXML As String, ByVal model As String) As Beams
        FuncReadBeamsFromXMLFiles = Nothing
        Try
            If IO.File.Exists(fullPatchXML) = True Then
                Dim xDoc As XmlDocument = New XmlDocument()
                Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "Beam" Then
                                If reader.HasAttributes = True Then
                                    Dim userBeam As Beams = New Beams
                                    Dim findModel As Boolean = False
                                    While (reader.MoveToNextAttribute())
                                        Dim name As String = reader.Name
                                        Dim valN As String = reader.Value
                                        If name Like "model" And valN Like model Then
                                            userBeam.model = valN
                                            findModel = True
                                        ElseIf name Like "fullLenght" And findModel = True Then
                                            userBeam.fullLenght = Val(valN) / 1000
                                        ElseIf name Like "widthTop" And findModel = True Then
                                            userBeam.widthTop = Val(valN) / 1000
                                        ElseIf name Like "widthBottom" And findModel = True Then
                                            userBeam.widthBottom = Val(valN) / 1000
                                        ElseIf name Like "height" And findModel = True Then
                                            userBeam.height = Val(valN) / 1000
                                        ElseIf name Like "a" And findModel = True Then
                                            userBeam.a = Val(valN) / 1000
                                        ElseIf name Like "b" And findModel = True Then
                                            userBeam.b = Val(valN) / 1000
                                        ElseIf name Like "mass" And findModel = True Then
                                            userBeam.mass = Val(valN)
                                        ElseIf name Like "concrete" And findModel = True Then
                                            userBeam.concrete = valN
                                        ElseIf name Like "volume" And findModel = True Then
                                            userBeam.volume = valN
                                        ElseIf name Like "steel" And findModel = True Then
                                            userBeam.steel = valN
                                        ElseIf name Like "modelTLS" And findModel = True Then
                                            userBeam.modelTLS = valN
                                        End If
                                    End While
                                    If findModel = True Then
                                        reader.Close()
                                        Return userBeam
                                    End If
                                End If
                            End If
                    End Select
                End While
                reader.Close()
                Return Nothing
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'запись XRecords в мостовую балку
    Public Shared Function FuncWriteXRecordsToBeam(ByRef acLineBeam As DwgLine, ByVal userBeam As Beams, ByVal idBridge As String, ByVal templateXML As String) As Boolean
        FuncWriteXRecordsToBeam = False
        If IsNothing(userBeam) = False Then
            Dim strJSON As String = ""
            Try
                strJSON = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
            Catch ex As Newtonsoft.Json.JsonException
            End Try
            If IsNothing(acLineBeam) = False Then
                Dim boolFindTable As Boolean = FuncXRecords.FuncFindTablePS(acLineBeam, "PROJECT_BRIDGE")
                Dim arrayRecAxis As String(,) = {}
                If boolFindTable = True Then
                    ReDim Preserve arrayRecAxis(1, 0)
                    arrayRecAxis(0, 0) = "KeyParameters"
                    arrayRecAxis(1, 0) = strJSON
                Else
                    ReDim Preserve arrayRecAxis(1, 5)
                    arrayRecAxis(0, 0) = "LABEL"
                    arrayRecAxis(1, 0) = "Мосты и путепроводы"

                    arrayRecAxis(0, 1) = "Name"
                    arrayRecAxis(1, 1) = "Ось балки"

                    arrayRecAxis(0, 2) = "KeyParameters"
                    arrayRecAxis(1, 2) = strJSON

                    Dim uniqueId As Guid = Guid.NewGuid()
                    Dim rng As Random = New Random
                    Dim idBeam As String = rng.Next
                    arrayRecAxis(0, 3) = "ElementID"
                    arrayRecAxis(1, 3) = uniqueId.ToString

                    arrayRecAxis(0, 4) = "BrigeID"
                    arrayRecAxis(1, 4) = idBridge

                    arrayRecAxis(0, 5) = "NOTE"
                    arrayRecAxis(1, 5) = "-"
                End If
                If IsArray(arrayRecAxis) = True Then
                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(acLineBeam, "PROJECT_BRIDGE", arrayRecAxis, templateXML)
                    Return True
                End If
            End If
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'перераскладка балок мостового сооружения
    Public Shared Sub PlacementUserBeams(ByRef activDoc As Topomatic.Dwg.Drawing, ByVal arrayPropertiesBridge As String(), ByVal arrayPropertiesRows As Double(,), ByVal dictionaryBeams As Dictionary(Of Integer, String(,)), ByVal userAlign As Alignment, ByVal userSurface As Surface, ByRef elementDictionary As Dictionary(Of String, UInteger()), ByRef hgBoundares As UInteger, ByVal putchBridge As String, ByVal templateXML As String)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'оформление
        'балка
        Dim boolBeamStyle As Boolean = False
        Dim layerBeam As DwgLayer = activDoc.ActiveLayer
        Dim colorBeam As CadColor = New CadColor(7)
        Dim nameTypeLineBeam As DwgLinetype = activDoc.ActiveLinetype
        Dim ScaleTypeLineBeam As Integer = 0.1
        Dim widthTypeLineBeam As Integer = 20
        Dim tablePSBeam As String = ""
        'Верх ребра балки
        Dim boolUpBeamStyle As Boolean = False
        Dim layerUpBeam As DwgLayer = activDoc.ActiveLayer
        Dim colorUpBeam As CadColor = New CadColor(7)
        Dim nameTypeLineUpBeam As DwgLinetype = activDoc.ActiveLinetype
        Dim ScaleTypeLineUpBeam As Integer = 1
        Dim widthTypeLineUpBeam As Integer = 20
        Dim tablePSUpBeam As String = ""
        'Низ ребра балки
        Dim boolDownBeamStyle As Boolean = False
        Dim layerDownBeam As DwgLayer = activDoc.ActiveLayer
        Dim colorDownBeam As CadColor = New CadColor(7)
        Dim nameTypeLineDownBeam As DwgLinetype = activDoc.ActiveLinetype
        Dim ScaleTypeLineDownBeam As Integer = 1
        Dim widthTypeLineDownBeam As Integer = 20
        Dim tablePSDownBeam As String = ""
        'находим стили оформления вспомогательных элементов
        If elementDictionary.Count > 0 Then
            For i As Integer = 0 To elementDictionary.Count - 1
                Dim keyPair As KeyValuePair(Of String, UInteger()) = elementDictionary.ElementAt(i)
                Dim valObj As UInteger() = keyPair.Value
                If IsArray(valObj) = True Then
                    For j As Integer = 0 To valObj.Length - 1
                        Dim idObj As UInteger = valObj(j)
                        Dim entObj As DwgObject = Nothing
                        Dim boolFindObject As Boolean = activDoc.ActiveSpace.Entities.TryGetObject(idObj, entObj)
                        If boolFindObject = True Then
                            Dim arrayData As String(,) = Nothing
                            Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(entObj, "PROJECT_BRIDGE", arrayData)
                            If IsArray(arrayData) Then
                                Dim nameElement As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                                If nameElement Like "Балка" Then
                                    If boolBeamStyle = False Then
                                        If TypeOf entObj Is DwgModel3DElement Then
                                            Dim beamModel3d As DwgModel3DElement = entObj
                                            layerBeam = beamModel3d.Layer
                                            colorBeam = beamModel3d.Color
                                            nameTypeLineBeam = beamModel3d.Linetype
                                            ScaleTypeLineBeam = beamModel3d.LinetypeScale
                                            widthTypeLineBeam = beamModel3d.Lineweight
                                            boolBeamStyle = True
                                        End If
                                    End If
                                ElseIf nameElement Like "Верх грани плиты балки" Then
                                    If boolUpBeamStyle = False Then
                                        If TypeOf entObj Is DwgLine Then
                                            Dim beamLine As DwgLine = entObj
                                            layerUpBeam = beamLine.Layer
                                            colorUpBeam = beamLine.Color
                                            nameTypeLineUpBeam = beamLine.Linetype
                                            ScaleTypeLineUpBeam = beamLine.LinetypeScale
                                            widthTypeLineUpBeam = beamLine.Lineweight
                                            boolUpBeamStyle = True
                                        End If
                                    End If
                                ElseIf nameElement Like "Низ ребра балки" Then
                                    If boolDownBeamStyle = False Then
                                        If TypeOf entObj Is DwgLine Then
                                            Dim beamLine As DwgLine = entObj
                                            layerDownBeam = beamLine.Layer
                                            colorDownBeam = beamLine.Color
                                            nameTypeLineDownBeam = beamLine.Linetype
                                            ScaleTypeLineDownBeam = beamLine.LinetypeScale
                                            widthTypeLineDownBeam = beamLine.Lineweight
                                            boolDownBeamStyle = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
                If boolBeamStyle = True And boolDownBeamStyle = True And boolUpBeamStyle = True Then
                    Exit For
                End If
            Next i
        End If
        'Ось опоры
        Dim boolAxisPillar As Boolean = False
        Dim layerAxisPillar As DwgLayer = activDoc.ActiveLayer
        Dim colorAxisPillar As CadColor = New CadColor(7)
        Dim nameTypeLineAxisPillar As DwgLinetype = activDoc.ActiveLinetype
        Dim ScaleTypeLineAxisPillar As Integer = 1
        Dim widthTypeLineAxisPillar As Integer = 20
        Dim tablePSAxisPillar As String = ""
        'Ось опирания балок
        Dim boolAxisBeamPillar As Boolean = False
        Dim layerAxisBeamPillar As DwgLayer = activDoc.ActiveLayer
        Dim colorAxisBeamPillar As CadColor = New CadColor(7)
        Dim nameTypeLineAxisBeamPillar As DwgLinetype = activDoc.ActiveLinetype
        Dim ScaleTypeLineAxisBeamPillar As Integer = 1
        Dim widthTypeLineAxisBeamPillar As Integer = 20
        Dim tablePSAxisBeamPillar As String = ""
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim rng As Random = New Random
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'параметры сооружения
        If IsArray(arrayPropertiesBridge) = False Then
            Exit Sub
        ElseIf arrayPropertiesBridge.Length < 6 Then
            Exit Sub
        End If
        Dim nameBridge As String = arrayPropertiesBridge(0) 'имя сооружения
        Dim idBridge As String = arrayPropertiesBridge(1) 'ID Сооружения
        Dim keyParam As String = arrayPropertiesBridge(2) 'параметры
        Dim poperOffset As Double = Val(arrayPropertiesBridge(5)) 'поперечное смещение
        Dim hgBridge As UInteger = CUInt(arrayPropertiesBridge(6)) 'ObjectID оси
        Dim startPositionBridge As Double = Val(arrayPropertiesBridge(7)) 'пикет начала раскладки балок
        Dim userBridge As Bridge = Nothing
        Try
            userBridge = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Bridge)(keyParam)
        Catch ex As Newtonsoft.Json.JsonException
            Exit Sub
        End Try
        'размер сооружения
        Dim dimBridgeLeft As Double = userBridge.dimLeftStructure
        Dim dimBrigeRight As Double = userBridge.dimRightStructure
        'горизонтальное и вертикальные смещения
        Dim horizontalOffset As Double = userBridge.offsetHPosition 'горизонтальное смещение
        Dim verticalOffset As Double = userBridge.offsetVPosition 'вертикальное смещение
        'доступ к оси сооружения
        Dim axisOldAlign As DwgPolyline = Nothing
        Dim boolPline As Boolean = activDoc.ActiveSpace.Entities.TryGetObject(hgBridge, axisOldAlign)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim acStartLine As DwgLine = Nothing 'ось первой опоры
        Dim acEndLine As DwgLine = Nothing 'ось последней опоры
        Dim arrayLineElements As DwgLine() = {}
        Dim countAtrrayLineElements As Integer = 0
        'проверка ости трассы и поверхности
        If IsNothing(userAlign) = True Then
            MsgBox("Трасса не найдена. Макро прервано!!!")
            Exit Sub
        End If
        If IsNothing(userSurface) = True Then
            MsgBox("Поверхность не найдена. Макро прервано!!!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'Преобразовываем трассу в полилинию
        Dim axisPline0 As DwgPolyline = Nothing 'полилиния из трассы  в прямом направлении
        Dim acPoly3dAlign As Polyline3D = New Polyline3D
        userAlign.Plan.CompoundLine.ToPolyLine(acPoly3dAlign)
        Dim ArrayCoord As Double(,) = Nothing
        Dim countarrayPos As Integer = 0
        For j As Integer = 0 To acPoly3dAlign.Count - 1
            Dim pos As BugleVector3D = acPoly3dAlign.Item(j)
            ReDim Preserve ArrayCoord(3, countarrayPos)
            ArrayCoord(0, countarrayPos) = pos.Vertex.X
            ArrayCoord(1, countarrayPos) = pos.Vertex.Y
            ArrayCoord(2, countarrayPos) = pos.Bugle
            countarrayPos += 1
        Next
        If IsArray(ArrayCoord) = True Then
            If ArrayCoord.GetUpperBound(1) > 0 Then
                axisPline0 = RoburFunc.FuncDrawPolylineToArrayCoord(activDoc, ArrayCoord)
            End If
        End If
        axisPline0.GetPolyline(acPoly3dAlign)
        'прямое направление в полилинию
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.FuncOffsetAlignment(activDoc, userAlign, -1 * dimBridgeLeft)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)
        'обратное направление в полилинию
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.FuncOffsetAlignment(activDoc, userAlign, -1 * dimBrigeRight, True)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смешение самых крайних осей (лево и право)
        Dim numberLeftRow As Integer = arrayPropertiesRows(0, 0)
        Dim offsetLeftLastBearm As Double = arrayPropertiesRows(1, 0)

        Dim numberRightRow As Integer = arrayPropertiesRows(0, arrayPropertiesRows.GetUpperBound(1))
        Dim offsetRightLastBearm As Double = arrayPropertiesRows(1, arrayPropertiesRows.GetUpperBound(1))
        '=============================================================================================================
        'формируем новый массив с рядами балок (обрабатываем 1 и последний ряд в первую очередь, потом все остальные)
        Dim newArrayPropertiesRows As Double(,) = {}
        ReDim newArrayPropertiesRows(2, arrayPropertiesRows.GetUpperBound(1))
        newArrayPropertiesRows(0, 0) = numberLeftRow
        newArrayPropertiesRows(1, 0) = offsetLeftLastBearm
        newArrayPropertiesRows(2, 0) = arrayPropertiesRows(2, 0)
        newArrayPropertiesRows(0, 1) = numberRightRow
        newArrayPropertiesRows(1, 1) = offsetRightLastBearm
        newArrayPropertiesRows(2, 1) = arrayPropertiesRows(2, arrayPropertiesRows.GetUpperBound(1))
        For i As Integer = 1 To arrayPropertiesRows.GetUpperBound(1) - 1
            newArrayPropertiesRows(0, i + 1) = arrayPropertiesRows(0, i)
            newArrayPropertiesRows(1, i + 1) = arrayPropertiesRows(1, i)
            newArrayPropertiesRows(2, i + 1) = arrayPropertiesRows(2, i)
        Next i
        '==============================================================================================================
        'ищем оси опор
        Dim dictionaryAxisPillars As Dictionary(Of Integer, DwgLine()) = New Dictionary(Of Integer, DwgLine())
        For Each acEnt As DwgEntity In activDoc.ActiveSpace.Entities
            If TypeOf (acEnt) Is DwgLine Then
                Dim acLine As DwgLine = acEnt
                Dim arrayData As String(,) = Nothing
                Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(acLine, "PROJECT_BRIDGE", arrayData)
                If IsArray(arrayData) Then
                    Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                    If nameObject Like "Ось опоры" Then
                        Dim keyParamAxis As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1)
                        Dim axisPillars As AxisHorizontalPillars = Nothing
                        Try
                            axisPillars = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisHorizontalPillars)(keyParamAxis)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                        If IsNothing(axisPillars) = False Then
                            If dictionaryAxisPillars.Count = 0 Then
                                Dim arrayLine As DwgLine() = {Nothing, acLine, Nothing}
                                dictionaryAxisPillars.Add(axisPillars.number, arrayLine)
                            ElseIf dictionaryAxisPillars.ContainsKey(axisPillars.number) = True Then
                                Dim arrayLine As DwgLine() = dictionaryAxisPillars.Item(axisPillars.number)
                                arrayLine(1) = acLine
                                dictionaryAxisPillars.Item(axisPillars.number) = arrayLine
                            Else
                                Dim arrayLine As DwgLine() = {Nothing, acLine, Nothing}
                                dictionaryAxisPillars.Add(axisPillars.number, arrayLine)
                            End If
                            'запоминаем стиль оформления
                            If boolAxisPillar = False Then
                                layerAxisPillar = acLine.Layer
                                colorAxisPillar = acLine.Color
                                nameTypeLineAxisPillar = acLine.Linetype
                                ScaleTypeLineAxisPillar = acLine.LinetypeScale
                                widthTypeLineAxisPillar = acLine.Lineweight
                                boolAxisPillar = True
                            End If
                        End If
                    End If
                End If
            End If
        Next
        'ищем оси опирания балок
        For Each acEnt As DwgEntity In activDoc.ActiveSpace.Entities
            If TypeOf (acEnt) Is DwgLine Then
                Dim acLine As DwgLine = acEnt
                Dim arrayData As String(,) = Nothing
                Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(acLine, "PROJECT_BRIDGE", arrayData)
                If IsArray(arrayData) Then
                    Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                    If nameObject Like "Ось опирания балок" Or nameObject Like "ось опирания балок" Then
                        Dim keyParamAxisBeams As String = MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1)
                        'если это старый вариант программы и в параметре записано только число идем по этой ветке
                        If IsNumeric(keyParamAxisBeams) = True Then
                            Dim numberPillars As Integer = Val(MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1))
                            If dictionaryAxisPillars.ContainsKey(numberPillars) = True Then
                                Dim arrayLine As DwgLine() = dictionaryAxisPillars.Item(numberPillars)
                                Dim axisPillars As DwgLine = arrayLine(1)
                                'пересечение оси опоры
                                Dim pointIntersectCollectionA As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, axisPillars.StartPoint.Pos, axisPillars.EndPoint.Pos)
                                Dim PKLine1 As Double = -1
                                Dim offLine1 As Double = -1
                                Dim boolPK1 As Boolean = False
                                If pointIntersectCollectionA.Count > 0 Then
                                    Dim pointI1 As Vector2D = pointIntersectCollectionA.ElementAt(0)
                                    boolPK1 = userAlign.Plan.CompoundLine.PosToStaOffset(pointI1, PKLine1, offLine1)
                                End If
                                'пересечение оси опирания балок
                                Dim pointIntersectCollectionB As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, acLine.StartPoint.Pos, acLine.EndPoint.Pos)
                                Dim PKLine2 As Double = -1
                                Dim offLine2 As Double = -1
                                Dim boolPK2 As Boolean = False
                                If pointIntersectCollectionB.Count > 0 Then
                                    Dim pointI2 As Vector2D = pointIntersectCollectionB.ElementAt(0)
                                    boolPK2 = userAlign.Plan.CompoundLine.PosToStaOffset(pointI2, PKLine2, offLine2)
                                End If
                                If boolPK1 = True And boolPK2 = True Then
                                    If PKLine1 > PKLine2 Then
                                        arrayLine(0) = acLine
                                    Else
                                        arrayLine(2) = acLine
                                    End If
                                End If
                                dictionaryAxisPillars.Item(numberPillars) = arrayLine
                            End If
                        Else
                            'читаем номер опоры и номер пролета
                            Dim axisPrBeams As AxisBeamsPillar = Nothing
                            Try
                                axisPrBeams = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisBeamsPillar)(keyParamAxisBeams)
                            Catch ex As Newtonsoft.Json.JsonException
                            End Try
                            If IsNothing(axisPrBeams) = False Then
                                Dim numPillars As Integer = axisPrBeams.numberPillar
                                Dim numProlet As Integer = axisPrBeams.numberColl
                                If dictionaryAxisPillars.ContainsKey(numPillars) = True Then
                                    Dim arrayLine As DwgLine() = dictionaryAxisPillars.Item(numPillars)
                                    If numProlet = numPillars Then
                                        arrayLine(2) = acLine
                                    Else
                                        arrayLine(0) = acLine
                                    End If
                                End If
                            End If
                        End If
                        'запоминаем стиль оформления
                        If boolAxisBeamPillar = False Then
                            layerAxisBeamPillar = acLine.Layer
                            colorAxisBeamPillar = acLine.Color
                            nameTypeLineAxisBeamPillar = acLine.Linetype
                            ScaleTypeLineAxisBeamPillar = acLine.LinetypeScale
                            widthTypeLineAxisBeamPillar = acLine.Lineweight
                            boolAxisBeamPillar = True
                        End If
                    End If
                End If
            End If
        Next
        If dictionaryAxisPillars.Count > 1 Then
            dictionaryAxisPillars = dictionaryAxisPillars.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
            Dim keyValStart As KeyValuePair(Of Integer, DwgLine()) = dictionaryAxisPillars.ElementAt(0)
            Dim arrayStartElement As DwgLine() = keyValStart.Value
            acStartLine = arrayStartElement(1)
            Dim keyValEnd As KeyValuePair(Of Integer, DwgLine()) = dictionaryAxisPillars.ElementAt(dictionaryAxisPillars.Count - 1)
            Dim arrayEndElement As DwgLine() = keyValEnd.Value
            acEndLine = arrayEndElement(1)
        End If
        '========================================================================================
        'создаем словарь с зазорами опор
        Dim dictionaryPillarAxisZazor As Dictionary(Of Integer, Double()) = New Dictionary(Of Integer, Double())
        For i As Integer = 0 To dictionaryAxisPillars.Count - 1
            Dim keyVal As KeyValuePair(Of Integer, DwgLine()) = dictionaryAxisPillars.ElementAt(i)
            Dim numberPillar As Integer = keyVal.Key
            Dim acLinePillarAxis As DwgLine = keyVal.Value(1)
            If IsNothing(acLinePillarAxis) = False Then
                Dim keyParamAxis As String = FuncXRecords.FuncReadValueXData(acLinePillarAxis, "KeyParameters", "PROJECT_PILLARS")
                If keyParamAxis.Length > 0 Then
                    Dim userPillars As AxisHorizontalPillars = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisHorizontalPillars)(keyParamAxis)
                    dictionaryPillarAxisZazor.Add(numberPillar, {userPillars.leftClearence, userPillars.rightClearence})
                End If
            End If
        Next
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'перемещаем первую опору (если необходимо мост сдвинуть относительно оси
        'ищем точку пересечения с трассой
        'находим точку пересечения трассы смещения и сечения
        Dim pointIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, acStartLine.StartPoint.Pos, acStartLine.EndPoint.Pos)
        If pointIntersectCollection.Count > 0 Then
            Dim tempStartPoint As Vector2D = pointIntersectCollection(0)
            Dim angleStart As Double = (acStartLine.StartPoint.Pos - tempStartPoint).Angle
            Dim distStart As Double = (acStartLine.StartPoint.Pos - tempStartPoint).Length

            Dim angleEnd As Double = (acStartLine.EndPoint.Pos - tempStartPoint).Angle
            Dim distEnd As Double = (acStartLine.EndPoint.Pos - tempStartPoint).Length

            Dim pkTemp As Double = 0
            Dim pkoff As Double = 0
            Dim boolPk As Boolean = userAlign.Plan.CompoundLine.PosToStaOffset(tempStartPoint, pkTemp, pkoff)
            'если стартовый пикет не записан, то записываем его
            If startPositionBridge = 0 Then
                startPositionBridge = pkTemp
                userBridge.startPosition = startPositionBridge
                'перезаписываем данные
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBridge)
                Dim arrayRec As String(,) = Nothing
                ReDim Preserve arrayRec(1, 0)
                arrayRec(0, 0) = "KeyParameters"
                arrayRec(1, 0) = strGSON
                If IsArray(arrayRec) = True Then
                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(axisOldAlign, "PROJECT_BRIDGE", arrayRec, templateXML)
                End If
            End If
            If boolPk = True Then
                startPositionBridge = startPositionBridge + horizontalOffset
                Dim newTempStartPoint As Vector2D = New Vector2D()
                Dim boolFindPoint As Boolean = userAlign.Plan.CompoundLine.StaOffsetToPos(startPositionBridge, 0, newTempStartPoint)
                If boolFindPoint = True Then
                    'создаем произвольный вектор 
                    Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angleStart, distStart + 10)
                    Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angleEnd, distEnd + 10)
                    Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, newTempStartPoint, pt1)
                    If pointIntersectCollection1.Count > 0 Then
                        acStartLine.StartPoint = New Vector3D(pointIntersectCollection1(0), acStartLine.StartPoint.Z)
                    End If
                    Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, newTempStartPoint, pt2)
                    If pointIntersectCollection2.Count > 0 Then
                        acStartLine.EndPoint = New Vector3D(pointIntersectCollection2(0), acStartLine.EndPoint.Z)
                    End If
                End If
            End If
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '==============================================================================================================
        'начинаем перебирать ряды
        For i As Integer = 0 To newArrayPropertiesRows.GetUpperBound(1)
            'номер ряда
            Dim numRowBridge As Integer = newArrayPropertiesRows(0, i)
            'смещение балки в плане
            Dim offsetLenghtAxis As Double = newArrayPropertiesRows(1, i)
            'смещение балки в высоте
            Dim dEarth As Double = newArrayPropertiesRows(2, i)
            'рисуем вспомогательную ось
            Dim axisPline As DwgPolyline = FuncAlignment.FuncOffsetAlignment(activDoc, userAlign, offsetLenghtAxis)
            Dim axisPline3D As IPolyline3D = New Polyline3D()
            axisPline.GetPolyline(axisPline3D)
            'если ось нарисована, продолжаем расчет
            If IsNothing(axisPline) = False Then
                If dictionaryBeams.ContainsKey(numRowBridge) = True Then
                    Dim arrayBeams As String(,) = dictionaryBeams.Item(numRowBridge)
                    If IsArray(arrayBeams) = True Then
                        'задаем 4 точки начального сечения (начальное сечение без наклона)
                        Dim startSectionPoint3d As List(Of Vector3D) = New List(Of Vector3D)
                        Dim endSectionPoint3d As List(Of Vector3D) = New List(Of Vector3D)
                        'начало раскладки балок
                        Dim startAlignPoint As Vector2D = Nothing
                        Dim startAlignDist As Double = -1
                        Dim offsetDist As Double = -1
                        Dim prevNumberBeam As Integer = 0
                        For j As Integer = 0 To arrayBeams.GetUpperBound(1)
                            'читаем название балки
                            Dim keyParamBeam As String = arrayBeams(2, j)
                            Dim idElement As String = arrayBeams(4, j)
                            Dim hgStr As UInteger = CUInt(arrayBeams(3, j))
                            Dim lineDownBalka As DwgLine = Nothing
                            Dim boolFindBalka As Boolean = activDoc.ActiveSpace.Entities.TryGetObject(hgStr, lineDownBalka)
                            If boolFindBalka = False Then Continue For
                            Dim userBeam As Beams = New Beams
                            Try
                                userBeam = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Beams)(keyParamBeam)
                            Catch ex As JsonException
                            End Try
                            If IsNothing(userBeam) = True Then Continue For
                            Dim numProlet As Integer = userBeam.numberProlet
                            Dim numRow As Integer = userBeam.numberRow
                            Dim heightBearm As Double = userBeam.height
                            Dim lenghtPointPrStart As Double = userBeam.a
                            Dim lenghtPointPrEnd As Double = userBeam.b
                            Dim downWidthBearm As Double = userBeam.widthBottom
                            Dim upWidthBearm As Double = userBeam.widthTop
                            Dim lenghtBearm As Double = userBeam.fullLenght
                            Dim zazor As Double = userBeam.clearence
                            Dim dLenghtEarth As Double = userBeam.dEarth + verticalOffset
                            If i < 2 Then
                                'если балка начальная ищем для нее ось опирания (предыдущая балка отсутствует)
                                If startSectionPoint3d.Count = 0 And endSectionPoint3d.Count = 0 Then
                                    Dim arrayLinePillars As DwgLine() = dictionaryAxisPillars.Item(numProlet)
                                    Dim DefiningPillar As DwgLine = arrayLinePillars(1)
                                    'находим точку пересечения трассы смещения и сечения
                                    Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, DefiningPillar.StartPoint.Pos, DefiningPillar.EndPoint.Pos)
                                    Dim elevST As Double = 0
                                    If pointIntersectCollection1.Count > 0 Then
                                        startAlignPoint = pointIntersectCollection1.ElementAt(0)
                                        'низ
                                        startSectionPoint3d.Add(New Vector3D(DefiningPillar.StartPoint.X, DefiningPillar.StartPoint.Y, elevST - heightBearm - dLenghtEarth))
                                        startSectionPoint3d.Add(New Vector3D(DefiningPillar.EndPoint.X, DefiningPillar.EndPoint.Y, elevST - heightBearm - dLenghtEarth))
                                        'верх\=
                                        startSectionPoint3d.Add(New Vector3D(DefiningPillar.StartPoint.X, DefiningPillar.StartPoint.Y, elevST - dLenghtEarth))
                                        startSectionPoint3d.Add(New Vector3D(DefiningPillar.EndPoint.X, DefiningPillar.EndPoint.Y, elevST - dLenghtEarth))
                                    Else
                                        MsgBox("Не удалось определить положение балки. Из возможных причин: недостаточная ширина сооружения!!!")
                                        Continue For
                                    End If
                                    '===========================================================================================================
                                    'находим расстояние на линии
                                    Dim boolStartPoint0 As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startAlignPoint, startAlignDist, offsetDist)
                                    If boolStartPoint0 = False Then
                                        'MsgBox("Ошибка в определении пикетажного значения начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - слишком короткая трасса.")
                                        Continue For
                                    End If
                                    zazor = 0
                                Else
                                    'считаем зазор
                                    Dim leftZazor As Double = zazor
                                    Dim rightZazor As Double = zazor
                                    If dictionaryPillarAxisZazor.Count > 0 Then
                                        If dictionaryPillarAxisZazor.ContainsKey(numProlet) = True Then
                                            Dim arrayZazor As Double() = dictionaryPillarAxisZazor.Item(numProlet)
                                            leftZazor = arrayZazor(0)
                                            rightZazor = arrayZazor(1)
                                        End If
                                    End If
                                    zazor = FuncCalculateZazorBeams(dimBridgeLeft, dimBrigeRight, offsetLenghtAxis, leftZazor, rightZazor)
                                    '======================================================================================================================
                                    Dim boolStartPoint1 As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, startAlignPoint, startAlignDist, offsetDist)
                                    If boolStartPoint1 = True Then
                                        Try
                                            startAlignPoint = PolylineExtentions.StaOffsetToPos(axisPline3D, startAlignDist + zazor, 0)
                                        Catch ex As System.ArgumentOutOfRangeException
                                            'MsgBox("Ошибка в определении пикетажного значения начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - слишком короткая трасса.")
                                            Exit For
                                        End Try
                                    Else
                                        'MsgBox("Ошибка в определении плановых координат начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - слишком короткая трасса.")
                                        Exit For
                                    End If
                                End If
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'находим отметку низа начала балки (с зазором)
                                Dim ElevStartAlignPoint As Double = -9999999
                                Try
                                    ElevStartAlignPoint = userSurface.GetElevation(startAlignPoint)
                                Catch ex As System.Exception
                                    'MsgBox("Ошибка в определении высоты начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности.")
                                    Continue For
                                End Try
                                'отметка низа балки
                                ElevStartAlignPoint = ElevStartAlignPoint - heightBearm - dLenghtEarth
                                'КООРДИНАТЫ Начала раскладки балок
                                Dim startAlignPoint3D As Vector3D = New Vector3D(startAlignPoint.X, startAlignPoint.Y, ElevStartAlignPoint)
                                'находим первую (начальную точку опирания) 
                                Dim newPoint1 As Vector3D = startAlignPoint3D
                                If zazor > 0 Then
                                    newPoint1 = FuncBridge.FuncFindPointBearm(axisPline, startAlignPoint3D, lenghtPointPrStart, userSurface, heightBearm, dLenghtEarth)
                                    If newPoint1.X = -1 And newPoint1.Y = -1 And newPoint1.Z = -1 Then
                                        'MsgBox("Ошибка в определении высоты или пикетажного положения начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности или трассы.")
                                        Exit For
                                    End If
                                End If
                                'находим вторую точку опирания
                                Dim lenghtAxisBeam As Double = Math.Abs(lenghtBearm - lenghtPointPrStart - lenghtPointPrEnd)
                                Dim newPoint2 As Vector3D = FuncBridge.FuncFindPointBearm(axisPline, newPoint1, lenghtAxisBeam, userSurface, heightBearm, dLenghtEarth)
                                If newPoint2.X = -1 And newPoint2.Y = -1 And newPoint2.Z = -1 Then
                                    'MsgBox("Ошибка в определении высоты или пикетажного положения конечной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности или трассы.")
                                    Exit For
                                End If
                                'получаем осевую линию по низу балки за минусом рассояний до точек опирания
                                lineDownBalka.StartPoint = newPoint1
                                lineDownBalka.EndPoint = newPoint2
                                '======================================================================================================================
                                'корректируем длину балки за наклон
                                Dim boolCorrBaem As Boolean = False
                                For k As Integer = 0 To 1000
                                    If j = 0 Then
                                        startSectionPoint3d = New List(Of Vector3D)
                                    End If
                                    Dim boolBalka As Boolean = FuncBridge.FuncCorrBalka(lineDownBalka, startSectionPoint3d, heightBearm, downWidthBearm, upWidthBearm, lenghtPointPrStart, lenghtPointPrEnd, zazor, endSectionPoint3d, axisPline3D, userSurface, dLenghtEarth, userAlign)
                                    If boolBalka = True Then
                                        boolCorrBaem = True
                                        Exit For
                                    End If
                                Next k
                                If boolCorrBaem = False Then
                                    'MsgBox("Ошибка в определении высоты или пикетажного положения конечной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности или трассы.")
                                    Exit For
                                End If
                                '=======================================================================================================================================
                                'вычисляем новое начало раскладки балок
                                Dim startVirtualPoint As Vector3D = Nothing
                                Dim endVirtualPoint As Vector3D = Nothing
                                Dim boolVitrual As Boolean = FuncBridge.FuncVirtualExtendBearm(lineDownBalka, lenghtPointPrStart, lenghtPointPrEnd, startVirtualPoint, endVirtualPoint)
                                If boolVitrual = True Then
                                    Dim boolNewStartPointBeam As Boolean = PolylineExtentions.PosToStaOffset(axisPline3D, endVirtualPoint, startAlignDist, offsetDist)
                                    If boolNewStartPointBeam = True Then
                                        Try
                                            startAlignPoint = PolylineExtentions.StaOffsetToPos(axisPline3D, startAlignDist, offsetDist)
                                        Catch ex As System.ArgumentOutOfRangeException
                                            'MsgBox("Ошибка в определении координат новой начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - слишком короткая трасса.")
                                            Exit For
                                        End Try
                                    Else
                                        'MsgBox("Ошибка в определении пикетажного значения новой начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - слишком короткая трасса.")
                                        Exit For
                                    End If
                                End If
                                'корректируем дополнительные построения данной балки
                                'рисуем вспомогательные линии
                                If endSectionPoint3d.Count > 3 Then
                                    Dim boolFindBeam As Boolean = False
                                    Dim boolFindUpBeamElement1 As Boolean = False
                                    Dim boolFindUpBeamElement2 As Boolean = False
                                    Dim boolFindDownBeamElement1 As Boolean = False
                                    Dim boolFindDownBeamElement2 As Boolean = False
                                    Dim arrayElements As UInteger() = {}
                                    If elementDictionary.Count > 0 Then
                                        If elementDictionary.ContainsKey(idElement) = True Then
                                            arrayElements = elementDictionary.Item(idElement)
                                            For k As Integer = 0 To arrayElements.Length - 1
                                                Dim hgObject As UInteger = arrayElements(k)
                                                Dim acObject As DwgObject = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                Dim arrayData As String(,) = {}
                                                Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(acObject, "PROJECT_BRIDGE", arrayData)
                                                If IsArray(arrayData) Then
                                                    Dim idElem As String = MathFunction.FuncFindValueToArray2d(arrayData, "ElementID", 0, 1)
                                                    If idElem Like idElement Then
                                                        Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                                                        Dim off As Double = Val(MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1))
                                                        If nameObject Like "Низ ребра балки" Then
                                                            If off > 0 Then
                                                                Dim line1 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                                line1.StartPoint = endSectionPoint3d.Item(0)
                                                                line1.EndPoint = startSectionPoint3d.Item(0)
                                                                If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                                                    ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                                                    arrayLineElements(countAtrrayLineElements) = line1
                                                                    countAtrrayLineElements += 1
                                                                End If
                                                                boolFindDownBeamElement1 = True
                                                            Else
                                                                Dim line2 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                                line2.StartPoint = endSectionPoint3d.Item(1)
                                                                line2.EndPoint = startSectionPoint3d.Item(1)
                                                                If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                                                    ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                                                    arrayLineElements(countAtrrayLineElements) = line2
                                                                    countAtrrayLineElements += 1
                                                                End If
                                                                boolFindDownBeamElement2 = True
                                                            End If
                                                        ElseIf nameObject Like "Верх грани плиты балки" Then
                                                            If off > 0 Then
                                                                Dim line3 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                                line3.StartPoint = endSectionPoint3d.Item(2)
                                                                line3.EndPoint = startSectionPoint3d.Item(2)
                                                                If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                                                    ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                                                    arrayLineElements(countAtrrayLineElements) = line3
                                                                    countAtrrayLineElements += 1
                                                                End If
                                                                boolFindUpBeamElement1 = True
                                                            Else
                                                                Dim line4 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                                line4.StartPoint = endSectionPoint3d.Item(3)
                                                                line4.EndPoint = startSectionPoint3d.Item(3)
                                                                If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                                                    ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                                                    arrayLineElements(countAtrrayLineElements) = line4
                                                                    countAtrrayLineElements += 1
                                                                End If
                                                                boolFindUpBeamElement2 = True
                                                            End If
                                                        ElseIf nameObject Like "Балка" Then
                                                            Try
                                                                Dim entity As DwgModel3DElement = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                                entity.Position = lineDownBalka.StartPoint
                                                                Dim element As ConstructedModel3dElement = entity.Element
                                                                'ориентируем балку по оси
                                                                Dim poly As Polyline3D = New Polyline3D()
                                                                poly.Add(New BugleVector3D(lineDownBalka.StartPoint.Pos, 0))
                                                                poly.Add(New BugleVector3D(lineDownBalka.EndPoint.Pos, 0))
                                                                Dim insert_point As Vector3D = poly(0).Vertex
                                                                Dim pivot As Matrix = Matrix.CreateTranslation(insert_point)
                                                                Dim invert As Matrix = Matrix.Invert(pivot)
                                                                Dim Aggregate As ImAggregates = ImAggregates.Create("SmdxManualPolyline")

                                                                Aggregate.ApplayOverridedProperties({
                                                     New ImProperty("mverteces", Nothing,
                                                ImAggregateExtentions.CreateSmdxManualPolyline(poly, invert))
        })
                                                                element.ApplayOverridedProperties({
        New ImProperty("axis_curve", Nothing, Aggregate)
        })
                                                                Dim elementPropertiesObject As ImElement = entity.Element
                                                                Dim generalPropertiesObject1 As ImProperties = elementPropertiesObject.GetProperties()
                                                                If generalPropertiesObject1.Count > 0 Then
                                                                    For k1 As Integer = 0 To generalPropertiesObject1.Count - 1
                                                                        Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k1)
                                                                        If generalPropLevel1.Name Like "Длина балки" Then
                                                                            generalPropLevel1.Value = lenghtBearm
                                                                        ElseIf generalPropLevel1.Name Like "Средняя толщина покрытия над балкой" Then
                                                                            Try
                                                                                Dim arrayElev As Double() = {}
                                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                                If arrayElev.Length > 1 Then
                                                                                    Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                                End If
                                                                            Catch ex As System.Exception
                                                                                generalPropLevel1.Value = dLenghtEarth
                                                                            End Try
                                                                        ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. А" Then
                                                                            Try
                                                                                Dim arrayElev As Double() = {}
                                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                                If arrayElev.Length > 1 Then
                                                                                    Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                                End If
                                                                            Catch ex As System.Exception
                                                                                generalPropLevel1.Value = dLenghtEarth
                                                                            End Try
                                                                        ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. B" Then
                                                                            Try
                                                                                Dim arrayElev As Double() = {}
                                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                                If arrayElev.Length > 1 Then
                                                                                    Dim dEarth1 As Double = arrayElev(1) - arrayElev(3)
                                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                                End If
                                                                            Catch ex As System.Exception
                                                                                generalPropLevel1.Value = dLenghtEarth
                                                                            End Try
                                                                        ElseIf generalPropLevel1.Name Like "Номер пролёта" Then
                                                                            generalPropLevel1.Value = userBeam.numberProlet
                                                                        ElseIf generalPropLevel1.Name Like "Номер балки" Then
                                                                            If userBeam.numberRow < 0 Then
                                                                                generalPropLevel1.Value = "Л" & userBeam.numberRow
                                                                            Else
                                                                                If userBridge.centreAxisBeam = False Then
                                                                                    generalPropLevel1.Value = "П" & userBeam.numberRow
                                                                                Else
                                                                                    generalPropLevel1.Value = userBeam.numberRow
                                                                                End If
                                                                            End If
                                                                        ElseIf generalPropLevel1.Name Like "Высота балки" Then
                                                                            generalPropLevel1.Value = heightBearm
                                                                        ElseIf generalPropLevel1.Name Like "Ширина ребра в зоне опирания" Then
                                                                            generalPropLevel1.Value = downWidthBearm
                                                                        ElseIf generalPropLevel1.Name Like "Ширина плиты" Then
                                                                            generalPropLevel1.Value = upWidthBearm
                                                                        ElseIf generalPropLevel1.Name Like "Участок опирания балки" Then
                                                                            generalPropLevel1.Value = lenghtPointPrStart
                                                                        End If
                                                                    Next k1
                                                                End If
                                                                Dim updateable = TryCast(element, IUpdatable)
                                                                If updateable Is Nothing Then Return
                                                                updateable.BeginUpdate()
                                                                Try
                                                                    element.ApplayOverridedProperties(generalPropertiesObject1)
                                                                Finally
                                                                    updateable.EndUpdate()
                                                                End Try
                                                                boolFindBeam = True
                                                            Catch ex As System.Exception
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                            Next k
                                        End If
                                    End If
                                    If boolFindBeam = False Then
                                        'тлс модель не найдена, вставляем новую
                                        Dim doc As ConstructionDocument = New ConstructionDocument()
                                        Dim nameFileTLC As String = ""
                                        If IsNothing(userBeam.model) = False Then
                                            Dim directoryTLS As String = putchBridge & userBeam.model & "\" & userBeam.modelTLS
                                            If Directory.Exists(directoryTLS) = True Then
                                                Dim arrayTLCFiles As String() = Directory.GetFiles(directoryTLS, "*.tlc")
                                                If IsArray(arrayTLCFiles) = True Then
                                                    If arrayTLCFiles.Length > 0 Then
                                                        nameFileTLC = arrayTLCFiles(0)
                                                        doc.LoadFromFile(nameFileTLC)
                                                    End If
                                                End If
                                            End If
                                        End If
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем многовидовой блок
                                        Try
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'вставляем конструкцию
                                            If File.Exists(nameFileTLC) = True Then
                                                Dim el As ImProperties = New ImProperties()
                                                Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(doc, el)
                                                Dim entity As DwgModel3DElement = New DwgModel3DElement
                                                entity.Position = lineDownBalka.StartPoint
                                                'ориентируем балку по оси
                                                Dim poly As Polyline3D = New Polyline3D()
                                                poly.Add(New BugleVector3D(lineDownBalka.StartPoint, 0))
                                                poly.Add(New BugleVector3D(lineDownBalka.EndPoint, 0))

                                                Dim insert_point As Vector3D = poly(0).Vertex
                                                Dim pivot As Matrix = Matrix.CreateTranslation(insert_point)
                                                Dim invert As Matrix = Matrix.Invert(pivot)
                                                Dim Aggregate As ImAggregates = ImAggregates.Create("SmdxManualPolyline")
                                                Aggregate.ApplayOverridedProperties({
                                         New ImProperty("mverteces", Nothing,
                                    ImAggregateExtentions.CreateSmdxManualPolyline(poly, invert))
                                        })
                                                element.ApplayOverridedProperties({
                                             New ImProperty("axis_curve", Nothing, Aggregate)
                                        })
                                                entity.Element = element
                                                Dim elementPropertiesObject As ImElement = entity.Element
                                                Dim generalPropertiesObject1 As ImProperties = elementPropertiesObject.GetProperties()
                                                If generalPropertiesObject1.Count > 0 Then
                                                    For k As Integer = 0 To generalPropertiesObject1.Count - 1
                                                        Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k)
                                                        If generalPropLevel1.Name Like "Длина балки" Then
                                                            generalPropLevel1.Value = lenghtBearm
                                                        ElseIf generalPropLevel1.Name Like "Средняя толщина покрытия над балкой" Then
                                                            Try
                                                                Dim arrayElev As Double() = {}
                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                If arrayElev.Length > 1 Then
                                                                    Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                End If
                                                            Catch ex As System.Exception
                                                                generalPropLevel1.Value = dLenghtEarth
                                                            End Try
                                                        ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. А" Then
                                                            Try
                                                                Dim arrayElev As Double() = {}
                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                If arrayElev.Length > 1 Then
                                                                    Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                End If
                                                            Catch ex As System.Exception
                                                                generalPropLevel1.Value = dLenghtEarth
                                                            End Try
                                                        ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. B" Then
                                                            Try
                                                                Dim arrayElev As Double() = {}
                                                                Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                If arrayElev.Length > 1 Then
                                                                    Dim dEarth1 As Double = arrayElev(1) - arrayElev(3)
                                                                    generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                End If
                                                            Catch ex As System.Exception
                                                                generalPropLevel1.Value = dLenghtEarth
                                                            End Try
                                                        ElseIf generalPropLevel1.Name Like "Номер пролёта" Then
                                                            generalPropLevel1.Value = userBeam.numberProlet
                                                        ElseIf generalPropLevel1.Name Like "Номер балки" Then
                                                            If userBeam.numberRow < 0 Then
                                                                generalPropLevel1.Value = "Л" & userBeam.numberRow
                                                            Else
                                                                If userBridge.centreAxisBeam = False Then
                                                                    generalPropLevel1.Value = "П" & userBeam.numberRow
                                                                Else
                                                                    generalPropLevel1.Value = userBeam.numberRow
                                                                End If
                                                            End If
                                                        ElseIf generalPropLevel1.Name Like "Высота балки" Then
                                                            generalPropLevel1.Value = heightBearm
                                                        ElseIf generalPropLevel1.Name Like "Ширина ребра в зоне опирания" Then
                                                            generalPropLevel1.Value = downWidthBearm
                                                        ElseIf generalPropLevel1.Name Like "Ширина плиты" Then
                                                            generalPropLevel1.Value = upWidthBearm
                                                        ElseIf generalPropLevel1.Name Like "Участок опирания балки" Then
                                                            generalPropLevel1.Value = lenghtPointPrStart
                                                        End If
                                                    Next
                                                End If
                                                entity.Prepare(activDoc)
                                                entity.Layer = layerBeam
                                                entity.Linetype = nameTypeLineBeam
                                                entity.LinetypeScale = ScaleTypeLineBeam
                                                entity.Lineweight = widthTypeLineBeam
                                                activDoc.ActiveSpace.Add(entity)
                                                Dim arrayRecTLC As String(,) = Nothing
                                                ReDim Preserve arrayRecTLC(1, 5)
                                                arrayRecTLC(0, 0) = "LABEL"
                                                arrayRecTLC(1, 0) = "Мосты и путепроводы"

                                                arrayRecTLC(0, 1) = "Name"
                                                arrayRecTLC(1, 1) = "Балка"

                                                arrayRecTLC(0, 2) = "KeyParameters"
                                                arrayRecTLC(1, 2) = userBeam.model

                                                arrayRecTLC(0, 3) = "ElementID"
                                                arrayRecTLC(1, 3) = idElement

                                                arrayRecTLC(0, 4) = "BrigeID"
                                                arrayRecTLC(1, 4) = idBridge

                                                arrayRecTLC(0, 5) = "NOTE"
                                                arrayRecTLC(1, 5) = "-"

                                                If IsArray(arrayRecTLC) = True Then
                                                    If tablePSBeam.Trim.Length > 0 Then
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(entity, tablePSBeam, arrayRecTLC, templateXML)
                                                    Else
                                                        Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(entity, "PROJECT_BRIDGE", arrayRecTLC)
                                                    End If
                                                End If
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                    End If
                                    If boolFindDownBeamElement1 = False Then
                                        Dim line1 As DwgLine = New DwgLine()
                                        line1.StartPoint = endSectionPoint3d.Item(0)
                                        line1.EndPoint = startSectionPoint3d.Item(0)
                                        If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                            ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                            arrayLineElements(countAtrrayLineElements) = line1
                                            countAtrrayLineElements += 1
                                        End If
                                        line1.Color = colorDownBeam
                                        line1.Layer = layerDownBeam
                                        line1.Linetype = nameTypeLineDownBeam
                                        line1.LinetypeScale = ScaleTypeLineDownBeam
                                        line1.Lineweight = widthTypeLineDownBeam
                                        activDoc.ActiveSpace.Add(line1)
                                        '======================================================================================================================
                                        Dim arrayRec As String(,) = Nothing
                                        ReDim Preserve arrayRec(1, 5)
                                        arrayRec(0, 0) = "LABEL"
                                        arrayRec(1, 0) = "Мосты и путепроводы"

                                        arrayRec(0, 1) = "Name"
                                        arrayRec(1, 1) = "Низ ребра балки"

                                        arrayRec(0, 2) = "KeyParameters"
                                        arrayRec(1, 2) = downWidthBearm / 2

                                        arrayRec(0, 3) = "ElementID"
                                        arrayRec(1, 3) = idElement

                                        arrayRec(0, 4) = "BrigeID"
                                        arrayRec(1, 4) = idBridge

                                        arrayRec(0, 5) = "NOTE"
                                        arrayRec(1, 5) = "-"

                                        If IsArray(arrayRec) = True Then
                                            If tablePSDownBeam.Trim.Length > 0 Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                            Else
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                            End If
                                        End If
                                    End If
                                    If boolFindDownBeamElement2 = False Then
                                        Dim line1 As DwgLine = New DwgLine()
                                        line1.StartPoint = endSectionPoint3d.Item(1)
                                        line1.EndPoint = startSectionPoint3d.Item(1)
                                        If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                            ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                            arrayLineElements(countAtrrayLineElements) = line1
                                            countAtrrayLineElements += 1
                                        End If
                                        line1.Color = colorDownBeam
                                        line1.Layer = layerDownBeam
                                        line1.Linetype = nameTypeLineDownBeam
                                        line1.LinetypeScale = ScaleTypeLineDownBeam
                                        line1.Lineweight = widthTypeLineDownBeam
                                        activDoc.ActiveSpace.Add(line1)
                                        '======================================================================================================================
                                        Dim arrayRec As String(,) = Nothing
                                        ReDim Preserve arrayRec(1, 5)
                                        arrayRec(0, 0) = "LABEL"
                                        arrayRec(1, 0) = "Мосты и путепроводы"

                                        arrayRec(0, 1) = "Name"
                                        arrayRec(1, 1) = "Низ ребра балки"

                                        arrayRec(0, 2) = "KeyParameters"
                                        arrayRec(1, 2) = -1 * downWidthBearm / 2

                                        arrayRec(0, 3) = "ElementID"
                                        arrayRec(1, 3) = idElement

                                        arrayRec(0, 4) = "BrigeID"
                                        arrayRec(1, 4) = idBridge

                                        arrayRec(0, 5) = "NOTE"
                                        arrayRec(1, 5) = "-"

                                        If IsArray(arrayRec) = True Then
                                            If tablePSDownBeam.Trim.Length > 0 Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                            Else
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                            End If
                                        End If
                                    End If
                                    If boolFindUpBeamElement1 = False Then
                                        Dim line1 As DwgLine = New DwgLine()
                                        line1.StartPoint = endSectionPoint3d.Item(2)
                                        line1.EndPoint = startSectionPoint3d.Item(2)
                                        If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                            ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                            arrayLineElements(countAtrrayLineElements) = line1
                                            countAtrrayLineElements += 1
                                        End If
                                        line1.Color = colorUpBeam
                                        line1.Layer = layerUpBeam
                                        line1.Linetype = nameTypeLineUpBeam
                                        line1.LinetypeScale = ScaleTypeLineUpBeam
                                        line1.Lineweight = widthTypeLineUpBeam
                                        activDoc.ActiveSpace.Add(line1)
                                        '======================================================================================================================
                                        Dim arrayRec As String(,) = Nothing
                                        ReDim Preserve arrayRec(1, 5)
                                        arrayRec(0, 0) = "LABEL"
                                        arrayRec(1, 0) = "Мосты и путепроводы"

                                        arrayRec(0, 1) = "Name"
                                        arrayRec(1, 1) = "Верх грани плиты балки"

                                        arrayRec(0, 2) = "KeyParameters"
                                        arrayRec(1, 2) = upWidthBearm / 2

                                        arrayRec(0, 3) = "ElementID"
                                        arrayRec(1, 3) = idElement

                                        arrayRec(0, 4) = "BrigeID"
                                        arrayRec(1, 4) = idBridge

                                        arrayRec(0, 5) = "NOTE"
                                        arrayRec(1, 5) = "-"

                                        If IsArray(arrayRec) = True Then
                                            If tablePSDownBeam.Trim.Length > 0 Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                            Else
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                            End If
                                        End If
                                    End If
                                    If boolFindUpBeamElement2 = False Then
                                        Dim line1 As DwgLine = New DwgLine()
                                        line1.StartPoint = endSectionPoint3d.Item(3)
                                        line1.EndPoint = startSectionPoint3d.Item(3)
                                        If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                            ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                            arrayLineElements(countAtrrayLineElements) = line1
                                            countAtrrayLineElements += 1
                                        End If
                                        line1.Color = colorUpBeam
                                        line1.Layer = layerUpBeam
                                        line1.Linetype = nameTypeLineUpBeam
                                        line1.LinetypeScale = ScaleTypeLineUpBeam
                                        line1.Lineweight = widthTypeLineUpBeam
                                        activDoc.ActiveSpace.Add(line1)
                                        '======================================================================================================================
                                        Dim arrayRec As String(,) = Nothing
                                        ReDim Preserve arrayRec(1, 5)
                                        arrayRec(0, 0) = "LABEL"
                                        arrayRec(1, 0) = "Мосты и путепроводы"

                                        arrayRec(0, 1) = "Name"
                                        arrayRec(1, 1) = "Верх грани плиты балки"

                                        arrayRec(0, 2) = "KeyParameters"
                                        arrayRec(1, 2) = -1 * upWidthBearm / 2

                                        arrayRec(0, 3) = "ElementID"
                                        arrayRec(1, 3) = idElement

                                        arrayRec(0, 4) = "BrigeID"
                                        arrayRec(1, 4) = idBridge

                                        arrayRec(0, 5) = "NOTE"
                                        arrayRec(1, 5) = "-"

                                        If IsArray(arrayRec) = True Then
                                            If tablePSDownBeam.Trim.Length > 0 Then
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                            Else
                                                Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                            End If
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'корректируем оси опирания балок
                                    If i = 0 Then 'первый ряд
                                        Dim arrayLinePillarsStart As DwgLine() = dictionaryAxisPillars.Item(numProlet)
                                        Dim arrayLinePillarsEnd As DwgLine() = dictionaryAxisPillars.Item(numProlet + 1)
                                        If j = 0 Then
                                            'начало моста корректируем ось балки
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                startPillar1.StartPoint = lineDownBalka.StartPoint
                                            End If
                                            'корректируем ось опирания последующей балки
                                            Dim startPillar2 As DwgLine = arrayLinePillarsEnd(0)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.StartPoint = lineDownBalka.EndPoint
                                            End If
                                            'корректируем ось балки
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.StartPoint = endVirtualPoint
                                            End If
                                        ElseIf j = arrayBeams.GetUpperBound(1) Then
                                            'предпоследняя опора
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                Dim startPoint As Vector3D = startPillar1.StartPoint
                                                Dim newStartPoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(startPoint, startVirtualPoint, 3)
                                                startPillar1.StartPoint = newStartPoint
                                            End If
                                            'ось опирания балки
                                            Dim startPillar2 As DwgLine = arrayLinePillarsStart(2)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.StartPoint = lineDownBalka.StartPoint
                                            End If
                                            'конец моста корректируем ось балки
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.StartPoint = lineDownBalka.EndPoint
                                            End If
                                        Else
                                            'ось в начале пролета
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                Dim startPoint As Vector3D = startPillar1.StartPoint
                                                Dim newStartPoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(startPoint, startVirtualPoint, 3)
                                                startPillar1.StartPoint = newStartPoint
                                            End If
                                            'ось опирания балки начало
                                            Dim startPillar2 As DwgLine = arrayLinePillarsStart(2)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.StartPoint = lineDownBalka.StartPoint
                                            End If
                                            'ось опирания балки в конце
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(0)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.StartPoint = lineDownBalka.EndPoint
                                            End If
                                            'ось в конце пролета
                                            Dim startPillar4 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar4) = False Then
                                                startPillar4.StartPoint = endVirtualPoint
                                            End If
                                        End If
                                    ElseIf i = 1 Then 'крайний рад но с другой стороны
                                        Dim arrayLinePillarsStart As DwgLine() = dictionaryAxisPillars.Item(numProlet)
                                        Dim arrayLinePillarsEnd As DwgLine() = dictionaryAxisPillars.Item(numProlet + 1)
                                        If j = 0 Then
                                            'начало моста корректируем ось балки
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                startPillar1.EndPoint = lineDownBalka.StartPoint
                                            End If
                                            'корректируем ось опирания последующей балки
                                            Dim startPillar2 As DwgLine = arrayLinePillarsEnd(0)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.EndPoint = lineDownBalka.EndPoint
                                            End If
                                            'корректируем ось балки
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.EndPoint = endVirtualPoint
                                            End If
                                        ElseIf j = arrayBeams.GetUpperBound(1) Then
                                            'предпоследняя опора
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                Dim startPoint As Vector3D = startPillar1.EndPoint
                                                Dim newStartPoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(startPoint, startVirtualPoint, 3)
                                                startPillar1.EndPoint = newStartPoint
                                            End If
                                            'ось опирания балки
                                            Dim startPillar2 As DwgLine = arrayLinePillarsStart(2)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.EndPoint = lineDownBalka.StartPoint
                                            End If
                                            'конец моста корректируем ось балки
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.EndPoint = lineDownBalka.EndPoint
                                            End If
                                        Else
                                            'ось в начале пролета
                                            Dim startPillar1 As DwgLine = arrayLinePillarsStart(1)
                                            If IsNothing(startPillar1) = False Then
                                                Dim startPoint As Vector3D = startPillar1.EndPoint
                                                Dim newStartPoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(startPoint, startVirtualPoint, 3)
                                                startPillar1.EndPoint = newStartPoint
                                            End If
                                            'ось опирания балки начало
                                            Dim startPillar2 As DwgLine = arrayLinePillarsStart(2)
                                            If IsNothing(startPillar2) = False Then
                                                startPillar2.EndPoint = lineDownBalka.StartPoint
                                            End If
                                            'ось опирания балки в конце
                                            Dim startPillar3 As DwgLine = arrayLinePillarsEnd(0)
                                            If IsNothing(startPillar3) = False Then
                                                startPillar3.EndPoint = lineDownBalka.EndPoint
                                            End If
                                            'ось в конце пролета
                                            Dim startPillar4 As DwgLine = arrayLinePillarsEnd(1)
                                            If IsNothing(startPillar4) = False Then
                                                startPillar4.EndPoint = endVirtualPoint
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                Dim arrayLinePillars1 As DwgLine() = dictionaryAxisPillars.Item(numProlet)
                                Dim arrayLinePillars2 As DwgLine() = dictionaryAxisPillars.Item(numProlet + 1)

                                Dim StartAxisPillar As DwgLine = arrayLinePillars1(2)
                                Dim EndAxisPillar As DwgLine = arrayLinePillars2(0)
                                If j = 0 Then
                                    StartAxisPillar = arrayLinePillars1(1)
                                ElseIf j = arrayBeams.GetUpperBound(1) Then
                                    EndAxisPillar = arrayLinePillars2(1)
                                End If
                                Dim startAxisBeamPoint As Vector3D = New Vector3D()
                                Dim endAxisBeamPoint As Vector3D = New Vector3D()

                                'находим точку пересечения трассы смещения и сечения
                                Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, StartAxisPillar.StartPoint.Pos, StartAxisPillar.EndPoint.Pos)
                                If pointIntersectCollection1.Count > 0 Then
                                    startAlignPoint = pointIntersectCollection1.ElementAt(0)
                                Else
                                    Continue For
                                End If
                                'находим отметку низа начала балки (с зазором)
                                Dim ElevStartAxisBeam As Double = -9999999
                                Try
                                    ElevStartAxisBeam = userSurface.GetElevation(startAlignPoint) - heightBearm - dLenghtEarth
                                Catch ex As System.Exception
                                    'MsgBox("Ошибка в определении высоты начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности.")
                                    Continue For
                                End Try
                                startAxisBeamPoint = New Vector3D(startAlignPoint, ElevStartAxisBeam)

                                'находим точку пересечения трассы смещения и последнего сечения
                                Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3D, EndAxisPillar.StartPoint.Pos, EndAxisPillar.EndPoint.Pos)
                                If pointIntersectCollection2.Count > 0 Then
                                    Dim endAlignPoint As Vector2D = pointIntersectCollection2.ElementAt(0)
                                    'находим отметку низа начала балки (с зазором)
                                    Dim ElevEndAxisBeam As Double = -9999999
                                    Try
                                        ElevEndAxisBeam = userSurface.GetElevation(endAlignPoint) - heightBearm - dLenghtEarth
                                    Catch ex As System.Exception
                                        'MsgBox("Ошибка в определении высоты начальной точки опирания балки № " & numbBearmStr & ", ряд № " & Val(arrayPr2(5, i)) & " построен не полностью!!! Возможная причина - расчетная балка за пределами проектной поверхности.")
                                        Continue For
                                    End Try
                                    endAxisBeamPoint = New Vector3D(endAlignPoint, ElevEndAxisBeam)
                                Else
                                    Continue For
                                End If
                                'заново формируем балку
                                lineDownBalka.StartPoint = startAxisBeamPoint
                                lineDownBalka.EndPoint = endAxisBeamPoint
                                'восстанавливаем вспомогательные построения
                                Dim startCollectPoint As List(Of Vector3D) = New List(Of Vector3D)
                                Dim endCollectPoint As List(Of Vector3D) = New List(Of Vector3D)
                                Dim restoreBeam As Boolean = FuncBridge.FuncRestoreCoordinatesBalka(lineDownBalka, userBeam.height, userBeam.widthBottom, userBeam.widthTop, userBeam.a, userBeam.b, startCollectPoint, endCollectPoint)
                                'корректируем балку по высоте
                                Dim elev1 As Double = userSurface.GetElevation(startCollectPoint.Item(2).Pos)
                                Dim elevTopStart1 As Double = elev1 - (dLenghtEarth + startCollectPoint.Item(2).Z)

                                Dim elev2 As Double = userSurface.GetElevation(startCollectPoint.Item(3).Pos)
                                Dim elevTopStart2 As Double = elev2 - (dLenghtEarth + startCollectPoint.Item(3).Z)

                                Dim elev3 As Double = userSurface.GetElevation(endCollectPoint.Item(2).Pos)
                                Dim elevTopStart3 As Double = elev3 - (dLenghtEarth + endCollectPoint.Item(2).Z)

                                Dim elev4 As Double = userSurface.GetElevation(endCollectPoint.Item(3).Pos)
                                Dim elevTopStart4 As Double = elev4 - (endCollectPoint.Item(3).Z + dLenghtEarth)


                                If elevTopStart1 < 0 And elevTopStart2 < 0 Then
                                    If Math.Abs(elevTopStart1) > Math.Abs(elevTopStart2) Then
                                        startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart1)
                                    Else
                                        startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart2)
                                    End If
                                ElseIf elevTopStart1 > 0 And elevTopStart2 > 0 Then
                                    If elevTopStart1 > elevTopStart2 Then
                                        startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart2)
                                    Else
                                        startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart1)
                                    End If
                                ElseIf elevTopStart1 < 0 And elevTopStart2 > 0 Then
                                    startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart1)
                                ElseIf elevTopStart1 > 0 And elevTopStart2 < 0 Then
                                    startAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(startAxisBeamPoint.X, startAxisBeamPoint.Y, startAxisBeamPoint.Z + elevTopStart2)
                                End If
                                lineDownBalka.StartPoint = startAxisBeamPoint

                                If elevTopStart3 < 0 And elevTopStart4 < 0 Then
                                    If Math.Abs(elevTopStart3) > Math.Abs(elevTopStart4) Then
                                        endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart3)
                                    Else
                                        endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart4)
                                    End If
                                ElseIf elevTopStart3 > 0 And elevTopStart4 > 0 Then
                                    If elevTopStart3 > elevTopStart4 Then
                                        endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart4)
                                    Else
                                        endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart3)
                                    End If
                                ElseIf elevTopStart3 < 0 And elevTopStart4 > 0 Then
                                    endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart3)
                                ElseIf elevTopStart3 > 0 And elevTopStart4 < 0 Then
                                    endAxisBeamPoint = New Topomatic.Cad.Foundation.Vector3D(endAxisBeamPoint.X, endAxisBeamPoint.Y, endAxisBeamPoint.Z + elevTopStart4)
                                End If
                                lineDownBalka.EndPoint = endAxisBeamPoint
                                startCollectPoint = New List(Of Topomatic.Cad.Foundation.Vector3D)
                                endCollectPoint = New List(Of Topomatic.Cad.Foundation.Vector3D)
                                restoreBeam = FuncBridge.FuncRestoreCoordinatesBalka(lineDownBalka, userBeam.height, userBeam.widthBottom, userBeam.widthTop, userBeam.a, userBeam.b, startCollectPoint, endCollectPoint)
                                Dim boolFindBeam As Boolean = False
                                Dim boolFindUpBeamElement1 As Boolean = False
                                Dim boolFindUpBeamElement2 As Boolean = False
                                Dim boolFindDownBeamElement1 As Boolean = False
                                Dim boolFindDownBeamElement2 As Boolean = False
                                Dim arrayElements As UInteger() = {}
                                If elementDictionary.Count > 0 Then
                                    If elementDictionary.ContainsKey(idElement) = True Then
                                        arrayElements = elementDictionary.Item(idElement)
                                        For k As Integer = 0 To arrayElements.Length - 1
                                            Dim hgObject As UInteger = arrayElements(k)
                                            Dim acObject As DwgObject = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                            Dim arrayData As String(,) = {}
                                            Dim boolFindData As Boolean = FuncXRecords.FuncReadXData(acObject, "PROJECT_BRIDGE", arrayData)
                                            If IsArray(arrayData) Then
                                                Dim idElem As String = MathFunction.FuncFindValueToArray2d(arrayData, "ElementID", 0, 1)
                                                If idElem Like idElement Then
                                                    Dim nameObject As String = MathFunction.FuncFindValueToArray2d(arrayData, "Name", 0, 1)
                                                    Dim off As Double = Val(MathFunction.FuncFindValueToArray2d(arrayData, "KeyParameters", 0, 1))
                                                    If nameObject Like "Низ ребра балки" Then
                                                        If off > 0 Then
                                                            Dim line1 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                            line1.StartPoint = startCollectPoint.Item(0)
                                                            line1.EndPoint = endCollectPoint.Item(0)
                                                            boolFindDownBeamElement1 = True
                                                        Else
                                                            Dim line2 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                            line2.StartPoint = startCollectPoint.Item(1)
                                                            line2.EndPoint = endCollectPoint.Item(1)
                                                            boolFindDownBeamElement2 = True
                                                        End If
                                                    ElseIf nameObject Like "Верх грани плиты балки" Then
                                                        If off > 0 Then
                                                            Dim line3 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                            line3.StartPoint = startCollectPoint.Item(2)
                                                            line3.EndPoint = endCollectPoint.Item(2)
                                                            boolFindUpBeamElement1 = True
                                                        Else
                                                            Dim line4 As DwgLine = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                            line4.StartPoint = startCollectPoint.Item(3)
                                                            line4.EndPoint = endCollectPoint.Item(3)
                                                            boolFindUpBeamElement2 = True
                                                        End If
                                                    ElseIf nameObject Like "Балка" Then
                                                        Try
                                                            Dim entity As DwgModel3DElement = activDoc.ActiveSpace.Entities.GetObject(hgObject)
                                                            entity.Position = lineDownBalka.StartPoint
                                                            Dim element As ConstructedModel3dElement = entity.Element
                                                            'ориентируем балку по оси
                                                            Dim poly As Polyline3D = New Polyline3D()
                                                            poly.Add(New BugleVector3D(lineDownBalka.StartPoint.Pos, 0))
                                                            poly.Add(New BugleVector3D(lineDownBalka.EndPoint.Pos, 0))
                                                            Dim insert_point As Vector3D = poly(0).Vertex
                                                            Dim pivot As Matrix = Matrix.CreateTranslation(insert_point)
                                                            Dim invert As Matrix = Matrix.Invert(pivot)
                                                            Dim Aggregate As ImAggregates = ImAggregates.Create("SmdxManualPolyline")

                                                            Aggregate.ApplayOverridedProperties({
                                                 New ImProperty("mverteces", Nothing,
                                            ImAggregateExtentions.CreateSmdxManualPolyline(poly, invert))
    })
                                                            element.ApplayOverridedProperties({
    New ImProperty("axis_curve", Nothing, Aggregate)
    })


                                                            Dim elementPropertiesObject As ImElement = entity.Element
                                                            Dim generalPropertiesObject1 As ImProperties = elementPropertiesObject.GetProperties()
                                                            If generalPropertiesObject1.Count > 0 Then
                                                                For k1 As Integer = 0 To generalPropertiesObject1.Count - 1
                                                                    Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k1)
                                                                    If generalPropLevel1.Name Like "Длина балки" Then
                                                                        generalPropLevel1.Value = lenghtBearm
                                                                    ElseIf generalPropLevel1.Name Like "Средняя толщина покрытия над балкой" Then
                                                                        Try
                                                                            Dim arrayElev As Double() = {}
                                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                            If arrayElev.Length > 1 Then
                                                                                Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                                generalPropLevel1.Value = dEarth1
                                                                            End If
                                                                        Catch ex As System.Exception
                                                                            generalPropLevel1.Value = dLenghtEarth
                                                                        End Try
                                                                    ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. А" Then
                                                                        Try
                                                                            Dim arrayElev As Double() = {}
                                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                            If arrayElev.Length > 1 Then
                                                                                Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                                generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                            End If
                                                                        Catch ex As System.Exception
                                                                            generalPropLevel1.Value = dLenghtEarth
                                                                        End Try
                                                                    ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. B" Then
                                                                        Try
                                                                            Dim arrayElev As Double() = {}
                                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                                            If arrayElev.Length > 1 Then
                                                                                Dim dEarth1 As Double = arrayElev(1) - arrayElev(3)
                                                                                generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                                            End If
                                                                        Catch ex As System.Exception
                                                                            generalPropLevel1.Value = dLenghtEarth
                                                                        End Try
                                                                    ElseIf generalPropLevel1.Name Like "Номер пролёта" Then
                                                                        generalPropLevel1.Value = userBeam.numberProlet
                                                                    ElseIf generalPropLevel1.Name Like "Номер балки" Then
                                                                        If userBeam.numberRow < 0 Then
                                                                            generalPropLevel1.Value = "Л" & userBeam.numberRow
                                                                        Else
                                                                            If userBridge.centreAxisBeam = False Then
                                                                                generalPropLevel1.Value = "П" & userBeam.numberRow
                                                                            Else
                                                                                generalPropLevel1.Value = userBeam.numberRow
                                                                            End If
                                                                        End If
                                                                    ElseIf generalPropLevel1.Name Like "Высота балки" Then
                                                                        generalPropLevel1.Value = heightBearm
                                                                    ElseIf generalPropLevel1.Name Like "Ширина ребра в зоне опирания" Then
                                                                        generalPropLevel1.Value = downWidthBearm
                                                                    ElseIf generalPropLevel1.Name Like "Ширина плиты" Then
                                                                        generalPropLevel1.Value = upWidthBearm
                                                                    ElseIf generalPropLevel1.Name Like "Участок опирания балки" Then
                                                                        generalPropLevel1.Value = lenghtPointPrStart
                                                                    End If
                                                                Next k1
                                                            End If
                                                            Dim updateable = TryCast(element, IUpdatable)
                                                            If updateable Is Nothing Then Return
                                                            updateable.BeginUpdate()
                                                            Try
                                                                element.ApplayOverridedProperties(generalPropertiesObject1)
                                                            Finally
                                                                updateable.EndUpdate()
                                                            End Try
                                                            boolFindBeam = True
                                                        Catch ex As System.Exception
                                                        End Try
                                                    End If
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                                If boolFindBeam = False Then
                                    'тлс модель не найдена, вставляем новую
                                    Dim doc As ConstructionDocument = New ConstructionDocument()
                                    Dim nameFileTLC As String = ""
                                    If IsNothing(userBeam.model) = False Then
                                        Dim directoryTLS As String = putchBridge & userBeam.nameAlbum & "\" & userBeam.modelTLS
                                        If Directory.Exists(directoryTLS) = True Then
                                            Dim arrayTLCFiles As String() = Directory.GetFiles(directoryTLS, "*.tlc")
                                            If IsArray(arrayTLCFiles) = True Then
                                                If arrayTLCFiles.Length > 0 Then
                                                    nameFileTLC = arrayTLCFiles(0)
                                                    doc.LoadFromFile(nameFileTLC)
                                                End If
                                            End If
                                        End If
                                    End If
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'вставляем многовидовой блок
                                    Try
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'вставляем конструкцию
                                        If File.Exists(nameFileTLC) = True Then
                                            Dim el As ImProperties = New ImProperties()
                                            Dim element As ConstructedModel3dElement = New ConstructedModel3dElement(doc, el)
                                            Dim entity As DwgModel3DElement = New DwgModel3DElement
                                            entity.Position = lineDownBalka.StartPoint
                                            'ориентируем балку по оси
                                            Dim poly As Polyline3D = New Polyline3D()
                                            poly.Add(New BugleVector3D(lineDownBalka.StartPoint, 0))
                                            poly.Add(New BugleVector3D(lineDownBalka.EndPoint, 0))

                                            Dim insert_point As Vector3D = poly(0).Vertex
                                            Dim pivot As Matrix = Matrix.CreateTranslation(insert_point)
                                            Dim invert As Matrix = Matrix.Invert(pivot)
                                            Dim Aggregate As ImAggregates = ImAggregates.Create("SmdxManualPolyline")
                                            Aggregate.ApplayOverridedProperties({
                                     New ImProperty("mverteces", Nothing,
                                ImAggregateExtentions.CreateSmdxManualPolyline(poly, invert))
                                    })
                                            element.ApplayOverridedProperties({
                                         New ImProperty("axis_curve", Nothing, Aggregate)
                                    })
                                            entity.Element = element
                                            Dim elementPropertiesObject As ImElement = entity.Element
                                            Dim generalPropertiesObject1 As ImProperties = elementPropertiesObject.GetProperties()
                                            If generalPropertiesObject1.Count > 0 Then
                                                For k As Integer = 0 To generalPropertiesObject1.Count - 1
                                                    Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k)
                                                    If generalPropLevel1.Name Like "Длина балки" Then
                                                        generalPropLevel1.Value = lenghtBearm
                                                    ElseIf generalPropLevel1.Name Like "Средняя толщина покрытия над балкой" Then
                                                        Try
                                                            Dim arrayElev As Double() = {}
                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                            If arrayElev.Length > 1 Then
                                                                dEarth = arrayElev(0) - arrayElev(2)
                                                                generalPropLevel1.Value = dEarth
                                                            End If
                                                        Catch ex As System.Exception
                                                            generalPropLevel1.Value = Math.Round(dEarth, 3)
                                                        End Try
                                                    ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. А" Then
                                                        Try
                                                            Dim arrayElev As Double() = {}
                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                            If arrayElev.Length > 1 Then
                                                                Dim dEarth1 As Double = arrayElev(0) - arrayElev(2)
                                                                generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                            End If
                                                        Catch ex As System.Exception
                                                            generalPropLevel1.Value = dLenghtEarth
                                                        End Try
                                                    ElseIf generalPropLevel1.Name Like "Толщина покрытия над балкой в т. B" Then
                                                        Try
                                                            Dim arrayElev As Double() = {}
                                                            Dim boolFindElev As Double = FuncBridge.FuncReturnTopElevationPointPrBeam(lineDownBalka, userBeam.height, userSurface, arrayElev)
                                                            If arrayElev.Length > 1 Then
                                                                Dim dEarth1 As Double = arrayElev(1) - arrayElev(3)
                                                                generalPropLevel1.Value = Math.Round(dEarth1, 3)
                                                            End If
                                                        Catch ex As System.Exception
                                                            generalPropLevel1.Value = dLenghtEarth
                                                        End Try
                                                    ElseIf generalPropLevel1.Name Like "Номер пролёта" Then
                                                        generalPropLevel1.Value = userBeam.numberProlet
                                                    ElseIf generalPropLevel1.Name Like "Номер балки" Then
                                                        If userBeam.numberRow < 0 Then
                                                            generalPropLevel1.Value = "Л" & userBeam.numberRow
                                                        Else
                                                            If userBridge.centreAxisBeam = False Then
                                                                generalPropLevel1.Value = "П" & userBeam.numberRow
                                                            Else
                                                                generalPropLevel1.Value = userBeam.numberRow
                                                            End If
                                                        End If
                                                    ElseIf generalPropLevel1.Name Like "Высота балки" Then
                                                        generalPropLevel1.Value = heightBearm
                                                    ElseIf generalPropLevel1.Name Like "Ширина ребра в зоне опирания" Then
                                                        generalPropLevel1.Value = downWidthBearm
                                                    ElseIf generalPropLevel1.Name Like "Ширина плиты" Then
                                                        generalPropLevel1.Value = upWidthBearm
                                                    ElseIf generalPropLevel1.Name Like "Участок опирания балки" Then
                                                        generalPropLevel1.Value = lenghtPointPrStart
                                                    End If
                                                Next
                                            End If
                                            entity.Prepare(activDoc)
                                            entity.Layer = layerBeam
                                            entity.Linetype = nameTypeLineBeam
                                            entity.LinetypeScale = ScaleTypeLineBeam
                                            entity.Lineweight = widthTypeLineBeam
                                            activDoc.ActiveSpace.Add(entity)
                                            Dim arrayRecTLC As String(,) = Nothing
                                            ReDim Preserve arrayRecTLC(1, 5)
                                            arrayRecTLC(0, 0) = "LABEL"
                                            arrayRecTLC(1, 0) = "Мосты и путепроводы"

                                            arrayRecTLC(0, 1) = "Name"
                                            arrayRecTLC(1, 1) = "Балка"

                                            arrayRecTLC(0, 2) = "KeyParameters"
                                            arrayRecTLC(1, 2) = userBeam.model

                                            arrayRecTLC(0, 3) = "ElementID"
                                            arrayRecTLC(1, 3) = idElement

                                            arrayRecTLC(0, 4) = "BrigeID"
                                            arrayRecTLC(1, 4) = idBridge

                                            arrayRecTLC(0, 5) = "NOTE"
                                            arrayRecTLC(1, 5) = "-"

                                            If IsArray(arrayRecTLC) = True Then
                                                If tablePSBeam.Trim.Length > 0 Then
                                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(entity, tablePSBeam, arrayRecTLC, templateXML)
                                                Else
                                                    Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(entity, "PROJECT_BRIDGE", arrayRecTLC)
                                                End If
                                            End If
                                        End If
                                    Catch ex As System.Exception
                                    End Try
                                End If
                                If boolFindDownBeamElement1 = False Then
                                    Dim line1 As DwgLine = New DwgLine()
                                    line1.StartPoint = endCollectPoint.Item(0)
                                    line1.EndPoint = startCollectPoint.Item(0)
                                    If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                        ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                        arrayLineElements(countAtrrayLineElements) = line1
                                        countAtrrayLineElements += 1
                                    End If
                                    line1.Color = colorDownBeam
                                    line1.Layer = layerDownBeam
                                    line1.Linetype = nameTypeLineDownBeam
                                    line1.LinetypeScale = ScaleTypeLineDownBeam
                                    line1.Lineweight = widthTypeLineDownBeam
                                    activDoc.ActiveSpace.Add(line1)
                                    '======================================================================================================================
                                    Dim arrayRec As String(,) = Nothing
                                    ReDim Preserve arrayRec(1, 5)
                                    arrayRec(0, 0) = "LABEL"
                                    arrayRec(1, 0) = "Мосты и путепроводы"

                                    arrayRec(0, 1) = "Name"
                                    arrayRec(1, 1) = "Низ ребра балки"

                                    arrayRec(0, 2) = "KeyParameters"
                                    arrayRec(1, 2) = downWidthBearm / 2

                                    arrayRec(0, 3) = "ElementID"
                                    arrayRec(1, 3) = idElement

                                    arrayRec(0, 4) = "BrigeID"
                                    arrayRec(1, 4) = idBridge

                                    arrayRec(0, 5) = "NOTE"
                                    arrayRec(1, 5) = "-"

                                    If IsArray(arrayRec) = True Then
                                        If tablePSDownBeam.Trim.Length > 0 Then
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                        Else
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                        End If
                                    End If
                                End If
                                If boolFindDownBeamElement2 = False Then
                                    Dim line1 As DwgLine = New DwgLine()
                                    line1.StartPoint = endCollectPoint.Item(1)
                                    line1.EndPoint = startCollectPoint.Item(1)
                                    If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                        ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                        arrayLineElements(countAtrrayLineElements) = line1
                                        countAtrrayLineElements += 1
                                    End If
                                    line1.Color = colorDownBeam
                                    line1.Layer = layerDownBeam
                                    line1.Linetype = nameTypeLineDownBeam
                                    line1.LinetypeScale = ScaleTypeLineDownBeam
                                    line1.Lineweight = widthTypeLineDownBeam
                                    activDoc.ActiveSpace.Add(line1)
                                    '======================================================================================================================
                                    Dim arrayRec As String(,) = Nothing
                                    ReDim Preserve arrayRec(1, 5)
                                    arrayRec(0, 0) = "LABEL"
                                    arrayRec(1, 0) = "Мосты и путепроводы"

                                    arrayRec(0, 1) = "Name"
                                    arrayRec(1, 1) = "Низ ребра балки"

                                    arrayRec(0, 2) = "KeyParameters"
                                    arrayRec(1, 2) = -1 * downWidthBearm / 2

                                    arrayRec(0, 3) = "ElementID"
                                    arrayRec(1, 3) = idElement

                                    arrayRec(0, 4) = "BrigeID"
                                    arrayRec(1, 4) = idBridge

                                    arrayRec(0, 5) = "NOTE"
                                    arrayRec(1, 5) = "-"

                                    If IsArray(arrayRec) = True Then
                                        If tablePSDownBeam.Trim.Length > 0 Then
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                        Else
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                        End If
                                    End If
                                End If
                                If boolFindUpBeamElement1 = False Then
                                    Dim line1 As DwgLine = New DwgLine()
                                    line1.StartPoint = endCollectPoint.Item(2)
                                    line1.EndPoint = startCollectPoint.Item(2)
                                    If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                        ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                        arrayLineElements(countAtrrayLineElements) = line1
                                        countAtrrayLineElements += 1
                                    End If
                                    line1.Color = colorUpBeam
                                    line1.Layer = layerUpBeam
                                    line1.Linetype = nameTypeLineUpBeam
                                    line1.LinetypeScale = ScaleTypeLineUpBeam
                                    line1.Lineweight = widthTypeLineUpBeam
                                    activDoc.ActiveSpace.Add(line1)
                                    '======================================================================================================================
                                    Dim arrayRec As String(,) = Nothing
                                    ReDim Preserve arrayRec(1, 5)
                                    arrayRec(0, 0) = "LABEL"
                                    arrayRec(1, 0) = "Мосты и путепроводы"

                                    arrayRec(0, 1) = "Name"
                                    arrayRec(1, 1) = "Верх грани плиты балки"

                                    arrayRec(0, 2) = "KeyParameters"
                                    arrayRec(1, 2) = upWidthBearm / 2

                                    arrayRec(0, 3) = "ElementID"
                                    arrayRec(1, 3) = idElement

                                    arrayRec(0, 4) = "BrigeID"
                                    arrayRec(1, 4) = idBridge

                                    arrayRec(0, 5) = "NOTE"
                                    arrayRec(1, 5) = "-"

                                    If IsArray(arrayRec) = True Then
                                        If tablePSDownBeam.Trim.Length > 0 Then
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                        Else
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                        End If
                                    End If
                                End If
                                If boolFindUpBeamElement2 = False Then
                                    Dim line1 As DwgLine = New DwgLine()
                                    line1.StartPoint = endCollectPoint.Item(3)
                                    line1.EndPoint = startCollectPoint.Item(3)
                                    If j = 0 Or j = arrayBeams.GetUpperBound(1) Then
                                        ReDim Preserve arrayLineElements(countAtrrayLineElements)
                                        arrayLineElements(countAtrrayLineElements) = line1
                                        countAtrrayLineElements += 1
                                    End If
                                    line1.Color = colorUpBeam
                                    line1.Layer = layerUpBeam
                                    line1.Linetype = nameTypeLineUpBeam
                                    line1.LinetypeScale = ScaleTypeLineUpBeam
                                    line1.Lineweight = widthTypeLineUpBeam
                                    activDoc.ActiveSpace.Add(line1)
                                    '======================================================================================================================
                                    Dim arrayRec As String(,) = Nothing
                                    ReDim Preserve arrayRec(1, 5)
                                    arrayRec(0, 0) = "LABEL"
                                    arrayRec(1, 0) = "Мосты и путепроводы"

                                    arrayRec(0, 1) = "Name"
                                    arrayRec(1, 1) = "Верх грани плиты балки"

                                    arrayRec(0, 2) = "KeyParameters"
                                    arrayRec(1, 2) = -1 * upWidthBearm / 2

                                    arrayRec(0, 3) = "ElementID"
                                    arrayRec(1, 3) = idElement

                                    arrayRec(0, 4) = "BrigeID"
                                    arrayRec(1, 4) = idBridge

                                    arrayRec(0, 5) = "NOTE"
                                    arrayRec(1, 5) = "-"

                                    If IsArray(arrayRec) = True Then
                                        If tablePSDownBeam.Trim.Length > 0 Then
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXData(line1, tablePSDownBeam, arrayRec, templateXML)
                                        Else
                                            Dim boolInsDataPS As Boolean = FuncXRecords.FuncCreateXDataSystem(line1, "PROJECT_BRIDGE", arrayRec)
                                        End If
                                    End If
                                End If
                            End If
                        Next j
                    End If
                End If
            End If
            axisPline.Clear()
        Next i
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'оформление
        'оси
        For i As Integer = 0 To dictionaryAxisPillars.Count - 1
            Dim keyVal As KeyValuePair(Of Integer, DwgLine()) = dictionaryAxisPillars.ElementAt(1)
            Dim arrayLineAxis As DwgLine() = keyVal.Value
            If IsArray(arrayLineAxis) = True Then
                For j As Integer = 0 To arrayLineAxis.Length - 1
                    If IsNothing(arrayLineAxis(j)) = False Then
                        Dim line1 As DwgLine = arrayLineAxis(j)
                        'временно удлинняем линию до пересечения с габаритами
                        FuncBridge.FuncExtendBearm(line1, 20, 20)
                        'ищем пересечение с лева
                        Dim tempElevLeft As Double = line1.StartPoint.Z
                        Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, line1.StartPoint.Pos, line1.EndPoint.Pos)
                        If pointIntersectCollection1.Count > 0 Then
                            Dim stLine As Vector3D = New Vector3D(pointIntersectCollection1.ElementAt(0), tempElevLeft)
                            line1.StartPoint = stLine
                        End If
                        'ищем пересечение с справа
                        Dim tempElevRight As Double = line1.EndPoint.Z
                        Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, line1.StartPoint.Pos, line1.EndPoint.Pos)
                        If pointIntersectCollection2.Count > 0 Then
                            Dim stLine As Vector3D = New Vector3D(pointIntersectCollection2.ElementAt(0), tempElevRight)
                            line1.EndPoint = stLine
                        End If
                        If j = 1 Then
                            FuncBridge.FuncExtendBearm(line1, 2, 2)
                        End If
                    End If
                Next j
            End If
        Next i
        'ось центральная
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'рисование графики
        Dim offsetLeftPointStart As Vector2D = Nothing
        Dim offsetRightPointStart As Vector2D = Nothing
        Dim offsetLeftPointEnd As Vector2D = Nothing
        Dim offsetRightPointEnd As Vector2D = Nothing

        Dim tempStartLine As DwgLine = acStartLine.Clone()
        Dim tempEndLine As DwgLine = acEndLine.Clone()

        Dim tempStartPointLeft As Vector3D = Nothing
        Dim tempStartPointRight As Vector3D = Nothing
        Dim boolV1 As Boolean = FuncBridge.FuncVirtualExtendBearm(acStartLine, 100, 100, tempStartPointLeft, tempStartPointRight)
        Dim tempEndPointLeft As Vector3D = Nothing
        Dim tempEndPointRight As Vector3D = Nothing
        Dim boolV2 As Boolean = FuncBridge.FuncVirtualExtendBearm(acEndLine, 100, 100, tempEndPointLeft, tempEndPointRight)
        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'делаем смещение крайних осей для поиска габарита
        Dim offsetLeftStart As Double = 0
        Dim offsetRightStart As Double = 0
        'ищем габарит в начале моста
        For Each acLine As DwgLine In arrayLineElements
            Dim pointIntersect As Vector2D = New Vector2D()
            Dim boolIntersect As Boolean = MathFunction.FuncIntersectionTwoSegments(acLine.StartPoint.Pos, acLine.EndPoint.Pos, tempStartPointLeft.Pos, tempStartPointRight.Pos, pointIntersect)
            If boolIntersect = True Then
                Dim off0 As Double = 0
                Dim off1 As Double = 0
                Dim off2 As Double = 0
                Dim pk0 As Double = 0
                Dim pk1 As Double = 0
                Dim pk2 As Double = 0
                Dim boolPK0 As Boolean = axisPline0.PosToStaOffset(pointIntersect, pk0, off0)
                Dim boolPK1 As Boolean = axisPline0.PosToStaOffset(acLine.StartPoint.Pos, pk1, off1)
                Dim boolPK2 As Boolean = axisPline0.PosToStaOffset(acLine.EndPoint.Pos, pk2, off2)
                If boolPK0 = True And boolPK1 = True And boolPK2 = True Then
                    If pk1 < pk0 Then
                        Dim tempdist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.StartPoint.Pos)
                        Dim tempdist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.EndPoint.Pos)
                        If tempdist > tempdist2 Then tempdist = tempdist2
                        If off1 < 0 Then
                            If tempdist > offsetLeftStart Then
                                offsetLeftStart = tempdist
                                offsetLeftPointStart = acLine.StartPoint.Pos
                            End If
                        Else
                            If tempdist > offsetRightStart Then
                                offsetRightStart = tempdist
                                offsetRightPointStart = acLine.StartPoint.Pos
                            End If
                        End If
                    ElseIf pk2 < pk0 Then
                        Dim tempdist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.EndPoint.Pos)
                        Dim tempdist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.StartPoint.Pos)
                        If tempdist > tempdist2 Then tempdist = tempdist2
                        If off2 < 0 Then
                            If tempdist > offsetLeftStart Then
                                offsetLeftStart = tempdist
                                offsetLeftPointStart = acLine.EndPoint.Pos
                            End If
                        Else
                            If tempdist > offsetRightStart Then
                                offsetRightStart = tempdist
                                offsetRightPointStart = acLine.EndPoint.Pos
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Dim offsetLeftEnd As Double = 0
        Dim offsetRightEnd As Double = 0

        For Each acLine As DwgLine In arrayLineElements
            Dim pointIntersect As Vector2D = New Vector2D()
            Dim boolIntersect As Boolean = MathFunction.FuncIntersectionTwoSegments(acLine.StartPoint.Pos, acLine.EndPoint.Pos, tempEndPointLeft.Pos, tempEndPointRight.Pos, pointIntersect)
            If boolIntersect = True Then
                Dim off0 As Double = 0
                Dim off1 As Double = 0
                Dim off2 As Double = 0
                Dim pk0 As Double = 0
                Dim pk1 As Double = 0
                Dim pk2 As Double = 0
                Dim boolPK0 As Boolean = axisPline0.PosToStaOffset(pointIntersect, pk0, off0)
                Dim boolPK1 As Boolean = axisPline0.PosToStaOffset(acLine.StartPoint.Pos, pk1, off1)
                Dim boolPK2 As Boolean = axisPline0.PosToStaOffset(acLine.EndPoint.Pos, pk2, off2)
                If boolPK0 = True And boolPK1 = True And boolPK2 = True Then
                    If pk1 > pk0 Then
                        Dim tempdist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.StartPoint.Pos)
                        Dim tempdist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.EndPoint.Pos)
                        If tempdist > tempdist2 Then tempdist = tempdist2
                        If off1 < 0 Then
                            If tempdist > offsetLeftEnd Then
                                offsetLeftEnd = tempdist
                                offsetLeftPointEnd = acLine.StartPoint.Pos
                            End If
                        Else
                            If tempdist > offsetRightEnd Then
                                offsetRightEnd = tempdist
                                offsetRightPointEnd = acLine.StartPoint.Pos
                            End If
                        End If
                    ElseIf pk2 > pk0 Then
                        Dim tempdist As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.EndPoint.Pos)
                        Dim tempdist2 As Double = MathFunction.funcCalcDistanceByToPoints2d(pointIntersect, acLine.StartPoint.Pos)
                        If tempdist > tempdist2 Then tempdist = tempdist2
                        If off2 < 0 Then
                            If tempdist > offsetLeftEnd Then
                                offsetLeftEnd = tempdist
                                offsetLeftPointEnd = acLine.EndPoint.Pos
                            End If
                        Else
                            If tempdist > offsetRightEnd Then
                                offsetRightEnd = tempdist
                                offsetRightPointEnd = acLine.EndPoint.Pos
                            End If
                        End If
                    End If
                End If
            End If
        Next
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If offsetLeftEnd > 0 And offsetLeftStart > 0 And offsetRightEnd > 0 And offsetRightStart > 0 Then
            tempStartLine.StartPoint = New Vector3D(offsetLeftPointStart, 0)
            tempStartLine.EndPoint = New Vector3D(offsetRightPointStart, 0)
            'удлинняем линию
            Dim boolExtLineStart As Boolean = FuncBridge.FuncExtendBearm(tempStartLine, 100, 100)
            tempEndLine.StartPoint = New Vector3D(offsetLeftPointEnd, 0)
            tempEndLine.EndPoint = New Vector3D(offsetRightPointEnd, 0)
            'удлинняем линию
            Dim boolExtLineend As Boolean = FuncBridge.FuncExtendBearm(tempEndLine, 100, 100)
        End If

        'рисуем ось трассы и габарит сооружения
        If IsNothing(acStartLine) = False And IsNothing(acEndLine) = False Then
            Dim pointIntersectCollection1 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, tempStartLine.StartPoint.Pos, tempStartLine.EndPoint.Pos)
            Dim PKStart As Double = -1
            Dim offStart As Double = -1
            Dim boolPKStart As Boolean = False
            If pointIntersectCollection1.Count > 0 Then
                Dim startPointAxis As Vector2D = pointIntersectCollection1(0)
                boolPKStart = acPoly3dAlign.PosToStaOffset(startPointAxis, PKStart, offStart)
            End If
            Dim pointIntersectCollection2 As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(acPoly3dAlign, tempEndLine.StartPoint.Pos, tempEndLine.EndPoint.Pos)
            Dim PKEnd As Double = -1
            Dim offEnd As Double = -1
            Dim boolPKEnd As Boolean = False
            If pointIntersectCollection2.Count > 0 Then
                Dim endPointAxis As Vector2D = pointIntersectCollection2(0)
                boolPKEnd = acPoly3dAlign.PosToStaOffset(endPointAxis, PKEnd, offEnd)
            End If
            If boolPKEnd = True And boolPKStart = True Then
                Dim newPoly3dAlign As Polyline3D = acPoly3dAlign.Trim(PKStart, PKEnd)
                Dim hgAlignBridje As UInteger = CUInt(arrayPropertiesBridge(4))
                If boolPline = True Then
                    axisOldAlign.Clear()
                    For i As Integer = 0 To newPoly3dAlign.Count - 1
                        Dim pos As BugleVector2D = New BugleVector2D(newPoly3dAlign.Item(i).Vertex.Pos, newPoly3dAlign.Item(i).Bugle)
                        axisOldAlign.Add(pos)
                    Next
                Else
                    Erase ArrayCoord
                    For i As Integer = 0 To newPoly3dAlign.Count - 1
                        Dim pos As BugleVector3D = newPoly3dAlign.Item(i)
                        ReDim Preserve ArrayCoord(3, i)
                        ArrayCoord(0, i) = pos.Vertex.X
                        ArrayCoord(1, i) = pos.Vertex.Y
                        ArrayCoord(2, i) = pos.Bugle
                    Next
                    If IsArray(ArrayCoord) = True Then
                        If ArrayCoord.GetUpperBound(1) > 0 Then
                            axisOldAlign = RoburFunc.FuncDrawPolylineToArrayCoord(activDoc, ArrayCoord)
                        End If
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            'рисуем габарит сооружения
            Dim PKStartDirect As Double = -1
            Dim PKStartReverse As Double = -1
            Dim PKEndDirect As Double = -1
            Dim PKEndReverse As Double = -1
            Dim pline2dCurv1 As Polyline2DCurve = New Polyline2DCurve()
            Dim pline2dCurv2 As Polyline2DCurve = New Polyline2DCurve()
            'виртуально удлинняем первую балку
            Dim startIntersectPoint1 As Vector3D = Nothing
            Dim startIntersectPoint2 As Vector3D = Nothing
            Dim boolExt1 As Boolean = FuncBridge.FuncVirtualExtendBearm(tempStartLine, 20, 20, startIntersectPoint1, startIntersectPoint2)

            Dim endIntersectPoint1 As Vector3D = Nothing
            Dim endIntersectPoint2 As Vector3D = Nothing
            Dim boolExt2 As Boolean = FuncBridge.FuncVirtualExtendBearm(tempEndLine, 20, 20, endIntersectPoint1, endIntersectPoint2)
            'ищем точку пересечения в начале трассы
            pointIntersectCollection1 = PolylineExtentions.GetIntersections(axisPline3DDirect, startIntersectPoint1.Pos, startIntersectPoint2.Pos)
            If pointIntersectCollection1.Count > 0 Then
                Dim pointSectStart As Vector2D = pointIntersectCollection1.ElementAt(0)
                Dim boolPK As Boolean = axisPlineDirect.PosToStaOffset(pointSectStart, PKStartDirect, offStart)
            End If
            'ищем точку пересечения в конце трассы
            pointIntersectCollection2 = PolylineExtentions.GetIntersections(axisPline3DDirect, endIntersectPoint1.Pos, endIntersectPoint2.Pos)
            If pointIntersectCollection2.Count > 0 Then
                Dim pointSectStart As Vector2D = pointIntersectCollection2.ElementAt(0)
                Dim boolPK As Boolean = axisPlineDirect.PosToStaOffset(pointSectStart, PKEndDirect, offEnd)
            End If

            'ищем точку пересечения в начале трассы
            pointIntersectCollection1 = PolylineExtentions.GetIntersections(axisPline3DReverse, startIntersectPoint1.Pos, startIntersectPoint2.Pos)
            If pointIntersectCollection1.Count > 0 Then
                Dim pointSectStart As Vector2D = pointIntersectCollection1.ElementAt(0)
                Dim boolPK As Boolean = axisPline3DReverse.PosToStaOffset(pointSectStart, PKEndReverse, offStart)
            End If
            'ищем точку пересечения в конце трассы
            pointIntersectCollection2 = PolylineExtentions.GetIntersections(axisPline3DReverse, endIntersectPoint1.Pos, endIntersectPoint2.Pos)
            If pointIntersectCollection2.Count > 0 Then
                Dim pointSectStart As Vector2D = pointIntersectCollection2.ElementAt(0)
                Dim boolPK As Boolean = axisPline3DReverse.PosToStaOffset(pointSectStart, PKStartReverse, offEnd)
            End If

            If PKStartDirect <> -1 And PKEndDirect <> -1 Then
                Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
                For i As Integer = 0 To axisPlineDirect.Count - 1
                    pline2dCurv.Add(axisPlineDirect.Item(i))
                Next i
                Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(PKStartDirect, PKEndDirect)
                If IsArray(arrayPlineCurve) = True Then
                    For i As Integer = 0 To arrayPlineCurve.Length - 1
                        Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                        If Math.Abs(plineCurv.Length - (PKEndDirect - PKStartDirect)) <= 0.001 Then
                            pline2dCurv1 = plineCurv
                        End If
                    Next i
                End If
            End If
            'право
            If PKStartReverse <> -1 And PKEndReverse <> -1 Then
                Dim pline2dCurv As Polyline2DCurve = New Polyline2DCurve()
                For i As Integer = 0 To axisPlineReverse.Count - 1
                    pline2dCurv.Add(axisPlineReverse.Item(i))
                Next i
                Dim arrayPlineCurve As Polyline2DCurve() = pline2dCurv.Break(PKStartReverse, PKEndReverse)
                If IsArray(arrayPlineCurve) = True Then
                    For i As Integer = 0 To arrayPlineCurve.Length - 1
                        Dim plineCurv As Polyline2DCurve = arrayPlineCurve(i)
                        If Math.Abs(plineCurv.Length - (PKEndReverse - PKStartReverse)) <= 0.001 Then
                            pline2dCurv2 = plineCurv
                        End If
                    Next i
                End If
            End If

            Dim boolBound As Boolean = False
            If pline2dCurv2.Length > 0 And pline2dCurv1.Length Then
                If IsNothing(hgBoundares) = False Then
                    Dim gb As DwgPolyline = New DwgPolyline()
                    Dim boolFindBoundares As Boolean = activDoc.ActiveSpace.Entities.TryGetObject(hgBoundares, gb)
                    If boolFindBoundares = True Then
                        gb.Clear()
                        For i As Integer = 0 To pline2dCurv1.Count - 1
                            gb.Add(pline2dCurv1.Item(i))
                        Next i
                        Dim count As Integer = pline2dCurv2.Count - 1
                        For i As Integer = 0 To pline2dCurv2.Count - 1
                            Dim bulg As Double = 0
                            If i <> 0 Then
                                bulg = pline2dCurv2.Item(i - 1).Bugle
                            End If
                            gb.Add(pline2dCurv2.Item(i))
                        Next i
                        boolBound = True
                    End If
                End If
            End If
            If boolBound = False Then
                Dim gb As DwgPolyline = New DwgPolyline()
                For i As Integer = 0 To pline2dCurv1.Count - 1
                    gb.Add(pline2dCurv1.Item(i))
                Next i
                Dim count As Integer = pline2dCurv2.Count - 1
                For i As Integer = 0 To pline2dCurv2.Count - 1
                    Dim bulg As Double = 0
                    If i <> 0 Then
                        bulg = pline2dCurv2.Item(i - 1).Bugle
                    End If
                    'Dim newVert As BugleVector2D = New BugleVector2D(pline2dCurv2.Item(i).Vertex, bulg)
                    gb.Add(pline2dCurv2.Item(i))
                Next i
                gb.Closed = True
                activDoc.ActiveSpace.Add(gb)
            End If
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'удаляем вспомогательные построения
        activDoc.ActiveSpace.Entities.Remove(axisPline0)
        activDoc.ActiveSpace.Entities.Remove(axisPlineDirect)
        activDoc.ActiveSpace.Entities.Remove(axisPlineReverse)
    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вставка крайней опоры
    Public Shared Sub PlacementExtremPillars(ByRef ActivDocument As Topomatic.Dwg.Drawing, ByRef dictinaryAllObjectBridge As Dictionary(Of String, String(,)), ByVal numberPillar As Integer, ByVal Align As Alignment, ByVal projectSurface As Surface, ByVal egSurface As Surface, ByVal idBridge As String, ByVal userBridge As Bridge, ByVal userNozzle As Bridge, ByVal templateXML As String)
        If dictinaryAllObjectBridge.Count = 0 Then
            MsgBox("Элементы мостового сооружения не найдены. Сооружение не построено!!!")
            Exit Sub
        End If
        Dim numberProlet As Integer = 1
        If numberPillar > 1 Then
            numberPillar = userBridge.countProlet + 1
            numberProlet = userBridge.countProlet
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. Находим ось крайней опоры
        If dictinaryAllObjectBridge.ContainsKey("Ось опоры") = False Then
            MsgBox("В выбранном сооружении отсутствуют балки пролетных строений. Сооружение не построено!!!")
            Exit Sub
        End If
        Dim userPillar As AxisHorizontalPillars = Nothing
        Dim axisLinePillar As DwgLine = Nothing
        Dim arrayPillars As String(,) = dictinaryAllObjectBridge.Item("Ось опоры")
        If IsArray(arrayPillars) = True Then
            For i As Integer = 0 To arrayPillars.GetUpperBound(1)
                Dim keyParam As String = arrayPillars(1, i)
                If IsNothing(keyParam) = False Then
                    If keyParam.Trim.Length > 0 Then
                        Try
                            userPillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of AxisHorizontalPillars)(keyParam)
                        Catch ex As Newtonsoft.Json.JsonException
                        End Try
                    End If
                End If
                If IsNothing(userPillar) = False Then
                    Dim tempNumberPillar As Integer = userPillar.number
                    If tempNumberPillar = numberPillar Then
                        Dim hgPillar As UInteger = CUInt(arrayPillars(4, i))
                        Dim boolFindPillar As Boolean = ActivDocument.ActiveSpace.Entities.TryGetObject(hgPillar, axisLinePillar)
                        If boolFindPillar = True Then
                            Exit For
                        End If
                    End If
                End If
            Next i
        End If
        If IsNothing(axisLinePillar) = True Then
            MsgBox("Ось крайней опоры №" & numberPillar & ", не найдена. Крайняя опора не построена!!!")
            Exit Sub
        End If
        If IsNothing(userPillar) = True Then
            MsgBox("Характеристики} крайней опоры №" & numberPillar & ", не найдены. Крайняя опора не построена!!!")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'строим насадку



        '1. Оформление
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'насадка
        Dim layerAxisBeam As DwgLayer = ActivDocument.ActiveLayer
        Dim colorAxisBeam As CadColor = New CadColor(7)
        Dim nameTypeLineAxisBeam As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineAxisBeam As Integer = 1
        Dim widthTypeLineAxisBeam As Integer = 20
        Dim tablePSAxisBeam As String = ""
        'балка
        Dim layerBeam As DwgLayer = ActivDocument.ActiveLayer
        Dim colorBeam As CadColor = New CadColor(7)
        Dim nameTypeLineBeam As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineBeam As Integer = 1
        Dim widthTypeLineBeam As Integer = 20
        Dim tablePSBeam As String = ""
        'Ось опоры
        Dim layerAxisPillar As DwgLayer = ActivDocument.ActiveLayer
        Dim colorAxisPillar As CadColor = New CadColor(7)
        Dim nameTypeLineAxisPillar As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineAxisPillar As Integer = 1
        Dim widthTypeLineAxisPillar As Integer = 20
        Dim tablePSAxisPillar As String = ""
        'Ось опирания балок
        Dim layerAxisBeamPillar As DwgLayer = ActivDocument.ActiveLayer
        Dim colorAxisBeamPillar As CadColor = New CadColor(7)
        Dim nameTypeLineAxisBeamPillar As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineAxisBeamPillar As Integer = 1
        Dim widthTypeLineAxisBeamPillar As Integer = 20
        Dim tablePSAxisBeamPillar As String = ""
        'Верх ребра балки
        Dim layerUpBeam As DwgLayer = ActivDocument.ActiveLayer
        Dim colorUpBeam As CadColor = New CadColor(7)
        Dim nameTypeLineUpBeam As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineUpBeam As Integer = 1
        Dim widthTypeLineUpBeam As Integer = 20
        Dim tablePSUpBeam As String = ""
        'Низ ребра балки
        Dim layerDownBeam As DwgLayer = ActivDocument.ActiveLayer
        Dim colorDownBeam As CadColor = New CadColor(7)
        Dim nameTypeLineDownBeam As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineDownBeam As Integer = 1
        Dim widthTypeLineDownBeam As Integer = 20
        Dim tablePSDownBeam As String = ""
        'главная ось путепровода
        Dim layerAxisBridge As DwgLayer = ActivDocument.ActiveLayer
        Dim colorAxisBridge As CadColor = New CadColor(7)
        Dim nameTypeLineAxisBridge As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineAxisBridge As Integer = 1
        Dim widthTypeLineAxisBridge As Integer = 20
        Dim tablePSAxisBridge As String = ""
        'граница путепровода
        Dim layerStructureBridge As DwgLayer = ActivDocument.ActiveLayer
        Dim colorStructureBridge As CadColor = New CadColor(5)
        Dim nameTypeLineStructureBridge As DwgLinetype = ActivDocument.ActiveLinetype
        Dim ScaleTypeLineStructureBridge As Integer = 1
        Dim widthTypeLineStructureBridge As Integer = 20
        Dim tablePSStructureBridge As String = ""
        '=======================================================================================================
        Dim categoryTables As String = "Искусственные сооружения"
        Dim nameTableBridge As String = "Мостовое сооружение"
        Dim nameTablePillars As String = "Опоры мостовых сооружений"
        Dim nameTableBeams As String = "Балки мостовых сооружений"



    End Sub
End Class
