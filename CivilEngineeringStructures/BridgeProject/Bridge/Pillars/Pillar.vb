
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media.Animation
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
'опора моста
Public Class Pillar
    Public Enum PillarType
        <Description("Крайняя опора")> LastPillar = 0
        <Description("Промежуточния опора")> MiddlePillar = 1
        <Description("Не определено")> None = 2
    End Enum
    ' Перечисление для левого или правого объекта
    Public Enum SidePillarElement
        <Description("Левый элемент")> Left = 0
        <Description("Правый элемент")> Right = 1
        <Description("Не определено")> None = 2
    End Enum
    'Inherits Bridges
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private _number As Integer 'номер опоры
    Private _defining As Boolean 'опора явлется определяющей
    Private _typePillar As PillarType 'крайняя опора\промежуточная опора
    Private _clearence As Double 'зазор
    Private _rightClearence As Double 'правый зазор
    Private _siteMonolit As Double 'участок омоноличивания балок
    Private _elevationLand As Double 'отметка земли
    Private _countSubPillars As Integer 'число подопор
    Private _singleSubFarmer As Boolean 'единый подферменник
    Private _fixedHeightCabinetWall As Boolean 'фиксированная шкафная стенка
    Private _presenceRacks As Boolean 'наличие стоек
    Private _presencGrillage As Boolean 'наличие ростверка
    Private _presencPreparation As Boolean 'наличие подготовки
    Private _pileInRack As Boolean 'вставить сваю в стойку
    Private _horizontalLevel As Boolean 'ось опоры горизонтальна
    Public _elementBridgePoint As PointsCollections
    'Private _entityPillar As DwgLine
    Public Sub New()
        _number = 0
        _defining = False
        _typePillar = PillarType.None
        _clearence = 0
        _rightClearence = 0
        _siteMonolit = 0
        _elevationLand = 0
        _countSubPillars = 1
        _singleSubFarmer = True
        _fixedHeightCabinetWall = False
        _presenceRacks = True
        _presencGrillage = True
        _presencPreparation = True
        _pileInRack = False
        _horizontalLevel = True
        _elementBridgePoint = New PointsCollections
    End Sub

    ' Деструктор класса
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property Number() As Integer
        Get
            Return _number
        End Get
        Set(value As Integer)
            _number = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Опора является определяющей")>
    <Category("Свойства")>
    <DisplayName("Определяющая опора")>
    Public Property Defining() As Boolean
        Get
            Return _defining
        End Get
        Set(value As Boolean)
            _defining = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Тип опоры (кряйняя или промежуточная)")>
    <Category("Свойства")>
    <DisplayName("Тип опоры")>
    Public Property TypePillar() As PillarType
        Get
            Return _typePillar
        End Get
        Set(value As PillarType)
            _typePillar = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Зазор между балками между соседними пролетами, м")>
    <Category("Свойства")>
    <DisplayName("Зазор между балками")>
    Public Property Clearence() As Double
        Get
            Return _clearence
        End Get
        Set(value As Double)
            _clearence = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Зазор (правый) между балками между соседними пролетами, м")>
    <Category("Свойства")>
    <DisplayName("Правый зазор")>
    Public Property RightClearence() As Double
        Get
            Return _rightClearence
        End Get
        Set(value As Double)
            _rightClearence = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Отметка земли, м")>
    <Category("Свойства")>
    <DisplayName("Отметка земли")>
    Public Property ElevationLand() As Double
        Get
            Return _elevationLand
        End Get
        Set(value As Double)
            _elevationLand = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Толщина участка омоноличивания балок, м")>
    <Category("Свойства")>
    <DisplayName("Толщина омоноличивания")>
    Public Property SiteMonolit() As Double
        Get
            Return _siteMonolit
        End Get
        Set(value As Double)
            _siteMonolit = value
        End Set
    End Property
    Public Property CountSubPillars() As Integer
        Get
            Return _countSubPillars
        End Get
        Set(value As Integer)
            _countSubPillars = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Единый подферменник")>
    <Category("Свойства")>
    <DisplayName("Единый подферменник")>
    Public Property SingleSubFarmer() As Boolean
        Get
            Return _singleSubFarmer
        End Get
        Set(value As Boolean)
            _singleSubFarmer = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Фиксированная высота шкафной стенки")>
    <Category("Свойства")>
    <DisplayName("Фиксировать высоту шк.стенки")>
    Public Property FixedHeightCabinetWall() As Boolean
        Get
            Return _fixedHeightCabinetWall
        End Get
        Set(value As Boolean)
            _fixedHeightCabinetWall = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Наличие стоек у опоры")>
    <Category("Свойства")>
    <DisplayName("Наличие стоек")>
    Public Property PresenceRacks() As Boolean
        Get
            Return _presenceRacks
        End Get
        Set(value As Boolean)
            _presenceRacks = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Наличие ростверка у опоры")>
    <Category("Свойства")>
    <DisplayName("Наличие ростверка")>
    Public Property PresencGrillage() As Boolean
        Get
            Return _presencGrillage
        End Get
        Set(value As Boolean)
            _presencGrillage = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Наличие подготовки у опоры")>
    <Category("Свойства")>
    <DisplayName("Наличие подготовки")>
    Public Property PresencPreparation() As Boolean
        Get
            Return _presencPreparation
        End Get
        Set(value As Boolean)
            _presencPreparation = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Стойки опоры находится в свае")>
    <Category("Свойства")>
    <DisplayName("Стойка в свае")>
    Public Property PileInRack() As Boolean
        Get
            Return _pileInRack
        End Get
        Set(value As Boolean)
            _pileInRack = value
        End Set
    End Property

    ' Свойство для доступа к признаку "горизонтальная ось"
    <Browsable(False)>
    Public Property HorizontalLevel() As Boolean
        Get
            Return _horizontalLevel
        End Get
        Set(value As Boolean)
            _horizontalLevel = value
        End Set
    End Property
    'создать новую структуру для оси опоры
    Public Shared Function createAxis(ByVal idBridge As String, typePillar As StructureElement.classStructure) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementAxis.ClassObject = typePillar
        elementAxis.Name = StructureElement.typeObject.axisPillar
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.axisPillar)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgLine
        Return elementAxis
    End Function
    'ищет существующую ось опоры
    Public Shared Function getAxisPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer) As StructureElement
        Dim dataAxisPillar As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
            Dim listAxisPillar = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillar)
            If IsNothing(listAxisPillar) = False Then
                If listAxisPillar.Count > 0 Then
                    For k As Integer = 0 To listAxisPillar.Count - 1
                        Dim tempData As StructureElement = listAxisPillar.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisPillar As Pillar = tempData.getPillar()
                            If IsNothing(userAxisPillar) = False Then
                                If numberPillar = userAxisPillar.Number Then
                                    dataAxisPillar = tempData
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataAxisPillar
    End Function
    'функция удаляет оси опоры и оси опирания балок в случае уменьшения числа пролетов сооружения
    Public Shared Function removeAxisPillarFromBridge(ByVal userBridge As Bridges, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Boolean
        Dim result As Boolean = False
        If IsNothing(userBridge) = True Then Return result
        If userBridge.ProletCount <= 0 Then Return result
        If IsNothing(dictionaryObjectsBridge) = True Then Return result
        'удаляем оси опор
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
            Dim listAxisPillar As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillar)
            If IsNothing(listAxisPillar) = False Then
                If listAxisPillar.Count > 0 Then
                    For i As Integer = 0 To listAxisPillar.Count - 1
                        Dim dataAxisPillar As StructureElement = listAxisPillar.Item(i)
                        If IsNothing(dataAxisPillar.DWGEntity) = False Then
                            Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                            If IsNothing(userAxisPillar) = False Then
                                If userAxisPillar.Number > userBridge.ProletCount + 1 Then
                                    Dim activDoc As Dwg.Drawing = dataAxisPillar.DWGEntity.Drawing
                                    activDoc.ActiveSpace.Entities.Remove(dataAxisPillar.DWGEntity)
                                    result = True
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        'удаляем оси опирания балок
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillarBeams) = True Then
            Dim listAxisBeamsPillar As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillarBeams)
            If IsNothing(listAxisBeamsPillar) = False Then
                If listAxisBeamsPillar.Count > 0 Then
                    For i As Integer = 0 To listAxisBeamsPillar.Count - 1
                        Dim dataAxisBeamPillar As StructureElement = listAxisBeamsPillar.Item(i)
                        If IsNothing(dataAxisBeamPillar.DWGEntity) = False Then
                            Dim userAxisBeamPillar As AxisBeamsPillars = dataAxisBeamPillar.getAxisBeamsPillar
                            If IsNothing(userAxisBeamPillar) = False Then
                                If userAxisBeamPillar.numberProlet > userBridge.ProletCount Then
                                    Dim activDoc As Dwg.Drawing = dataAxisBeamPillar.DWGEntity.Drawing
                                    activDoc.ActiveSpace.Entities.Remove(dataAxisBeamPillar.DWGEntity)
                                    result = True
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция проверяет и корректирует ось опоры если неверно ее направление
    Public Shared Function correctDirectionAxisPillar(ByRef axisPillar As DwgLine, ByRef align As Alignment) As Boolean
        correctDirectionAxisPillar = False
        If IsNothing(axisPillar) = True Then
            Return False
        End If
        If IsNothing(align) = True Then
            Return False
        End If
        If axisPillar.Length = 0 Then
            Return False
        End If
        'находим пикеты начала и конца предудущей балки
        Try
            Dim pkStart As Double = 0
            Dim offStart As Double = 0
            Dim boolFindPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisPillar.StartPoint.Pos, pkStart, offStart)
            Dim pkEnd As Double = 0
            Dim offend As Double = 0
            Dim boolFindPk1 As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisPillar.EndPoint.Pos, pkEnd, offend)
            If boolFindPk = True And boolFindPk1 = True Then
                If offStart > 0 And offend < 0 Then
                    Dim tempStartPoint As Vector3D = axisPillar.StartPoint
                    Dim tempEndPoint As Vector3D = axisPillar.EndPoint
                    axisPillar.StartPoint = tempEndPoint
                    axisPillar.EndPoint = tempStartPoint
                    Return True
                ElseIf offStart < 0 And offend > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'рисование оси опоры
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleAxisPillar As ProjectCivilStructuresStyle = Nothing, Optional ByVal templateXML As String = "") As StructureElement
        Dim axisLinePillar As DwgLine = Nothing
        Dim dataStructureBeamsPillar As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim dataAxisPillar As StructureElement = getAxisPillar(dictionaryBridgeElements, Number)
        If IsNothing(dataAxisPillar) = True Then
            If TypePillar = PillarType.LastPillar Then
                dataAxisPillar = createAxis(idBridge, StructureElement.classStructure.LastPillar)
            Else
                dataAxisPillar = createAxis(idBridge, StructureElement.classStructure.MiddlePillar)
            End If
        End If
        If IsNothing(dataAxisPillar) Then Return Nothing
        axisLinePillar = dataAxisPillar.DWGEntity
        If IsNothing(axisLinePillar) = True Then Return Nothing
        If IsNothing(styleAxisPillar) = True And File.Exists(templateXML) = True Then
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            styleAxisPillar = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisPillar.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опоры")
            styleAxisPillar.setObjectStyle(axisLinePillar)
        End If
        'ось опирания балок
        axisLinePillar.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLinePillar.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxis As Double = (axisLinePillar.StartPoint.Pos - axisLinePillar.EndPoint.Pos).Length
        If lenghtAxis = 0 Then
            MsgBox("Ось опоры имеет нулевое значение.")
            Return dataAxisPillar
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLinePillar) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLinePillar)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataAxisPillar.KeyParameter = strJson
        dataAxisPillar.DWGEntity = axisLinePillar
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLinePillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataAxisPillar)
        Return dataAxisPillar
    End Function
    'функция перемещает ось опоры вдоль трассы на указанную величину
    Public Shared Function moveAxisPillar(ByVal alignPline3D As Polyline3D, ByRef axisPillar As DwgLine, ByVal offset As Double) As Boolean
        moveAxisPillar = False
        If IsNothing(alignPline3D) = True Then Return False
        If alignPline3D.Length2D = 0 Then Return False
        If IsNothing(axisPillar) = True Then Return False
        If axisPillar.Length = 0 Then Return False
        If offset = 0 Then Return True
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'поиск объектов в общей библиотеке элементов
    'ищет определенный элемент опоры
    Public Shared Function getElementPillar(ByVal typeElement As StructureElement.typeObject, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberRow As Integer = 0, Optional ByVal numberColumn As Integer = 0, Optional ByVal numberSubPillar As Integer = 0, Optional sideObject As SidePillarElement = SidePillarElement.None) As List(Of StructureElement)
        Dim listDataElements As New List(Of StructureElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(typeElement) = True Then
            Dim listAxisElements = dictionaryObjectsBridge.Item(typeElement)
            If IsNothing(listAxisElements) = False Then
                If listAxisElements.Count > 0 Then
                    For k As Integer = 0 To listAxisElements.Count - 1
                        Dim tempData As StructureElement = listAxisElements.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim tempNumberPillar As Integer = 0
                            Dim tempNumberSubPillar As Integer = 0
                            Dim tempSideObject As SidePillarElement = SidePillarElement.None
                            Dim tempNumberRow As Integer = 0
                            Dim tempNumberColumn As Integer = 0
                            If typeElement = StructureElement.typeObject.axisNozzle Then
                                Dim userAxisObject As NozzlePillar = tempData.getNozzlePillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisRigel Then
                                Dim userAxisObject As RigelPillar = tempData.getRigelPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisCabinetWall Then
                                Dim userAxisObject As CabinetWallPillar = tempData.getCabinetWallPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisGrillage Then
                                Dim userAxisObject As GrillagePillar = tempData.getGrillagePillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisLeftHand Then
                                Dim userAxisObject As HandPillar = tempData.getHandPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillar
                                tempSideObject = userAxisObject.SideHand
                            ElseIf typeElement = StructureElement.typeObject.axisRightHand Then
                                Dim userAxisObject As HandPillar = tempData.getHandPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillar
                                tempSideObject = userAxisObject.SideHand
                            ElseIf typeElement = StructureElement.typeObject.axisLeftPostcard Then
                                Dim userAxisObject As PostcardPillar = tempData.getPostcardPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillar
                                tempSideObject = userAxisObject.SidePostcard
                            ElseIf typeElement = StructureElement.typeObject.axisrightPostcard Then
                                Dim userAxisObject As PostcardPillar = tempData.getPostcardPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillar
                                tempSideObject = userAxisObject.SidePostcard
                            ElseIf typeElement = StructureElement.typeObject.axisPreparation Then
                                Dim userAxisObject As PreparationPillar = tempData.getPreparationPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisRack Then
                                Dim userAxisObject As RackPillar = tempData.getRackPillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillars
                                tempNumberColumn = userAxisObject.Number
                            ElseIf typeElement = StructureElement.typeObject.axisSubFermenters Then
                                Dim userAxisObject As SubFermenters = tempData.getSubFermenters
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillar
                                tempNumberColumn = userAxisObject.NumberProlet
                                tempNumberRow = userAxisObject.NumberRow
                            ElseIf typeElement = StructureElement.typeObject.axisPile Then
                                Dim userAxisObject As PilePillar = tempData.getPilePillar
                                tempNumberPillar = userAxisObject.NumberPillar
                                tempNumberSubPillar = userAxisObject.NumberSubPillars
                                tempNumberColumn = userAxisObject.NumberColumn
                                tempNumberRow = userAxisObject.NumberRow
                            End If
                            'начинаем сравненеие
                            Dim boolNumberPillar As Boolean = True
                            Dim boolNumberSubPillar As Boolean = True
                            Dim boolSideObject As Boolean = True
                            Dim boolNumberRow As Boolean = True
                            Dim boolNumberColumn As Boolean = True
                            If numberPillar <> tempNumberPillar Then
                                boolNumberPillar = False
                            End If
                            If numberSubPillar > 0 Then
                                If numberSubPillar <> tempNumberSubPillar Then
                                    boolNumberSubPillar = False
                                End If
                            End If
                            If boolSideObject <> Pillar.SidePillarElement.None Then
                                If sideObject <> tempSideObject Then
                                    boolSideObject = False
                                End If
                            End If
                            If numberRow > 0 Then
                                If numberRow <> tempNumberRow Then
                                    boolNumberRow = False
                                End If
                            End If
                            If numberColumn > 0 Then
                                If numberColumn <> tempNumberColumn Then
                                    boolNumberColumn = False
                                End If
                            End If
                            If boolNumberColumn = True And boolNumberPillar = True And boolNumberRow = True And boolNumberSubPillar = True And boolSideObject = True Then
                                listDataElements.Add(tempData)
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return listDataElements
    End Function
    'функция возвращает все балки мостового сооружения для выбранной опоры
    Public Shared Function getBeamsPillarByNumber(ByVal numberPillar As Integer, ByVal dictionaryElementBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As List(Of Dictionary(Of Integer, StructureElement))
        Dim result As List(Of Dictionary(Of Integer, StructureElement)) = New List(Of Dictionary(Of Integer, StructureElement))
        Dim lastBeamsProlet As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        Dim nextBeamsProlet As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
        If numberPillar < 1 Then Return result
        If IsNothing(dictionaryElementBridge) = True Then Return result
        If dictionaryElementBridge.Count = 0 Then Return result
        If dictionaryElementBridge.ContainsKey(StructureElement.typeObject.axisBeam) = True Then
            Dim listBeams As List(Of StructureElement) = dictionaryElementBridge.Item(StructureElement.typeObject.axisBeam)
            If IsNothing(listBeams) = False Then
                If listBeams.Count > 0 Then
                    For i As Integer = 0 To listBeams.Count - 1
                        Dim dataBeam As StructureElement = listBeams.Item(i)
                        Dim keyParamBeam As String = dataBeam.KeyParameter
                        Dim userBeam As BeamI = dataBeam.getBeamI()
                        If IsNothing(userBeam) = False Then
                            Dim numColl As Integer = userBeam.numberProlet
                            Dim numRow As Integer = userBeam.numberRow
                            'если номер пролета равен номеру опоры или номер пролета на 1 меньше номера опоры
                            If numColl = numberPillar Then
                                If nextBeamsProlet.ContainsKey(numRow) = False Then
                                    nextBeamsProlet.Add(numRow, dataBeam)
                                End If
                            ElseIf numColl = numberPillar - 1 Then
                                If lastBeamsProlet.ContainsKey(numRow) = False Then
                                    lastBeamsProlet.Add(numRow, dataBeam)
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        If lastBeamsProlet.Count > 1 Then
            lastBeamsProlet = lastBeamsProlet.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        If lastBeamsProlet.Count > 0 Then

        End If
        If nextBeamsProlet.Count > 1 Then
            nextBeamsProlet = nextBeamsProlet.OrderBy(Function(pair) pair.Key).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)
        End If
        If nextBeamsProlet.Count > 0 Then

        End If
        result.Add(lastBeamsProlet)
        result.Add(nextBeamsProlet)
        Return result
    End Function
    'функция удаляет лишние модели
    Public Shared Function removeModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Boolean
        If IsNothing(dictionaryObjectsBridge) = True Then Return False
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelPile) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelPile)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim model As DwgModel3DElement = tempData.DWGEntity
                            If activProjectDocument.ActiveSpace.Entities.Contains(model) = True Then
                                activProjectDocument.ActiveSpace.Entities.Remove(model)
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return False
    End Function
    'функция удаляет лишние точки со списка
    Public Shared Function removePointInList(ByVal dictProjectPoint As List(Of Dictionary(Of String, ProjectionPoint))) As List(Of Dictionary(Of String, ProjectionPoint))
        Dim result As List(Of Dictionary(Of String, ProjectionPoint)) = New List(Of Dictionary(Of String, ProjectionPoint))
        Dim destTransformTop As New Dictionary(Of String, ProjectionPoint)
        Dim destTransformBottom As New Dictionary(Of String, ProjectionPoint)
        If IsNothing(dictProjectPoint) = True Then Return result
        If dictProjectPoint.Count = 2 Then
            Dim transformTop As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(0)
            Dim transformBottom As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(1)
            If IsNothing(transformTop) = False Then
                If transformTop.Count > 3 Then
                    Dim t1 As Vector3D = New Vector3D
                    Dim t2 As Vector3D = New Vector3D
                    Dim t3 As Vector3D = New Vector3D
                    Dim code As String = ""
                    For i As Integer = 0 To transformTop.Count - 1
                        code = transformTop.ElementAt(i).Key
                        If i = 0 Then
                            t1 = transformTop.ElementAt(transformTop.Count - 1).Value.originPoint
                            t2 = transformTop.ElementAt(i).Value.originPoint
                            t3 = transformTop.ElementAt(i + 1).Value.originPoint
                        ElseIf i = transformTop.Count - 1 Then
                            t1 = transformTop.ElementAt(i - 1).Value.originPoint
                            t2 = transformTop.ElementAt(i).Value.originPoint
                            t3 = transformTop.ElementAt(0).Value.originPoint
                        Else
                            t1 = transformTop.ElementAt(i - 1).Value.originPoint
                            t2 = transformTop.ElementAt(i).Value.originPoint
                            t3 = transformTop.ElementAt(i + 1).Value.originPoint
                        End If
                        Dim boolPointInLine As Boolean = BridgeGeometry.crossPointInLine(t1, t3, t2)
                        If boolPointInLine = False Then
                            destTransformTop.Add(code, transformTop.ElementAt(i).Value)
                        End If
                    Next i
                End If
            End If
            If IsNothing(transformBottom) = False Then
                If transformBottom.Count > 3 Then
                    Dim t1 As Vector3D = New Vector3D
                    Dim t2 As Vector3D = New Vector3D
                    Dim t3 As Vector3D = New Vector3D
                    Dim code As String = ""
                    For i As Integer = 0 To transformBottom.Count - 1
                        code = transformBottom.ElementAt(i).Key
                        If i = 0 Then
                            t1 = transformBottom.ElementAt(transformBottom.Count - 1).Value.originPoint
                            t2 = transformBottom.ElementAt(i).Value.originPoint
                            t3 = transformBottom.ElementAt(i + 1).Value.originPoint
                        ElseIf i = transformBottom.Count - 1 Then
                            t1 = transformBottom.ElementAt(i - 1).Value.originPoint
                            t2 = transformBottom.ElementAt(i).Value.originPoint
                            t3 = transformBottom.ElementAt(0).Value.originPoint
                        Else
                            t1 = transformBottom.ElementAt(i - 1).Value.originPoint
                            t2 = transformBottom.ElementAt(i).Value.originPoint
                            t3 = transformBottom.ElementAt(i + 1).Value.originPoint
                        End If
                        Dim boolPointInLine As Boolean = BridgeGeometry.crossPointInLine(t1, t3, t2)
                        If boolPointInLine = False Then
                            destTransformBottom.Add(code, transformBottom.ElementAt(i).Value)
                        End If
                    Next i
                End If
            End If
        End If
        result.Add(destTransformTop)
        result.Add(destTransformBottom)
        Return result
    End Function
    'функция готовит точки для рисования фигуры
    Public Shared Function selectPointForDraw(ByVal dictProjectPoint As List(Of Dictionary(Of String, ProjectionPoint)), ByVal prjView As ProjectionPoint.projectView) As Dictionary(Of String, ProjectionPoint)
        Dim result As New Dictionary(Of String, ProjectionPoint)
        If IsNothing(dictProjectPoint) = False Then
            If dictProjectPoint.Count = 2 Then
                Dim transformTop As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(0)
                Dim transformBottom As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(1)
                If prjView = ProjectionPoint.projectView.Top Then
                    result = transformTop
                ElseIf prjView = ProjectionPoint.projectView.Bottom Then
                    result = transformBottom
                ElseIf prjView = ProjectionPoint.projectView.Front Then
                    For i As Integer = 0 To transformTop.Count - 1
                        Dim code As String = transformTop.ElementAt(i).Key
                        If code.IndexOf("left") > -1 Then
                            result.Add(code & "-top", transformTop.ElementAt(i).Value)
                        End If
                    Next i
                    For i As Integer = transformBottom.Count - 1 To 0 Step -1
                        Dim code As String = transformBottom.ElementAt(i).Key
                        If code.IndexOf("left") > -1 Then
                            result.Add(code & "-bottom", transformBottom.ElementAt(i).Value)
                        End If
                    Next i
                ElseIf prjView = ProjectionPoint.projectView.Back Then
                    For i As Integer = 0 To transformTop.Count - 1
                        Dim code As String = transformTop.ElementAt(i).Key
                        If code.IndexOf("right") > -1 Then
                            result.Add(code & "-top", transformTop.ElementAt(i).Value)
                        End If
                    Next i
                    For i As Integer = transformBottom.Count - 1 To 0 Step -1
                        Dim code As String = transformBottom.ElementAt(i).Key
                        If code.IndexOf("right") > -1 Then
                            result.Add(code & "-bottom", transformBottom.ElementAt(i).Value)
                        End If
                    Next i
                ElseIf prjView = ProjectionPoint.projectView.Left Then
                    If transformTop.ContainsKey("leftPt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("leftPt1")
                        result.Add("leftPt1-top", point)
                    End If
                    If transformTop.ContainsKey("middlePt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("middlePt1")
                        result.Add("middlePt1-top", point)
                    End If
                    If transformTop.ContainsKey("rightPt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("rightPt1")
                        result.Add("rightPt1-top", point)
                    End If
                    If transformBottom.ContainsKey("rightPt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("rightPt1")
                        result.Add("rightPt1-bottom", point)
                    End If
                    If transformBottom.ContainsKey("middlePt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("middlePt1")
                        result.Add("middlePt1-bottom", point)
                    End If
                    If transformBottom.ContainsKey("leftPt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("leftPt1")
                        result.Add("leftPt1-bottom", point)
                    End If
                ElseIf prjView = ProjectionPoint.projectView.Right Then
                    If transformTop.ContainsKey("leftPt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("leftPt2")
                        result.Add("leftPt2-top", point)
                    End If
                    If transformTop.ContainsKey("middlePt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("middlePt2")
                        result.Add("middlePt2-top", point)
                    End If
                    If transformTop.ContainsKey("rightPt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("rightPt2")
                        result.Add("rightPt2-top", point)
                    End If
                    If transformBottom.ContainsKey("rightPt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("rightPt2")
                        result.Add("rightPt2-bottom", point)
                    End If
                    If transformBottom.ContainsKey("middlePt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("middlePt2")
                        result.Add("middlePt2-bottom", point)
                    End If
                    If transformBottom.ContainsKey("leftPt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("leftPt2")
                        result.Add("leftPt2-bottom", point)
                    End If
                End If
            End If
        End If
        Return result
    End Function
    'функция готовит точки для рисования фигуры
    Public Shared Function selectPointCircleRackForDraw(ByVal dictProjectPoint As List(Of Dictionary(Of String, ProjectionPoint)), ByVal prjView As ProjectionPoint.projectView) As Dictionary(Of String, ProjectionPoint)
        Dim result As New Dictionary(Of String, ProjectionPoint)
        If IsNothing(dictProjectPoint) = False Then
            If dictProjectPoint.Count = 2 Then
                Dim transformTop As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(0)
                Dim transformBottom As Dictionary(Of String, ProjectionPoint) = dictProjectPoint.Item(1)
                If prjView = ProjectionPoint.projectView.Top Then
                    result = transformTop
                ElseIf prjView = ProjectionPoint.projectView.bottom Then
                    result = transformBottom
                ElseIf prjView = ProjectionPoint.projectView.front Then
                    For i As Integer = 0 To transformTop.Count - 1
                        Dim code As String = transformTop.ElementAt(i).Key
                        If code.IndexOf("frontPt") > -1 Then
                            result.Add(code & "-top", transformTop.ElementAt(i).Value)
                        End If
                    Next i
                    For i As Integer = transformBottom.Count - 1 To 0 Step -1
                        Dim code As String = transformBottom.ElementAt(i).Key
                        If code.IndexOf("left") > -1 Then
                            result.Add(code & "-bottom", transformBottom.ElementAt(i).Value)
                        End If
                    Next i
                ElseIf prjView = ProjectionPoint.projectView.back Then
                    For i As Integer = 0 To transformTop.Count - 1
                        Dim code As String = transformTop.ElementAt(i).Key
                        If code.IndexOf("right") > -1 Then
                            result.Add(code & "-top", transformTop.ElementAt(i).Value)
                        End If
                    Next i
                    For i As Integer = transformBottom.Count - 1 To 0 Step -1
                        Dim code As String = transformBottom.ElementAt(i).Key
                        If code.IndexOf("right") > -1 Then
                            result.Add(code & "-bottom", transformBottom.ElementAt(i).Value)
                        End If
                    Next i
                ElseIf prjView = ProjectionPoint.projectView.left Then
                    If transformTop.ContainsKey("leftPt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("leftPt1")
                        result.Add("leftPt1-top", point)
                    End If
                    If transformTop.ContainsKey("middlePt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("middlePt1")
                        result.Add("middlePt1-top", point)
                    End If
                    If transformTop.ContainsKey("rightPt1") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("rightPt1")
                        result.Add("rightPt1-top", point)
                    End If
                    If transformBottom.ContainsKey("rightPt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("rightPt1")
                        result.Add("rightPt1-bottom", point)
                    End If
                    If transformBottom.ContainsKey("middlePt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("middlePt1")
                        result.Add("middlePt1-bottom", point)
                    End If
                    If transformBottom.ContainsKey("leftPt1") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("leftPt1")
                        result.Add("leftPt1-bottom", point)
                    End If
                ElseIf prjView = ProjectionPoint.projectView.right Then
                    If transformTop.ContainsKey("leftPt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("leftPt2")
                        result.Add("leftPt2-top", point)
                    End If
                    If transformTop.ContainsKey("middlePt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("middlePt2")
                        result.Add("middlePt2-top", point)
                    End If
                    If transformTop.ContainsKey("rightPt2") = True Then
                        Dim point As ProjectionPoint = transformTop.Item("rightPt2")
                        result.Add("rightPt2-top", point)
                    End If
                    If transformBottom.ContainsKey("rightPt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("rightPt2")
                        result.Add("rightPt2-bottom", point)
                    End If
                    If transformBottom.ContainsKey("middlePt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("middlePt2")
                        result.Add("middlePt2-bottom", point)
                    End If
                    If transformBottom.ContainsKey("leftPt2") = True Then
                        Dim point As ProjectionPoint = transformBottom.Item("leftPt2")
                        result.Add("leftPt2-bottom", point)
                    End If
                End If
            End If
        End If
        Return result
    End Function
    'функция рисует контур в pictureBox
    Public Shared Function drawCounterPictureBox(ByRef bmp As Bitmap, ByVal projectPoint As Dictionary(Of String, ProjectionPoint), ByVal k As Double, ByVal offsetX As Single, ByVal offsetY As String, ByVal prjView As ProjectionPoint.projectView, Optional drawDimText As Boolean = False, Optional drawElevationText As Boolean = False) As Boolean
        Dim colorPen As System.Drawing.Color = System.Drawing.Color.Black
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            If projectPoint.Count > 2 Then
                Dim codeStartPoint As String = ""
                Dim codeEndPoint As String = ""
                For i As Integer = 0 To projectPoint.Count - 1
                    Dim sourceStartPoint As ProjectionPoint = projectPoint.ElementAt(i).Value
                    codeStartPoint = projectPoint.ElementAt(i).Key
                    Dim sourceEndPoint As ProjectionPoint = projectPoint.ElementAt(i).Value
                    codeEndPoint = projectPoint.ElementAt(i).Key
                    If i < projectPoint.Count - 1 Then
                        sourceEndPoint = projectPoint.ElementAt(i + 1).Value
                        codeEndPoint = projectPoint.ElementAt(i + 1).Key
                    Else
                        sourceEndPoint = projectPoint.First.Value
                        codeEndPoint = projectPoint.First.Key
                    End If
                    Dim startPoint As Vector3D = sourceStartPoint.projectPoint
                    Dim endPoint As Vector3D = sourceEndPoint.projectPoint
                    Dim x1 As Double = 0
                    Dim y1 As Double = 0
                    Dim x2 As Double = 0
                    Dim y2 As Double = 0
                    If prjView = ProjectionPoint.projectView.Top Or prjView = ProjectionPoint.projectView.Bottom Then
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Y * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Y * k + offsetY
                    ElseIf prjView = ProjectionPoint.projectView.Back Or prjView = ProjectionPoint.projectView.Front Then
                        x1 = startPoint.X * k + offsetX
                        y1 = startPoint.Z * k + offsetY
                        x2 = endPoint.X * k + offsetX
                        y2 = endPoint.Z * k + offsetY
                    ElseIf prjView = ProjectionPoint.projectView.Left Or prjView = ProjectionPoint.projectView.Right Then
                        x1 = startPoint.Y * k + offsetX
                        y1 = startPoint.Z * k + offsetY
                        x2 = endPoint.Y * k + offsetX
                        y2 = endPoint.Z * k + offsetY
                    End If
                    Dim startLine As New System.Drawing.Point(x1, y1)
                    Dim endLine As New System.Drawing.Point(x2, y2)
                    Using pen As New System.Drawing.Pen(colorPen, 1)
                        'If shLine = True Then
                        '    pen.DashPattern = New Single() {8, 2}
                        'End If
                        g.DrawLine(pen, startLine, endLine)
                    End Using
                    'подписываем значение
                    Using font As New System.Drawing.Font("Arial", 11)
                        Using format As New StringFormat()
                            format.Alignment = StringAlignment.Center  ' Горизонтальное выравнивание по центру
                            format.LineAlignment = StringAlignment.Center ' Вертикальное выравнивание по центру
                            If drawDimText = True Then
                                Using brush As New SolidBrush(System.Drawing.Color.Blue)
                                    Dim l As Double = Math.Sqrt((startPoint.X - endPoint.X) ^ 2 + (startPoint.Y - endPoint.Y) ^ 2)
                                    Dim deltaX As Double = ((x2 + x1) / 2)
                                    Dim deltaZ As Double = ((y2 + y1) / 2)
                                    If codeStartPoint.IndexOf("top") > -1 And codeEndPoint.IndexOf("top") > -1 Like "top" Then
                                        deltaZ -= 25
                                    ElseIf codeStartPoint.IndexOf("bottom") > -1 And codeEndPoint.IndexOf("bottom") > -1 Then
                                        deltaZ += 15
                                    ElseIf codeStartPoint.IndexOf("top") > -1 And codeEndPoint.IndexOf("bottom") > -1 Then
                                        deltaX += 25
                                    ElseIf codeEndPoint.IndexOf("bottom") > -1 And codeStartPoint.IndexOf("top") > -1 Then
                                        deltaX -= 25
                                    End If
                                    g.DrawString(Math.Round(Math.Abs(l), 3), font, brush, New PointF(deltaX, deltaZ), format)
                                End Using
                            End If
                            If drawElevationText = True Then
                                Using brush As New SolidBrush(System.Drawing.Color.Red)
                                    Dim posZ As Double = y2
                                    If codeEndPoint.IndexOf("left") > -1 Then
                                        posZ -= 15
                                    ElseIf codeEndPoint.IndexOf("right") > -1 Then
                                        posZ -= 15
                                    End If
                                    g.DrawString(Math.Round(132.1, 3), font, brush, New PointF(x2, posZ), format)
                                End Using
                            End If
                        End Using
                    End Using
                Next i
                If prjView = ProjectionPoint.projectView.Top Then
                    Dim sourceStartPoint As ProjectionPoint = Nothing
                    Dim sourceEndPoint As ProjectionPoint = Nothing
                    If projectPoint.ContainsKey("middlePt1") = True Then
                        sourceStartPoint = projectPoint.Item("middlePt1")
                    End If
                    If projectPoint.ContainsKey("middlePt2") = True Then
                        sourceEndPoint = projectPoint.Item("middlePt2")
                    End If
                    If IsNothing(sourceEndPoint) = False And IsNothing(sourceStartPoint) = False Then
                        Dim startPoint As Vector3D = sourceStartPoint.projectPoint
                        Dim endPoint As Vector3D = sourceEndPoint.projectPoint
                        Dim x1 As Double = 0
                        Dim y1 As Double = 0
                        Dim x2 As Double = 0
                        Dim y2 As Double = 0
                        If prjView = ProjectionPoint.projectView.Top Or prjView = ProjectionPoint.projectView.Bottom Then
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Y * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Y * k + offsetY
                        ElseIf prjView = ProjectionPoint.projectView.back Or prjView = ProjectionPoint.projectView.front Then
                            x1 = startPoint.X * k + offsetX
                            y1 = startPoint.Z * k + offsetY
                            x2 = endPoint.X * k + offsetX
                            y2 = endPoint.Z * k + offsetY
                        ElseIf prjView = ProjectionPoint.projectView.left Or prjView = ProjectionPoint.projectView.right Then
                            x1 = startPoint.Y * k + offsetX
                            y1 = startPoint.Z * k + offsetY
                            x2 = endPoint.Y * k + offsetX
                            y2 = endPoint.Z * k + offsetY
                        End If
                        Dim startLine As New System.Drawing.Point(x1, y1)
                        Dim endLine As New System.Drawing.Point(x2, y2)
                        Using pen As New System.Drawing.Pen(colorPen, 1)
                            g.DrawLine(pen, startLine, endLine)
                        End Using
                    End If
                End If
            End If
        End Using
    End Function
    'функция рисует землю в pictureBox
    Public Shared Function drawSurfaceLine(ByRef bmp As Bitmap, userSurface As Surface, ByVal axisLine As DwgLine, ByVal maxElevation As Double, ByVal k As Double, ByVal offsetX As Single, ByVal offsetY As String, Optional drawElevationText As Boolean = False, Optional egSurface As Boolean = False) As Boolean
        '=======================================================================================================
        'рисуем поверхность земли
        If IsNothing(userSurface) = False Then
            Dim listPoly As List(Of Vector2D) = New List(Of Vector2D)
            listPoly.Add(axisLine.StartPoint.Pos)
            listPoly.Add(axisLine.EndPoint.Pos)
            Dim sectEgSurf As Sfc.Sections.Section = userSurface.CreateSection(listPoly, Sfc.Sections.SectionFlags.Default)
            If sectEgSurf.Count > 1 Then
                Using g As Graphics = Graphics.FromImage(bmp)
                    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    Dim sect1 As Sfc.Sections.SectionNode = sectEgSurf.ElementAt(0)
                    Dim dist1 As Double = Math.Round(sect1.Vertex.X, 3)
                    Dim elev1 As Double = Math.Round(maxElevation - sect1.Vertex.Y, 3)
                    Dim dh1 As Double = offsetY + elev1 * k
                    Dim startEgLine As New System.Drawing.Point(dist1 * k + offsetX, dh1)
                    For i As Integer = 1 To sectEgSurf.Count - 1
                        Dim sect2 As Sfc.Sections.SectionNode = sectEgSurf.ElementAt(i)
                        Dim dist2 As Double = Math.Round(sect2.Vertex.X, 3)
                        Dim elev2 As Double = Math.Round(maxElevation - sect2.Vertex.Y, 3)
                        Dim dh2 As Double = offsetY + elev2 * k
                        Dim endEgLine As New System.Drawing.Point(dist2 * k + offsetX, dh2)
                        Dim colorPenEG As Color = Color.Red
                        If egSurface = True Then
                            colorPenEG = Color.Green
                        End If
                        Using pen As New Pen(colorPenEG, 1)
                            g.DrawLine(pen, startEgLine, endEgLine)
                        End Using
                        startEgLine = endEgLine
                    Next
                End Using
            End If
        End If
        Return True
    End Function
    'функция выбирает все солиды для показа на поперечнике
    Public Shared Function getModel3DElement(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberProlet As Integer) As List(Of DwgModel3DElement)
        Dim result As New List(Of DwgModel3DElement)
        If IsNothing(dictionaryObjectsBridge) = True Then Return result
        If numberPillar > 0 Then
            For i As Integer = 0 To dictionaryObjectsBridge.Count - 1
                Dim typeSelectObject As StructureElement.typeObject = dictionaryObjectsBridge.ElementAt(i).Key
                If typeSelectObject = StructureElement.typeObject.modelCabinetWall Or typeSelectObject = StructureElement.typeObject.modelCabinetWallPlate Or typeSelectObject = StructureElement.typeObject.modelNozzle Or typeSelectObject = StructureElement.typeObject.modelSubFermenters Or typeSelectObject = typeSelectObject.modelLeftHand Or typeSelectObject = typeSelectObject.modelLeftHandCornice Or typeSelectObject = typeSelectObject.modelRightHand Or typeSelectObject = typeSelectObject.modelRightHandCornice Or typeSelectObject = typeSelectObject.modelLeftPostcard Or typeSelectObject = typeSelectObject.modelRightPostcard Or typeSelectObject = typeSelectObject.modelRack Or typeSelectObject = typeSelectObject.modelGrillage Or typeSelectObject = typeSelectObject.modelPreparation Or typeSelectObject = typeSelectObject.modelPile Or typeSelectObject = typeSelectObject.modelRigel Then
                    Dim listDataElement As List(Of StructureElement) = dictionaryObjectsBridge.ElementAt(i).Value
                    For j As Integer = 0 To listDataElement.Count - 1
                        Dim dataElement As StructureElement = listDataElement.Item(j)
                        Dim tempNumberPillar As Integer = FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar")
                        If tempNumberPillar = numberPillar Then
                            Dim entity As DwgEntity = dataElement.DWGEntity
                            If IsNothing(entity) = False Then
                                If TypeOf entity Is DwgModel3DElement Then
                                    result.Add(entity)
                                End If
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
        If numberProlet > 0 Then
            For i As Integer = 0 To dictionaryObjectsBridge.Count - 1
                Dim typeSelectObject As StructureElement.typeObject = dictionaryObjectsBridge.ElementAt(i).Key
                If typeSelectObject = StructureElement.typeObject.modelBeam Or typeSelectObject = StructureElement.typeObject.modelSiteMonolitBeams Or typeSelectObject = StructureElement.typeObject.modelSiteMonolitPillar Then
                    Dim listDataElement As List(Of StructureElement) = dictionaryObjectsBridge.ElementAt(i).Value
                    For j As Integer = 0 To listDataElement.Count - 1
                        Dim dataElement As StructureElement = listDataElement.Item(j)
                        Dim tempNumberPillar As Integer = FuncGSON.getValue(dataElement.KeyParameter, "NumberPillar")
                        Dim tempNumberProlet As Integer = FuncGSON.getValue(dataElement.KeyParameter, "NumberProlet")
                        If tempNumberPillar = numberPillar Then
                            If tempNumberProlet = numberProlet Then
                                Dim entity As DwgEntity = dataElement.DWGEntity
                                If IsNothing(entity) = False Then
                                    If TypeOf entity Is DwgModel3DElement Then
                                        result.Add(entity)
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
        Return result
    End Function
End Class

