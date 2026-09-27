
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media.Animation
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Layers
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
    Private _minClearence As Double 'минимальный зазор
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
        _minClearence = 0
        _siteMonolit = 0
        _elevationLand = 0
        _countSubPillars = 1
        _singleSubFarmer = False
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
    <DisplayName("Левый зазор")>
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
    <Description("Минимальный зазор между балками смежных пролетов, м")>
    <Category("Свойства")>
    <DisplayName("Минимальный зазор")>
    Public Property MinClearence() As Double
        Get
            Return _minClearence
        End Get
        Set(value As Double)
            _minClearence = value
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
    Public Shared Function getAxisPillar(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional removeDictionary As Boolean = False) As StructureElement
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
                                    If removeDictionary = True Then
                                        listAxisPillar.RemoveAt(k)
                                    End If
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
    '=========================================================================================================
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
    '=========================================================================================================
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
    '=========================================================================================================
    'функция перемещает ось опоры вдоль трассы на указанную величину
    Public Shared Function moveAxisPillarToStation(ByVal projectAlignmrnt As Alignment, ByRef axisPillar As DwgLine, ByVal station As Double) As Boolean
        Dim result As Boolean = False
        If IsNothing(projectAlignmrnt) = True Then Return False
        If projectAlignmrnt.Plan.CompoundLine.Length = 0 Then Return False
        If IsNothing(axisPillar) = True Then Return False
        If axisPillar.Length = 0 Then Return False
        Dim centerPoint As Vector2D = New Vector2D
        Dim boolPosition As Boolean = projectAlignmrnt.Plan.CompoundLine.StaOffsetToPos(station, 0, centerPoint)
        If boolPosition = True Then
            Dim angleStart As Double = (axisPillar.StartPoint.Pos - centerPoint).Angle
            Dim distStart As Double = (axisPillar.StartPoint.Pos - centerPoint).Length
            Dim angleEnd As Double = (axisPillar.EndPoint.Pos - centerPoint).Angle
            Dim distEnd As Double = (axisPillar.EndPoint.Pos - centerPoint).Length
            'создаем произвольный вектор 
            Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerPoint, angleStart, distStart)
            Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(centerPoint, angleEnd, distEnd)
            axisPillar.StartPoint = New Vector3D(pt1, axisPillar.StartPoint.Z)
            axisPillar.EndPoint = New Vector3D(pt2, axisPillar.EndPoint.Z)
            result = True
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция вычисляет положение осей опор и осей опирания балок по словалю с балками (по крайним рядам)
    Public Shared Function calculateAxisBeamsPillar(ByRef beamsStructure As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByRef dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByVal projectAlignment As Alignment) As Boolean
        If IsNothing(beamsStructure) = True Then Return False
        If beamsStructure.Count = 0 Then Return False
        If IsNothing(dictAxisPillar) = False Then
            If dictAxisPillar.Count > 0 Then
                For i As Integer = 0 To dictAxisPillar.Count - 2
                    Dim numberPillar As Integer = dictAxisPillar.ElementAt(i).Key 'номер опоры
                    Dim leftAxisBeam As DwgLine = Nothing
                    Dim rightAxisBeam As DwgLine = Nothing
                    'первая опора
                    Dim listPillar1 As List(Of StructureElement) = dictAxisPillar.ElementAt(i).Value 'список осей
                    If IsNothing(listPillar1) = True Then Continue For
                    If listPillar1.Count < 3 Then Continue For
                    'ось первой опоры
                    Dim dataAxisPillar1 As StructureElement = listPillar1.Item(1)
                    If IsNothing(listPillar1) = True Then Continue For
                    If listPillar1.Count < 3 Then Continue For
                    Dim axisPillar1 As DwgLine = dataAxisPillar1.DWGEntity 'ось опоры
                    Dim userAxisPillar1 As Pillar = dataAxisPillar1.getPillar()
                    'ось опирания балок последующего пролета
                    Dim dataAxisBeamsPillar As StructureElement = Nothing
                    Dim userAxisBeamsPillar As AxisBeamsPillars = Nothing
                    Dim axisBeamsPillar As DwgLine = Nothing
                    If IsNothing(listPillar1.Item(2)) = False Then
                        dataAxisBeamsPillar = listPillar1.Item(2)
                        userAxisBeamsPillar = dataAxisBeamsPillar.getAxisBeamsPillar()
                        axisBeamsPillar = dataAxisBeamsPillar.DWGEntity 'ось опирания балок последующего пролета
                    End If
                    'вторая опора
                    Dim listPillar2 As List(Of StructureElement) = dictAxisPillar.ElementAt(i + 1).Value 'список осей
                    If IsNothing(listPillar2) = True Then Continue For
                    If listPillar2.Count < 3 Then Continue For
                    Dim dataAxisPillar2 As StructureElement = listPillar2.Item(1)
                    Dim userAxisPillar2 As Pillar = dataAxisPillar2.getPillar()
                    Dim axisPillar2 As DwgLine = dataAxisPillar2.DWGEntity
                    'ось опирания балок предыдущего пролета
                    Dim dataAxisPrevBeamsPillar As StructureElement = Nothing
                    Dim userAxisPrevBeamsPillar As AxisBeamsPillars = Nothing
                    Dim axisPrevBeamsPillar As DwgLine = Nothing
                    'забираем оси оси опирания балок
                    If IsNothing(listPillar2.Item(0)) = False Then
                        dataAxisPrevBeamsPillar = listPillar2.Item(0)
                        userAxisPrevBeamsPillar = dataAxisPrevBeamsPillar.getAxisBeamsPillar()
                        axisPrevBeamsPillar = listPillar2.Item(0).DWGEntity 'ось опирания балок последующего пролета
                    End If
                    If beamsStructure.ContainsKey(numberPillar) = True Then
                        'получаем список балок текущего пролета
                        Dim listBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar)
                        If IsNothing(listBeams) = True Then Continue For
                        If listBeams.Count = 0 Then Continue For
                        '=========================================================================================================================
                        'крайняя левая балка
                        Dim dataLeftBeam As StructureElement = listBeams.First.Value
                        Dim userLeftBeam As BeamI = dataLeftBeam.getBeamI()
                        leftAxisBeam = dataLeftBeam.DWGEntity
                        leftAxisBeam.StartPoint = userLeftBeam._elementBridgePoint.StartAxisPoint
                        leftAxisBeam.EndPoint = userLeftBeam._elementBridgePoint.EndAxisPoint
                        If leftAxisBeam.Length > 0 Then
                            'проверяе направление балки (по ходу пикетажа)
                            Dim boolCorrDirection As Boolean = AxisBeamsPillars.correctionAxisDirectionBeam(leftAxisBeam, projectAlignment)
                            If boolCorrDirection = True Then
                                userLeftBeam._elementBridgePoint.StartAxisPoint = leftAxisBeam.StartPoint
                                userLeftBeam._elementBridgePoint.EndAxisPoint = leftAxisBeam.EndPoint
                            End If
                        Else
                            MsgBox("Не удалось вычислить длину балки.")
                            Continue For
                        End If
                        '=========================================================================================================================
                        'крайняя правая балка
                        If listBeams.Count > 1 Then
                            Dim dataRightBeam As StructureElement = listBeams.Last.Value
                            Dim userRightBeam As BeamI = dataRightBeam.getBeamI()
                            rightAxisBeam = dataRightBeam.DWGEntity
                            rightAxisBeam.StartPoint = userRightBeam._elementBridgePoint.StartAxisPoint
                            rightAxisBeam.EndPoint = userRightBeam._elementBridgePoint.EndAxisPoint
                            If rightAxisBeam.Length > 0 Then
                                Dim boolCorrDirection As Boolean = AxisBeamsPillars.correctionAxisDirectionBeam(rightAxisBeam, projectAlignment)
                                If boolCorrDirection = True Then
                                    userRightBeam._elementBridgePoint.StartAxisPoint = rightAxisBeam.StartPoint
                                    userRightBeam._elementBridgePoint.EndAxisPoint = rightAxisBeam.EndPoint
                                End If
                            End If
                            'в пролете более 1 ряда
                            If userRightBeam.numberRow <> userLeftBeam.numberRow Then
                                'если первая опора или последняя опора (ось опоры - это ось опирания балок)
                                'ось строим по концам балок
                                If i = 0 Then
                                    'ось строим по концам балок
                                    axisPillar1.StartPoint = leftAxisBeam.StartPoint
                                    axisPillar1.EndPoint = rightAxisBeam.StartPoint
                                    userAxisPillar1._elementBridgePoint.StartAxisPoint = leftAxisBeam.StartPoint
                                    userAxisPillar1._elementBridgePoint.EndAxisPoint = rightAxisBeam.StartPoint
                                    dataAxisPillar1.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar1)
                                    If dictAxisPillar.Count = 2 Then
                                        axisPillar2.StartPoint = leftAxisBeam.EndPoint
                                        axisPillar2.EndPoint = rightAxisBeam.EndPoint
                                        userAxisPillar2._elementBridgePoint.StartAxisPoint = leftAxisBeam.EndPoint
                                        userAxisPillar2._elementBridgePoint.EndAxisPoint = rightAxisBeam.EndPoint
                                        dataAxisPillar2.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar2)
                                    Else
                                        axisPrevBeamsPillar.StartPoint = leftAxisBeam.EndPoint
                                        axisPrevBeamsPillar.EndPoint = rightAxisBeam.EndPoint
                                        userAxisPrevBeamsPillar._elementBridgePoint.StartAxisPoint = axisPrevBeamsPillar.StartPoint
                                        userAxisPrevBeamsPillar._elementBridgePoint.EndAxisPoint = axisPrevBeamsPillar.EndPoint
                                        dataAxisPrevBeamsPillar.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPrevBeamsPillar)
                                    End If
                                ElseIf i = dictAxisPillar.Count - 2 Then
                                    'ось строим по концам балок
                                    If IsNothing(axisBeamsPillar) = False Then
                                        axisBeamsPillar.StartPoint = leftAxisBeam.StartPoint
                                        axisBeamsPillar.EndPoint = rightAxisBeam.StartPoint
                                        userAxisBeamsPillar._elementBridgePoint.StartAxisPoint = leftAxisBeam.EndPoint
                                        userAxisBeamsPillar._elementBridgePoint.EndAxisPoint = rightAxisBeam.EndPoint
                                        dataAxisBeamsPillar.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisBeamsPillar)
                                    End If
                                    axisPillar2.StartPoint = leftAxisBeam.EndPoint
                                    axisPillar2.EndPoint = rightAxisBeam.EndPoint
                                    userAxisPillar2._elementBridgePoint.StartAxisPoint = leftAxisBeam.EndPoint
                                    userAxisPillar2._elementBridgePoint.EndAxisPoint = rightAxisBeam.EndPoint
                                    dataAxisPillar2.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar2)
                                Else
                                    'опора промежуточная
                                    axisPrevBeamsPillar.StartPoint = leftAxisBeam.EndPoint
                                    axisPrevBeamsPillar.EndPoint = rightAxisBeam.EndPoint
                                    userAxisPrevBeamsPillar._elementBridgePoint.StartAxisPoint = leftAxisBeam.EndPoint
                                    userAxisPrevBeamsPillar._elementBridgePoint.EndAxisPoint = rightAxisBeam.EndPoint
                                    dataAxisPrevBeamsPillar.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPrevBeamsPillar)
                                    axisBeamsPillar.StartPoint = leftAxisBeam.StartPoint
                                    axisBeamsPillar.EndPoint = rightAxisBeam.StartPoint
                                    userAxisBeamsPillar._elementBridgePoint.StartAxisPoint = leftAxisBeam.StartPoint
                                    userAxisBeamsPillar._elementBridgePoint.EndAxisPoint = rightAxisBeam.StartPoint
                                    dataAxisBeamsPillar.KeyParameter = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisBeamsPillar)
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return True
    End Function
    '=========================================================================================================
    'функция вычисляет или корректирует положение осей опор и осей опирания балок по словалю с балками
    Public Shared Function calculateAxisPillarWithCorrctBeams(ByRef beamsStructure As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement))) As Boolean
        Dim result As Boolean = True
        If IsNothing(dictAxisPillar) = False Then
            If dictAxisPillar.Count > 2 Then
                'проверяем только промежуточные пролеты
                For i As Integer = 1 To dictAxisPillar.Count - 2
                    Dim numberPillar As Integer = dictAxisPillar.ElementAt(i).Key 'номер опоры
                    Dim listPillar As List(Of StructureElement) = dictAxisPillar.ElementAt(i).Value 'список осей
                    If IsNothing(listPillar) = True Then Continue For
                    If listPillar.Count < 3 Then Continue For
                    Dim dataPillar As StructureElement = listPillar(1)
                    If IsNothing(dataPillar) = True Then Continue For
                    Dim userPillar As Pillar = dataPillar.getPillar
                    If IsNothing(userPillar) = True Then Continue For
                    '==========================================================================================
                    'смотрим зазор
                    Dim leftZazor As Double = userPillar.Clearence
                    Dim rightZazor As Double = userPillar.RightClearence
                    'пкремещение оси опирания
                    Dim moveAxisBeamsPillar As Double = 0
                    '==========================================================================================
                    If beamsStructure.ContainsKey(numberPillar) = True Then
                        'получаем список балок текущего пролета
                        Dim listPrevBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar - 1)
                        Dim listBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar)
                        If IsNothing(listBeams) = True Then Continue For
                        If listBeams.Count > 2 Then
                            'проверяем только промежуточных балок
                            For j As Integer = 1 To listBeams.Count - 2
                                Dim dataBeamI As StructureElement = listBeams.ElementAt(j).Value
                                If IsNothing(dataBeamI) = True Then Continue For
                                Dim userBeamI As BeamI = dataBeamI.getBeamI()
                                If IsNothing(userBeamI) = True Then Continue For
                                Dim numberRow As Integer = userBeamI.numberRow
                                Dim prevUserBeam As BeamI = Nothing
                                For k As Integer = 1 To listPrevBeams.Count - 2
                                    Dim prevDataBeamI As StructureElement = listPrevBeams.ElementAt(k).Value
                                    If IsNothing(prevDataBeamI) = True Then Continue For
                                    Dim tempPevUserBeam As BeamI = prevDataBeamI.getBeamI()
                                    If IsNothing(tempPevUserBeam) = True Then Continue For
                                    If tempPevUserBeam.numberRow = numberRow Then
                                        prevUserBeam = tempPevUserBeam
                                        Exit For
                                    End If
                                Next k
                                If IsNothing(userBeamI) = False And IsNothing(prevUserBeam) = False Then
                                    Dim userListClearence As List(Of Double) = New List(Of Double)
                                    Dim clearence As Double = CalculationBeams.getClearenceBeams(prevUserBeam, userBeamI, userListClearence)
                                    If clearence < userPillar.MinClearence Then
                                        Dim deltaMove As Double = userPillar.MinClearence - clearence
                                        If moveAxisBeamsPillar < deltaMove Then
                                            moveAxisBeamsPillar = Math.Round(deltaMove, 3)
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    End If
                    If moveAxisBeamsPillar > 0.0006 Then
                        If leftZazor < rightZazor Then
                            userPillar.Clearence += moveAxisBeamsPillar
                        ElseIf leftZazor > rightZazor Then
                            userPillar.RightClearence += moveAxisBeamsPillar
                        Else
                            userPillar.Clearence += moveAxisBeamsPillar
                            userPillar.RightClearence += moveAxisBeamsPillar
                        End If
                        Dim strGSONBeamI As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
                        dataPillar.KeyParameter = strGSONBeamI
                        result = False
                    End If
                Next i
            End If
        End If
        Return result
    End Function
    '=========================================================================================================
    'функция вычисляет положение осей опор
    Public Shared Function calculateAxisPillar(ByVal userBridge As Bridges, ByRef beamsStructure As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)), ByVal dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByVal projectAlignment As Alignment) As Boolean
        calculateAxisPillar = False
        If IsNothing(userBridge) = True Then Return False
        If IsNothing(beamsStructure) = True Then Return False
        If beamsStructure.Count = 0 Then Return False
        'граница габарита моста слева
        Dim axisPlineDirect As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, -1 * userBridge.LeftStructureWidth + userBridge.TransverseOffset)
        Dim axisPline3DDirect = New Polyline3D()
        axisPlineDirect.GetPolyline(axisPline3DDirect)
        'граница габарита моста справа
        Dim axisPlineReverse As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(projectAlignment, userBridge.RightStructureWidth + userBridge.TransverseOffset)
        Dim axisPline3DReverse = New Polyline3D()
        axisPlineReverse.GetPolyline(axisPline3DReverse)
        For i As Integer = 0 To dictAxisPillar.Count - 1
            Dim numberPillar As Integer = dictAxisPillar.ElementAt(i).Key 'номер опоры
            Dim listPillar As List(Of StructureElement) = dictAxisPillar.ElementAt(i).Value 'список осей
            If IsNothing(listPillar) = True Then Continue For
            If listPillar.Count < 3 Then Continue For
            Dim dataPillar As StructureElement = listPillar(1)
            If IsNothing(dataPillar) = True Then Continue For
            Dim userPillar As Pillar = dataPillar.getPillar
            If IsNothing(userPillar) = True Then Continue For
            Dim axisPillar As DwgLine = dataPillar.DWGEntity
            '==========================================================================================
            If beamsStructure.ContainsKey(numberPillar) = True Then
                Dim leftAxisPoint As Vector3D = New Vector3D()
                Dim rightAxisPoint As Vector3D = New Vector3D()
                If i = 0 Then
                    'получаем список балок следующего пролета
                    Dim listBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar)
                    If IsNothing(listBeams) = True Then Continue For
                    Dim dataLeftBeam As StructureElement = listBeams.First.Value
                    If IsNothing(dataLeftBeam) = True Then Continue For
                    Dim userLeftBeam As BeamI = dataLeftBeam.getBeamI
                    If IsNothing(userLeftBeam) = True Then Continue For
                    Dim lineLeftBeam As DwgLine = dataLeftBeam.DWGEntity
                    If IsNothing(lineLeftBeam) = True Then Continue For
                    leftAxisPoint = userLeftBeam._elementBridgePoint.StartAxisPoint
                    'правая балка
                    Dim dataRightBeam As StructureElement = listBeams.Last.Value
                    If IsNothing(dataRightBeam) = True Then Continue For
                    Dim userRightBeam As BeamI = dataRightBeam.getBeamI
                    If IsNothing(userRightBeam) = True Then Continue For
                    Dim lineRightBeam As DwgLine = dataRightBeam.DWGEntity
                    If IsNothing(lineRightBeam) = True Then Continue For
                    rightAxisPoint = userRightBeam._elementBridgePoint.StartAxisPoint
                    If (leftAxisPoint - rightAxisPoint).Length = 0 Then
                        'один центральный ряд балок
                        Dim angle As Double = lineLeftBeam.Rotation + Math.PI / 2
                        If angle > Math.PI * 2 Then
                            angle -= Math.PI
                        End If
                        Dim tempLeftAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftAxisPoint, angle, userBridge.LeftStructureWidth + 2)
                        angle = lineLeftBeam.Rotation - Math.PI / 2
                        If angle < 0 Then
                            angle += Math.PI * 2
                        End If
                        Dim tempRightAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightAxisPoint, angle, userBridge.RightStructureWidth + 2)
                        userPillar._elementBridgePoint.StartAxisPoint = New Vector3D(tempLeftAxisPoint, leftAxisPoint.Z)
                        userPillar._elementBridgePoint.EndAxisPoint = New Vector3D(tempRightAxisPoint, leftAxisPoint.Z)
                    Else
                        userPillar._elementBridgePoint.StartAxisPoint = leftAxisPoint
                        userPillar._elementBridgePoint.EndAxisPoint = rightAxisPoint
                    End If
                    axisPillar.StartPoint = userPillar._elementBridgePoint.StartAxisPoint
                    axisPillar.EndPoint = userPillar._elementBridgePoint.EndAxisPoint
                    Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
                    dataPillar.KeyParameter = keyParam
                ElseIf i = dictAxisPillar.Count - 1 Then
                    'получаем список балок следующего пролета
                    Dim listBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar)
                    If IsNothing(listBeams) = True Then Continue For
                    Dim dataLeftBeam As StructureElement = listBeams.First.Value
                    If IsNothing(dataLeftBeam) = True Then Continue For
                    Dim userLeftBeam As BeamI = dataLeftBeam.getBeamI
                    If IsNothing(userLeftBeam) = True Then Continue For
                    Dim lineLeftBeam As DwgLine = dataLeftBeam.DWGEntity
                    If IsNothing(lineLeftBeam) = True Then Continue For
                    leftAxisPoint = userLeftBeam._elementBridgePoint.EndAxisPoint
                    'правая балка
                    Dim dataRightBeam As StructureElement = listBeams.Last.Value
                    If IsNothing(dataRightBeam) = True Then Continue For
                    Dim userRightBeam As BeamI = dataRightBeam.getBeamI
                    If IsNothing(userRightBeam) = True Then Continue For
                    Dim lineRightBeam As DwgLine = dataRightBeam.DWGEntity
                    If IsNothing(lineRightBeam) = True Then Continue For
                    rightAxisPoint = userRightBeam._elementBridgePoint.EndAxisPoint
                    If (leftAxisPoint - rightAxisPoint).Length = 0 Then
                        'один центральный ряд балок
                        Dim angle As Double = lineLeftBeam.Rotation + Math.PI / 2
                        If angle > Math.PI * 2 Then
                            angle -= Math.PI
                        End If
                        Dim tempLeftAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(leftAxisPoint, angle, userBridge.LeftStructureWidth + 2)
                        angle = lineLeftBeam.Rotation - Math.PI / 2
                        If angle < 0 Then
                            angle += Math.PI * 2
                        End If
                        Dim tempRightAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(rightAxisPoint, angle, userBridge.RightStructureWidth + 2)
                        userPillar._elementBridgePoint.StartAxisPoint = New Vector3D(tempLeftAxisPoint, leftAxisPoint.Z)
                        userPillar._elementBridgePoint.EndAxisPoint = New Vector3D(tempRightAxisPoint, leftAxisPoint.Z)
                    Else
                        userPillar._elementBridgePoint.StartAxisPoint = leftAxisPoint
                        userPillar._elementBridgePoint.EndAxisPoint = rightAxisPoint
                    End If
                    axisPillar.StartPoint = userPillar._elementBridgePoint.StartAxisPoint
                    axisPillar.EndPoint = userPillar._elementBridgePoint.EndAxisPoint
                    Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
                    dataPillar.KeyParameter = keyParam
                Else
                    'получаем список балок предыдущего пролета
                    Dim listPrevBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar - 1)
                    If IsNothing(listPrevBeams) = True Then Continue For
                    'получаем список балок следующего пролета
                    Dim listBeams As Dictionary(Of Integer, StructureElement) = beamsStructure.Item(numberPillar)
                    If IsNothing(listBeams) = True Then Continue For
                    'получаем ось левой балки предыдущего пролета
                    Dim dataPrevLeftBeam As StructureElement = listPrevBeams.First.Value
                    If IsNothing(dataPrevLeftBeam) = True Then Continue For
                    Dim userPrevLeftBeam As BeamI = dataPrevLeftBeam.getBeamI
                    If IsNothing(userPrevLeftBeam) = True Then Continue For
                    Dim prevLeftAxisBeam As DwgLine = dataPrevLeftBeam.DWGEntity
                    If IsNothing(prevLeftAxisBeam) = True Then
                        prevLeftAxisBeam = New DwgLine
                        prevLeftAxisBeam.StartPoint = userPrevLeftBeam._elementBridgePoint.StartAxisPoint
                        prevLeftAxisBeam.EndPoint = userPrevLeftBeam._elementBridgePoint.EndAxisPoint
                    ElseIf prevLeftAxisBeam.Length = 0 Then
                        prevLeftAxisBeam.StartPoint = userPrevLeftBeam._elementBridgePoint.StartAxisPoint
                        prevLeftAxisBeam.EndPoint = userPrevLeftBeam._elementBridgePoint.EndAxisPoint
                    End If
                    If prevLeftAxisBeam.Length = 0 Then Continue For
                    Dim boolExtendPrevLeftLine As Boolean = BridgeGeometry.extendLine(prevLeftAxisBeam, 0, userPrevLeftBeam.b)
                    'получаем ось левой балки следующего пролета
                    Dim dataLeftBeam As StructureElement = listBeams.First.Value
                    If IsNothing(dataLeftBeam) = True Then Continue For
                    Dim userLeftBeam As BeamI = dataLeftBeam.getBeamI
                    If IsNothing(userLeftBeam) = True Then Continue For
                    Dim leftAxisBeam As DwgLine = dataLeftBeam.DWGEntity
                    If IsNothing(leftAxisBeam) = True Then
                        leftAxisBeam = New DwgLine
                        leftAxisBeam.StartPoint = userLeftBeam._elementBridgePoint.StartAxisPoint
                        leftAxisBeam.EndPoint = userLeftBeam._elementBridgePoint.EndAxisPoint
                        leftAxisBeam.StartPoint = userLeftBeam._elementBridgePoint.StartAxisPoint
                    ElseIf leftAxisBeam.Length = 0 Then
                        leftAxisBeam.StartPoint = userLeftBeam._elementBridgePoint.StartAxisPoint
                        leftAxisBeam.EndPoint = userLeftBeam._elementBridgePoint.EndAxisPoint
                    End If
                    If leftAxisBeam.Length = 0 Then Continue For
                    Dim boolExtendLeftLine As Boolean = BridgeGeometry.extendLine(leftAxisBeam, userLeftBeam.a, 0)
                    '=========================================================================================
                    'получаем ось правой балки предыдущего пролета
                    Dim dataPrevRightBeam As StructureElement = listPrevBeams.Last.Value
                    If IsNothing(dataPrevRightBeam) = True Then Continue For
                    Dim userPrevRightBeam As BeamI = dataPrevRightBeam.getBeamI
                    If IsNothing(userPrevRightBeam) = True Then Continue For
                    Dim prevRightAxisBeam As DwgLine = dataPrevRightBeam.DWGEntity
                    If IsNothing(prevRightAxisBeam) = True Then
                        prevRightAxisBeam = New DwgLine
                        prevRightAxisBeam.StartPoint = userPrevRightBeam._elementBridgePoint.StartAxisPoint
                        prevRightAxisBeam.EndPoint = userPrevRightBeam._elementBridgePoint.EndAxisPoint
                    ElseIf prevRightAxisBeam.Length = 0 Then
                        prevRightAxisBeam.StartPoint = userPrevRightBeam._elementBridgePoint.StartAxisPoint
                        prevRightAxisBeam.EndPoint = userPrevRightBeam._elementBridgePoint.EndAxisPoint
                    End If
                    If prevRightAxisBeam.Length = 0 Then Continue For
                    Dim boolExtendPrevRightLine As Boolean = BridgeGeometry.extendLine(prevRightAxisBeam, 0, userPrevRightBeam.b)
                    'получаем ось правой балки следующего пролета
                    Dim dataRightBeam As StructureElement = listBeams.Last.Value
                    If IsNothing(dataRightBeam) = True Then Continue For
                    Dim userRightBeam As BeamI = dataRightBeam.getBeamI
                    If IsNothing(userRightBeam) = True Then Continue For
                    Dim rightAxisBeam As DwgLine = dataRightBeam.DWGEntity
                    If IsNothing(rightAxisBeam) = True Then
                        rightAxisBeam = New DwgLine
                        rightAxisBeam.StartPoint = userRightBeam._elementBridgePoint.StartAxisPoint
                        rightAxisBeam.EndPoint = userRightBeam._elementBridgePoint.EndAxisPoint
                    ElseIf rightAxisBeam.Length = 0 Then
                        rightAxisBeam.StartPoint = userRightBeam._elementBridgePoint.StartAxisPoint
                        rightAxisBeam.EndPoint = userRightBeam._elementBridgePoint.EndAxisPoint
                    End If
                    Dim boolExtendRightLine As Boolean = BridgeGeometry.extendLine(rightAxisBeam, userRightBeam.a, 0)
                    'строим ось опоры
                    'находим среднюю точку слева
                    Dim middlePointLeft As Vector3D = CalculationBeams.calculateMiddlePointBeams(dataPrevLeftBeam, dataLeftBeam)
                    'находим среднюю точку справа
                    Dim middlePointRight As Vector3D = CalculationBeams.calculateMiddlePointBeams(dataPrevRightBeam, dataRightBeam)
                    If (middlePointLeft - middlePointRight).Length = 0 Then
                        'один центральный ряд балок
                        Dim angle As Double = (prevLeftAxisBeam.Rotation + leftAxisBeam.Rotation) / 2 + Math.PI / 2
                        If angle > Math.PI * 2 Then
                            angle -= Math.PI
                        End If
                        Dim tempLeftAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middlePointLeft, angle, userBridge.LeftStructureWidth + 2)
                        angle = (prevLeftAxisBeam.Rotation + leftAxisBeam.Rotation) / 2 - Math.PI / 2
                        If angle < 0 Then
                            angle += Math.PI * 2
                        End If
                        Dim tempRightAxisPoint As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(middlePointLeft, angle, userBridge.RightStructureWidth + 2)
                        userPillar._elementBridgePoint.StartAxisPoint = New Vector3D(tempLeftAxisPoint, middlePointLeft.Z)
                        userPillar._elementBridgePoint.EndAxisPoint = New Vector3D(tempRightAxisPoint, middlePointLeft.Z)
                    Else
                        userPillar._elementBridgePoint.StartAxisPoint = middlePointLeft
                        userPillar._elementBridgePoint.EndAxisPoint = middlePointRight
                    End If
                    axisPillar.StartPoint = userPillar._elementBridgePoint.StartAxisPoint
                    axisPillar.EndPoint = userPillar._elementBridgePoint.EndAxisPoint
                    Dim boolExtendAxis As Boolean = BridgeGeometry.extendLine(axisPillar, 1000, 1000)
                    'находим пересечение оси опоры с левым габаритом
                    Dim pointLeftIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DDirect, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                    Dim pointRightIntersectCollection As IEnumerable(Of Vector2D) = PolylineExtentions.GetIntersections(axisPline3DReverse, axisPillar.StartPoint.Pos, axisPillar.EndPoint.Pos)
                    If pointLeftIntersectCollection.Count > 0 And pointRightIntersectCollection.Count > 0 Then
                        axisPillar.StartPoint = New Vector3D(pointLeftIntersectCollection(0), middlePointLeft.Z)
                        axisPillar.EndPoint = New Vector3D(pointRightIntersectCollection(0), middlePointRight.Z)
                        Dim boolExtendAxisPillar As Boolean = BridgeGeometry.extendLine(axisPillar, 2, 2)
                    End If
                    userPillar._elementBridgePoint.StartAxisPoint = axisPillar.StartPoint
                    userPillar._elementBridgePoint.EndAxisPoint = axisPillar.EndPoint
                    axisPillar.StartPoint = userPillar._elementBridgePoint.StartAxisPoint
                    axisPillar.EndPoint = userPillar._elementBridgePoint.EndAxisPoint
                    Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
                    dataPillar.KeyParameter = keyParam
                    Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(axisPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
                End If
            End If
        Next i
        Return True
    End Function
    'функция рисует оси опор
    Public Shared Function drawAxisPillar(ByRef drawingDocument As Dwg.Drawing, ByRef axisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleAxisPillar As ProjectCivilStructuresStyle = Nothing) As Boolean
        If IsNothing(styleAxisPillar) = True Then
            styleAxisPillar = New ProjectCivilStructuresStyle(drawingDocument)
        End If
        If IsNothing(axisPillar) = False Then
            If axisPillar.Count > 0 Then
                For i As Integer = 0 To axisPillar.Count - 1
                    Dim listAxisPillar As List(Of StructureElement) = axisPillar.ElementAt(i).Value
                    Dim dataAxisPillar As StructureElement = listAxisPillar.Item(1)
                    If IsNothing(dataAxisPillar) = False Then
                        Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                        Dim acLineAxisPillar As DwgLine = Nothing
                        Dim oldDataAxisPillar As StructureElement = Pillar.getAxisPillar(dictionaryObjectsBridge, userAxisPillar.Number, True)
                        If IsNothing(oldDataAxisPillar) = False Then
                            acLineAxisPillar = oldDataAxisPillar.DWGEntity
                            acLineAxisPillar.StartPoint = userAxisPillar._elementBridgePoint.StartAxisPoint
                            acLineAxisPillar.EndPoint = userAxisPillar._elementBridgePoint.EndAxisPoint
                            Dim oldUserAxisPillar As Pillar = oldDataAxisPillar.getPillar
                            userAxisPillar.ElevationLand = oldUserAxisPillar.ElevationLand
                            userAxisPillar.CountSubPillars = oldUserAxisPillar.CountSubPillars
                            userAxisPillar.SingleSubFarmer = oldUserAxisPillar.SingleSubFarmer
                            userAxisPillar.FixedHeightCabinetWall = oldUserAxisPillar.FixedHeightCabinetWall
                            userAxisPillar.PresenceRacks = oldUserAxisPillar.PresenceRacks
                            userAxisPillar.PresencGrillage = oldUserAxisPillar.PresencGrillage
                            userAxisPillar.PresencPreparation = oldUserAxisPillar.PresencPreparation
                            userAxisPillar.PileInRack = oldUserAxisPillar.PileInRack
                            userAxisPillar.HorizontalLevel = oldUserAxisPillar.HorizontalLevel
                        Else
                            acLineAxisPillar = dataAxisPillar.DWGEntity
                        End If
                        If IsNothing(acLineAxisPillar) = True Then
                            acLineAxisPillar = New DwgLine
                            dataAxisPillar.DWGEntity = acLineAxisPillar
                        End If
                        If acLineAxisPillar.Length = 0 Then
                            acLineAxisPillar.StartPoint = userAxisPillar._elementBridgePoint.StartAxisPoint
                            acLineAxisPillar.EndPoint = userAxisPillar._elementBridgePoint.EndAxisPoint
                        End If
                        If drawingDocument.ActiveSpace.Entities.Contains(acLineAxisPillar) = False Then
                            drawingDocument.ActiveSpace.Entities.Add(acLineAxisPillar)
                            Dim boolSetStyleBeam As Boolean = styleAxisPillar.setObjectStyle(acLineAxisPillar)
                        End If
                        Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisPillar)
                        dataAxisPillar.KeyParameter = keyParam
                        Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(acLineAxisPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataAxisPillar)
                    End If
                Next i
            End If
        End If
        'удаляем лишние элементы
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillar) = True Then
            Dim listAxisBeams As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillar)
            If IsNothing(listAxisBeams) = False Then
                For Each dataBeam As StructureElement In listAxisBeams
                    If IsNothing(dataBeam) = False Then
                        Dim axisLine As DwgEntity = dataBeam.DWGEntity
                        If IsNothing(axisLine) = False Then
                            If drawingDocument.ActiveSpace.Entities.Contains(axisLine) = True Then
                                drawingDocument.ActiveSpace.Entities.Remove(axisLine)
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Return True
    End Function
    'функция возвращает номер определяющей опоры
    Public Shared Function getDefinitNumberPillar(ByRef axisPillar As Dictionary(Of Integer, List(Of StructureElement))) As Integer
        Dim result As Integer = 0
        If IsNothing(axisPillar) = False Then
            If axisPillar.Count > 0 Then
                For i As Integer = 0 To axisPillar.Count - 1
                    Dim listAxisPillar As List(Of StructureElement) = axisPillar.ElementAt(i).Value
                    If IsNothing(listAxisPillar) = False Then
                        If listAxisPillar.Count > 1 Then
                            Dim dataAxisPillar As StructureElement = listAxisPillar.Item(1)
                            If IsNothing(dataAxisPillar) = False Then
                                Dim userAxisPillar As Pillar = dataAxisPillar.getPillar()
                                If userAxisPillar.Defining = True Then
                                    result = userAxisPillar.Number
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return result
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
                            ElseIf typeElement = StructureElement.typeObject.axisRightPostcard Then
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
                        Dim dataBeam As StructureElement = listBeams.ElementAt(i)
                        If IsNothing(dataBeam) = False Then
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
                ElseIf prjView = ProjectionPoint.projectView.Bottom Then
                    result = transformBottom
                ElseIf prjView = ProjectionPoint.projectView.Front Then
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
    'функция рисует контур в pictureBox
    Public Shared Function drawCounterPictureBox(ByRef bmp As Bitmap, ByVal projectPoint As Dictionary(Of String, ProjectionPoint), ByVal k As Double, ByVal offsetX As Single, ByVal offsetY As String, ByVal prjView As ProjectionPoint.projectView, Optional drawDimText As Boolean = False, Optional drawElevationText As Boolean = False) As Boolean
        Dim colorPen As System.Drawing.Color = System.Drawing.Color.Black
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            If projectPoint.Count > 2 Then
                Dim codeStartPoint As String = ""
                Dim codeEndPoint As String = ""
                For i As Integer = 0 To projectPoint.Count - 1
                    Try

                    Catch ex As Exception
                    Catch ex As Exception
                    End Try
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

