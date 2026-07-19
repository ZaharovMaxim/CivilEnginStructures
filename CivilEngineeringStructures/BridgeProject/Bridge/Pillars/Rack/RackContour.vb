Imports System.ComponentModel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports CivilEnginStructures.PileContour
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class RackContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberRack As Integer 'номер стойки
    Private _type As StructureElement.typeObject
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberRack = 0
        _type = StructureElement.typeObject.OtherElement
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberRack As Integer, TypeContour As StructureElement.typeObject)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberRack = NumberRack
        _type = TypeContour
    End Sub
    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property NumberPillar() As Integer
        Get
            Return _numberPillar
        End Get
        Set(value As Integer)
            _numberPillar = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер элемента")>
    <Category("Свойства")>
    <DisplayName("Номер элемента")>
    <[ReadOnly](True)>
    Public Property NumberSubPillar() As Integer
        Get
            Return _numberSubPillar
        End Get
        Set(value As Integer)
            _numberSubPillar = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер стойки")>
    <Category("Свойства")>
    <DisplayName("Номер стойки")>
    <[ReadOnly](True)>
    Public Property NumberRack() As Integer
        Get
            Return _numberRack
        End Get
        Set(value As Integer)
            _numberRack = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Тип объекта")>
    <Category("Свойства")>
    <DisplayName("Тип объекта")>
    <[ReadOnly](True)>
    Public Property TypeContour() As StructureElement.typeObject
        Get
            Return _type
        End Get
        Set(value As StructureElement.typeObject)
            _type = value
        End Set
    End Property

    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject, ByVal typeCounter As RackPillar.TypeRack) As StructureElement
        Dim elementPillar As StructureElement = New StructureElement()
        elementPillar.Label = "Мосты и путепроводы"
        elementPillar.ClassObject = StructureElement.classStructure.RackPillar
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterRackBottom Then
            elementPillar.Description = "Контур стойки по низу"
        ElseIf type = StructureElement.typeObject.counterRackTop Then
            elementPillar.Description = "Контур стойки по верху"
        End If
        elementPillar.KeyParameter = ""
        elementPillar.IdElement = Guid.NewGuid.ToString
        elementPillar.IdStructure = idBridge
        elementPillar.Note = ""
        If typeCounter = RackPillar.TypeRack.Circle Then
            elementPillar.DWGEntity = New DwgCircle
        Else
            elementPillar.DWGEntity = New DwgPolyline3D
        End If

        Return elementPillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий контур 
    Public Shared Function getContour(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberRack As Integer, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataContour As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(typeContour) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(typeContour)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As RackContour = tempData.getContourRackPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If numberRack = userObject.NumberRack Then
                                            dataContour = tempData
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataContour
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование контура 
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userRack As RackPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgEntity)
        Dim result As New Dictionary(Of StructureElement.typeObject, DwgEntity)
        If IsNothing(userRack) Then Return result
        'вспомогательные построения
        Dim layerRack As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorRack As CadColor = New CadColor(7)
        Dim nameTypeLineRack As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineRack As Integer = 1
        Dim widthTypeLineRack As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Стойка (верх контура)")
        'ещчки по верху и по низу
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userRack._elementBridgePoint.ListPointModel.Count > 3 Then
            For j As Integer = 0 To userRack._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userRack._elementBridgePoint.ListPointModel.ElementAt(j).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next j
        End If
        Dim topCenterPoint As Vector3D = userRack._elementBridgePoint.CenterTopPoint
        Dim bottomCenterPoint As Vector3D = userRack._elementBridgePoint.CenterBottomPoint
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = RackContour.getContour(dictionaryObjectsBridge, userRack.NumberPillar, userRack.Number, StructureElement.typeObject.counterRackTop, userRack.NumberSubPillars)
        Dim poly3dCounterTop As DwgPolyline3D = New DwgPolyline3D
        Dim circle As DwgCircle = New DwgCircle
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = RackContour.createContour(idBridge, StructureElement.typeObject.counterRackTop, userRack.RackType)
            If IsNothing(dataTopCounter.DWGEntity) = False Then
                If TypeOf dataTopCounter.DWGEntity Is DwgCircle Then
                    circle = dataTopCounter.DWGEntity
                ElseIf TypeOf dataTopCounter.DWGEntity Is DwgPolyline3D Then
                    poly3dCounterTop = dataTopCounter.DWGEntity
                End If
            End If
        Else
            If TypeOf dataTopCounter.DWGEntity Is DwgPolyline3D Then
                poly3dCounterTop = dataTopCounter.DWGEntity
            ElseIf TypeOf dataTopCounter.DWGEntity Is DwgCircle Then
                circle = dataTopCounter.DWGEntity
            End If
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        Dim arrayCounter As DwgPolyline3D() = {Nothing, Nothing}
        'рисуем сваю
        If userRack.RackType = RackPillar.TypeRack.Circle Then
            'рисуем окружность
            If IsNothing(poly3dCounterTop) = False Then
                activProjectDocument.ActiveSpace.Entities.Remove(poly3dCounterTop)
            End If
            If IsNothing(circle) = False Then
                circle.Center = topCenterPoint
                circle.Diametr = userRack.Diameter
            Else
                circle = drawClass.CreateCircle(topCenterPoint, userRack.Diameter)
                If activProjectDocument.ActiveSpace.Entities.Contains(circle) = False Then
                    styleCounter.setObjectStyle(circle)
                End If
                Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackTop)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataTopCounter.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(circle, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
            End If
            result.Add(StructureElement.typeObject.counterRackTop, circle)
        ElseIf userRack.RackType = RackPillar.TypeRack.Octagonal Then
            If listPointTopCounter.Count > 0 Then
                If IsNothing(circle) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(circle)
                End If
                Dim listVertexPolygon As List(Of Vector3D) = BridgeGeometry.calculateVertexPolygon(userRack._elementBridgePoint.EndAxisPoint, userRack.Diameter / 2, 8, userRack.Rotation)
                If poly3dCounterTop.Count = 0 Then
                    poly3dCounterTop = drawClass.createPolyline3D(listVertexPolygon, True)
                    Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackTop)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                    dataTopCounter.KeyParameter = strGSON
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
                Else
                    Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listVertexPolygon)
                End If
            End If
            result.Add(StructureElement.typeObject.counterRackTop, poly3dCounterTop)
        Else
            If listPointTopCounter.Count > 0 Then
                If IsNothing(circle) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(circle)
                End If
                If poly3dCounterTop.Count = 0 Then
                    poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
                    Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackTop)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                    dataTopCounter.KeyParameter = strGSON
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
                Else
                    Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounter)
                End If
            End If
            result.Add(StructureElement.typeObject.counterRackTop, poly3dCounterTop)
        End If

        '===============================================================================================================================
        'контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Стойка (низ контура)")
        Dim dataBottomCounter As StructureElement = RackContour.getContour(dictionaryObjectsBridge, userRack.NumberPillar, userRack.Number, StructureElement.typeObject.counterRackBottom, userRack.NumberSubPillars)
        Dim poly3dCounterBottom As DwgPolyline3D = New DwgPolyline3D
        Dim bottomCircle As DwgCircle = New DwgCircle()
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = RackContour.createContour(idBridge, StructureElement.typeObject.counterRackBottom, userRack.RackType)
            If IsNothing(dataBottomCounter.DWGEntity) = False Then
                If TypeOf dataBottomCounter.DWGEntity Is DwgCircle Then
                    bottomCircle = dataBottomCounter.DWGEntity
                ElseIf TypeOf dataBottomCounter.DWGEntity Is DwgPolyline3D Then
                    poly3dCounterBottom = dataBottomCounter.DWGEntity
                End If
            End If
        Else
            If TypeOf dataBottomCounter.DWGEntity Is DwgPolyline3D Then
                poly3dCounterBottom = dataBottomCounter.DWGEntity
            ElseIf TypeOf dataBottomCounter.DWGEntity Is DwgCircle Then
                bottomCircle = dataBottomCounter.DWGEntity
            End If
        End If
        'рисуем контур
        If userRack.RackType = RackPillar.TypeRack.Circle Then
            'рисуем окружность
            If IsNothing(poly3dCounterBottom) = False Then
                activProjectDocument.ActiveSpace.Entities.Remove(poly3dCounterBottom)
            End If
            If IsNothing(bottomcircle) = False Then
                bottomCircle.Center = bottomCenterPoint
                bottomCircle.Diametr = userRack.Diameter
            Else
                bottomCircle = drawClass.CreateCircle(bottomCenterPoint, userRack.Diameter)
                If activProjectDocument.ActiveSpace.Entities.Contains(bottomCircle) = False Then
                    styleCounter.setObjectStyle(bottomCircle)
                End If
                Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackBottom)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataBottomCounter.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(bottomCircle, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
            End If
            result.Add(StructureElement.typeObject.counterRackBottom, bottomCircle)
        ElseIf userRack.RackType = RackPillar.TypeRack.Octagonal Then
            If listPointBottomCounter.Count > 0 Then
                If IsNothing(bottomCircle) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(bottomCircle)
                End If
                Dim listVertexPolygon As List(Of Vector3D) = BridgeGeometry.calculateVertexPolygon(userRack._elementBridgePoint.StartAxisPoint, userRack.Diameter / 2, 8, userRack.Rotation)
                If poly3dCounterBottom.Count = 0 Then
                    poly3dCounterBottom = drawClass.createPolyline3D(listVertexPolygon, True)
                    Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackBottom)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                    dataBottomCounter.KeyParameter = strGSON
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
                Else
                    Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listVertexPolygon)
                End If
            End If
            result.Add(StructureElement.typeObject.counterRackBottom, poly3dCounterBottom)
        Else
            If listPointBottomCounter.Count > 0 Then
                If IsNothing(bottomCircle) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(bottomCircle)
                End If
                If poly3dCounterBottom.Count = 0 Then
                    poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
                    Dim userCounter As RackContour = New RackContour(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number, StructureElement.typeObject.counterRackBottom)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                    dataBottomCounter.KeyParameter = strGSON
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
                Else
                    Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounter)
                End If
            End If
            result.Add(StructureElement.typeObject.counterRackBottom, poly3dCounterBottom)
        End If
        Return result
    End Function
End Class
