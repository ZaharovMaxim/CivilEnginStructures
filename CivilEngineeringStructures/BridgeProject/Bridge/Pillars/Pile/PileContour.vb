Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class PileContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberColumn As Integer
    Private _numberRow As Integer
    Private _type As StructureElement.typeObject
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberColumn = 0
        _numberRow = 0
        _type = StructureElement.typeObject.counterPileTop
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberColumn As Integer, NumberRow As Integer, TypeCounter As StructureElement.typeObject)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberColumn = NumberColumn
        _numberRow = NumberRow
        _type = TypeCounter
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
    <Browsable(False)>
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
    <Description("Номер столбца")>
    <Category("Свойства")>
    <DisplayName("Номер столбца")>
    <[ReadOnly](True)>
    Public Property NumberColumn() As Integer
        Get
            Return _numberColumn
        End Get
        Set(value As Integer)
            _numberColumn = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    'создать контур сваи
    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject, ByVal typeCounter As PilePillar.TypePile) As StructureElement
        Dim elementPillar As StructureElement = New StructureElement()
        elementPillar.Label = "Мосты и путепроводы"
        elementPillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPillar.ClassObject = StructureElement.classStructure.PilePillar
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterPileTop Then
            elementPillar.Description = "Контур сваи по верху"
        ElseIf type = StructureElement.typeObject.counterPileBottom Then
            elementPillar.Description = "Контур сваи по низу"
        End If
        elementPillar.KeyParameter = ""
        elementPillar.IdElement = Guid.NewGuid.ToString
        elementPillar.IdStructure = idBridge
        elementPillar.Note = ""
        If typeCounter = PilePillar.TypePile.Drilling Then
            elementPillar.DWGEntity = New DwgCircle
        Else
            elementPillar.DWGEntity = New DwgPolyline3D
        End If
        Return elementPillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий контур сваи
    Public Shared Function getContour(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberColumn As Integer, ByVal numberRow As Integer, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
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
                            Dim userObject As PileContour = tempData.getPileCounterPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If numberColumn = userObject.NumberColumn Then
                                            If numberRow = userObject.NumberRow Then
                                                dataContour = tempData
                                                Exit For
                                            End If
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
    'рисование контура сваи
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPile As PilePillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userPile) Then Return False
        If IsNothing(userPile._elementBridgePoint.CenterTopPoint) = True Then Return False
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Свая (верх контура)")
        Dim centerPoint As Vector3D = userPile._elementBridgePoint.CenterTopPoint
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = PileContour.getContour(dictionaryObjectsBridge, userPile.NumberPillar, userPile.NumberColumn, userPile.NumberRow, StructureElement.typeObject.counterPileTop, userPile.NumberSubPillars)
        Dim poly3dCounterTop As DwgPolyline3D = New DwgPolyline3D
        Dim circle As DwgCircle = New DwgCircle
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = PileContour.createContour(idBridge, StructureElement.typeObject.counterPileTop, userPile.Type)
            If IsNothing(dataTopCounter.DWGEntity) = False Then
                If userPile.Type = PilePillar.TypePile.Drilling Then
                    circle = dataTopCounter.DWGEntity
                Else
                    poly3dCounterTop = dataTopCounter.DWGEntity
                End If
            End If
        Else
            If TypeOf dataTopCounter.DWGEntity Is DwgCircle Then
                circle = dataTopCounter.DWGEntity
            ElseIf TypeOf dataTopCounter.DWGEntity Is DwgPolyline3D Then
                poly3dCounterTop = dataTopCounter.DWGEntity
            End If
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        'рисуем сваю
        If userPile.Type = PilePillar.TypePile.Drilling Then
            'рисуем окружность
            If IsNothing(poly3dCounterTop) = False Then
                activProjectDocument.ActiveSpace.Entities.Remove(poly3dCounterTop)
            End If
            If IsNothing(circle) = False Then
                circle.Center = centerPoint
                circle.Diametr = userPile.Diameter
            Else
                circle = drawClass.CreateCircle(centerPoint, userPile.Diameter)
                styleCounter.setObjectStyle(circle)
                Dim userCounter As PileContour = New PileContour(userPile.NumberPillar, userPile.NumberSubPillars, userPile.NumberColumn, userPile.NumberRow, StructureElement.typeObject.counterPileTop)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataTopCounter.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(circle, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
            End If
        Else
            Dim listPoint As List(Of Vector3D) = BridgeGeometry.calculateRotatedRectangle3d(centerPoint, userPile.Width, userPile.Width, userPile.Rotation)
            If listPoint.Count > 0 Then
                If IsNothing(circle) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(circle)
                End If
                If poly3dCounterTop.Count = 0 Then
                    poly3dCounterTop = drawClass.createPolyline3D(listPoint, True)
                    Dim userCounter As PileContour = New PileContour(userPile.NumberPillar, userPile.NumberSubPillars, userPile.NumberColumn, userPile.NumberRow, StructureElement.typeObject.counterPileTop)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                    dataTopCounter.KeyParameter = strGSON
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
                    styleCounter.setObjectStyle(poly3dCounterTop)
                Else
                    Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPoint)
                End If
            End If
        End If
        Return True
    End Function
End Class
