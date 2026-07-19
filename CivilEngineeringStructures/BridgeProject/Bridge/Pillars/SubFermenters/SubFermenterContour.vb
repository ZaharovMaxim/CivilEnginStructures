Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization.Runtime
Public Class SubFermenterContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberColumn As Integer
    Private _numberRow As Integer
    Private _type As StructureElement.classStructure
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberColumn = 0
        _numberRow = 0
        _type = StructureElement.typeObject.OtherElement
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberColumn As Integer, NumberRow As Integer, TypeContour As StructureElement.typeObject)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberColumn = NumberColumn
        _numberRow = NumberRow
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
    <Description("Номер пролета")>
    <Category("Свойства")>
    <DisplayName("Номер пролета")>
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
    <Description("Номер ряда балок")>
    <Category("Свойства")>
    <DisplayName("Номер ряда балок")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
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

    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementPillar As StructureElement = New StructureElement()
        elementPillar.Label = "Мосты и путепроводы"
        elementPillar.ClassObject = StructureElement.classStructure.SubFermenters
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterSubFermentersBottom Then
            elementPillar.Description = "Контур подферменника по низу"
        ElseIf type = StructureElement.typeObject.counterSubFermentersTop Then
            elementPillar.Description = "Контур подферменника по верху"
        End If
        elementPillar.KeyParameter = ""
        elementPillar.IdElement = Guid.NewGuid.ToString
        elementPillar.IdStructure = idBridge
        elementPillar.Note = ""
        elementPillar.DWGEntity = New DwgPolyline3D
        Return elementPillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий контур 
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
                            Dim userObject As SubFermenterContour = tempData.getSubFermentersContour
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
    'рисование контура 
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userSubFermenter As SubFermenters, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContour = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userSubFermenter) Then Return drawContour
        If userSubFermenter._elementBridgePoint.ListPointModel.Count < 2 Then Return drawContour
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Подферменник (верх контура)")

        Dim arrayCounters As DwgPolyline3D() = {Nothing, Nothing}
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userSubFermenter._elementBridgePoint.ListPointModel.Count > 3 Then
            For k As Integer = 0 To userSubFermenter._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userSubFermenter._elementBridgePoint.ListPointModel.ElementAt(k).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next k
        End If
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = SubFermenterContour.getContour(dictionaryObjectsBridge, userSubFermenter.NumberPillar, userSubFermenter.NumberProlet, userSubFermenter.NumberRow, StructureElement.typeObject.counterSubFermentersTop, userSubFermenter.NumberSubPillar)
        Dim poly3dCounterTop As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = SubFermenterContour.createContour(idBridge, StructureElement.typeObject.counterSubFermentersTop)
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As SubFermenterContour = New SubFermenterContour(userSubFermenter.NumberPillar, userSubFermenter.NumberSubPillar, userSubFermenter.NumberProlet, userSubFermenter.NumberRow, StructureElement.typeObject.counterSubFermentersTop)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTop) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTop)
            styleCounter.setObjectStyle(poly3dCounterTop)
        End If
        arrayCounters(0) = poly3dCounterTop
        '============================================================================================================================
        'находим старый контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Подферменник (низ контура)")
        Dim dataBottomCounter As StructureElement = SubFermenterContour.getContour(dictionaryObjectsBridge, userSubFermenter.NumberPillar, userSubFermenter.NumberProlet, userSubFermenter.NumberRow, StructureElement.typeObject.counterSubFermentersBottom, userSubFermenter.NumberSubPillar)
        Dim poly3dCounterBottom As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = SubFermenterContour.createContour(idBridge, StructureElement.typeObject.counterSubFermentersBottom)
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As SubFermenterContour = New SubFermenterContour(userSubFermenter.NumberPillar, userSubFermenter.NumberSubPillar, userSubFermenter.NumberProlet, userSubFermenter.NumberRow, StructureElement.typeObject.counterSubFermentersBottom)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottom) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottom)
            styleCounter.setObjectStyle(poly3dCounterBottom)
        End If
        arrayCounters(1) = poly3dCounterBottom
        drawContour.Add(StructureElement.typeObject.counterSubFermentersTop, poly3dCounterTop)
        drawContour.Add(StructureElement.typeObject.counterSubFermentersBottom, poly3dCounterBottom)
        Return drawContour
    End Function
End Class
