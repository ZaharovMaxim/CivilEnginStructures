Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class PreparationContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
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
    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementPillar As StructureElement = New StructureElement()
        elementPillar.Label = "Мосты и путепроводы"
        elementPillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPillar.ClassObject = StructureElement.classStructure.PreparationPillar
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterPreparationBottom Then
            elementPillar.Description = "Контур подготовки по низу"
        ElseIf type = StructureElement.typeObject.counterPreparationTop Then
            elementPillar.Description = "Контур подготовки по верху"
        End If
        elementPillar.KeyParameter = ""
        elementPillar.IdElement = Guid.NewGuid.ToString
        elementPillar.IdStructure = idBridge
        elementPillar.Note = ""
        elementPillar.DWGEntity = New DwgPolyline3D
        Return elementPillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий контур подготовки
    Public Shared Function getContour(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
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
                            Dim userObject As PreparationContour = tempData.getPreparationContourPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        dataContour = tempData
                                        Exit For
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
    'рисование контура подготовки
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPreparation As PreparationPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContour = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userPreparation) Then Return drawContour
        If IsNothing(userPreparation._elementBridgePoint.StartAxisPoint) = True Then Return drawContour
        If IsNothing(userPreparation._elementBridgePoint.EndAxisPoint) = True Then Return drawContour
        'вспомогательные построения
        Dim layerPreparation As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorPreparation As CadColor = New CadColor(7)
        Dim nameTypeLinePreparation As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLinePreparation As Integer = 1
        Dim widthTypeLinePreparation As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Подготовка (верх контура)")
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userPreparation._elementBridgePoint.ListPointModel.Count > 3 Then
            For i As Integer = 0 To userPreparation._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userPreparation._elementBridgePoint.ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next i
        End If
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = PreparationContour.getContour(dictionaryObjectsBridge, userPreparation.NumberPillar, StructureElement.typeObject.counterPreparationTop, userPreparation.Number)
        Dim poly3dCounterTop As DwgPolyline3D = New DwgPolyline3D
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = PreparationContour.createContour(idBridge, StructureElement.typeObject.counterPreparationTop)
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As PreparationContour = New PreparationContour(userPreparation.NumberPillar, userPreparation.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTop) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTop)
        End If
        styleCounter.setObjectStyle(poly3dCounterTop)
        drawContour.Add(StructureElement.typeObject.counterPreparationTop, poly3dCounterTop)
        '============================================================================================================================
        'находим старый контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Подготовка (низ контура)")
        Dim dataBottomCounter As StructureElement = PreparationContour.getContour(dictionaryObjectsBridge, userPreparation.NumberPillar, StructureElement.typeObject.counterPreparationBottom, userPreparation.Number)
        Dim poly3dCounterBottom As DwgPolyline3D = New DwgPolyline3D()
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = PreparationContour.createContour(idBridge, StructureElement.typeObject.counterPreparationBottom)
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As PreparationContour = New PreparationContour(userPreparation.NumberPillar, userPreparation.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottom) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottom)
        End If
        styleCounter.setObjectStyle(poly3dCounterBottom)
        drawContour.Add(StructureElement.typeObject.counterPreparationBottom, poly3dCounterBottom)
        Return drawContour
    End Function
End Class
