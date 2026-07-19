Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class CabinetWallContour
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
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassObject = StructureElement.classStructure.CabinetWallPillar
        elementCounter.Name = type
        If type = StructureElement.typeObject.contourCabinetWallBottom Then
            elementCounter.Description = "Контур шкафной стенки по низу"
        ElseIf type = StructureElement.typeObject.contourCabinetWallTop Then
            elementCounter.Description = "Контур шкафной стенки по верху"
        ElseIf type = StructureElement.typeObject.contourCabinetWallPlateBottom Then
            elementCounter.Description = "Контур зуба упора шкафной стенки по низу"
        ElseIf type = StructureElement.typeObject.contourCabinetWallPlateTop Then
            elementCounter.Description = "Контур зуба упора шкафной стенки по верху"
        End If
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgPolyline3D
        Return elementCounter
    End Function

    'функция ищет существующий контур шкафной стенки
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
                            Dim userObject As CabinetWallContour = tempData.getCabinetWallCounterPillar
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

    Public Shared Function drawContours(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userCabinetWall As CabinetWallPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContours = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userCabinetWall) Then Return drawContours
        If IsNothing(userCabinetWall._elementBridgePoint.StartAxisPoint) = True Then Return drawContours
        If IsNothing(userCabinetWall._elementBridgePoint.EndAxisPoint) = True Then Return drawContours
        'вспомогательные построения
        Dim layerContour As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorContour As CadColor = New CadColor(7)
        Dim nameTypeLineContour As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineContour As Integer = 1
        Dim widthTypeLineContour As Integer = 20
        'контур шкафной стенки
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userCabinetWall._elementBridgePoint.ListPointModel.Count > 3 Then
            For i As Integer = 0 To userCabinetWall._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userCabinetWall._elementBridgePoint.ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next i
        End If
        'контур зуба упора
        Dim listPointTopPlate As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomPlate As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userCabinetWall._elementBridgePoint.ListPointSecondModel.Count > 3 Then
            For i As Integer = 0 To userCabinetWall._elementBridgePoint.ListPointSecondModel.Count - 1
                Dim pointBridge As PointStructure = userCabinetWall._elementBridgePoint.ListPointSecondModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopPlate.Add(topPoint)
                listPointBottomPlate.Add(bottomPoint)
            Next i
        End If
        '============================================================================================================================
        'стиль по верху
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Шкафная стенка (верх контура)")
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = CabinetWallContour.getContour(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.contourCabinetWallTop, userCabinetWall.Number)
        Dim poly3dCounterTop As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = CabinetWallContour.createContour(idBridge, StructureElement.typeObject.contourCabinetWallTop)
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As CabinetWallContour = New CabinetWallContour(userCabinetWall.NumberPillar, userCabinetWall.Number)
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
        drawContours.Add(StructureElement.typeObject.contourCabinetWallTop, poly3dCounterTop)
        '============================================================================================================================
        'верх зуба упора
        Dim dataTopCounterPlate As StructureElement = CabinetWallContour.getContour(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.contourCabinetWallPlateTop, userCabinetWall.Number)
        Dim poly3dCounterTopPlate As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounterPlate) = True Then
            dataTopCounterPlate = CabinetWallContour.createContour(idBridge, StructureElement.typeObject.contourCabinetWallPlateTop)
            poly3dCounterTopPlate = dataTopCounterPlate.DWGEntity
        Else
            poly3dCounterTopPlate = dataTopCounterPlate.DWGEntity
        End If
        If poly3dCounterTopPlate.Count = 0 Then
            poly3dCounterTopPlate = drawClass.createPolyline3D(listPointTopPlate, True)
            Dim userCounter As CabinetWallContour = New CabinetWallContour(userCabinetWall.NumberPillar, userCabinetWall.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounterPlate.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTopPlate, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounterPlate)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTopPlate, listPointTopPlate)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTopPlate) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTopPlate)
            styleCounter.setObjectStyle(poly3dCounterTopPlate)
        End If
        drawContours.Add(StructureElement.typeObject.contourCabinetWallPlateTop, poly3dCounterTopPlate)
        '============================================================================================================================
        'находим старый контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Шкафная стенка (низ контура)")
        Dim dataBottomCounter As StructureElement = CabinetWallContour.getContour(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.contourCabinetWallBottom, userCabinetWall.Number)
        Dim poly3dCounterBottom As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = CabinetWallContour.createContour(idBridge, StructureElement.typeObject.contourCabinetWallBottom)
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As CabinetWallContour = New CabinetWallContour(userCabinetWall.NumberPillar, userCabinetWall.Number)
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
        drawContours.Add(StructureElement.typeObject.contourCabinetWallBottom, poly3dCounterBottom)
        '============================================================================================================================
        'находим старый контур по низу зуба упора
        Dim dataBottomCounterPlate As StructureElement = Nothing
        dataBottomCounterPlate = CabinetWallContour.getContour(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.contourCabinetWallPlateBottom, userCabinetWall.Number)
        Dim poly3dCounterBottomPlate As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounterPlate) = True Then
            dataBottomCounterPlate = CabinetWallContour.createContour(idBridge, StructureElement.typeObject.contourCabinetWallPlateBottom)
            poly3dCounterBottomPlate = dataBottomCounterPlate.DWGEntity
        Else
            poly3dCounterBottomPlate = dataBottomCounterPlate.DWGEntity
        End If
        If poly3dCounterBottomPlate.Count = 0 Then
            poly3dCounterBottomPlate = drawClass.createPolyline3D(listPointBottomPlate, True)
            Dim userCounter As CabinetWallContour = New CabinetWallContour(userCabinetWall.NumberPillar, userCabinetWall.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounterPlate.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottomPlate, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounterPlate)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottomPlate, listPointBottomPlate)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottomPlate) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottomPlate)
            styleCounter.setObjectStyle(poly3dCounterBottomPlate)
        End If
        drawContours.Add(StructureElement.typeObject.contourCabinetWallPlateBottom, poly3dCounterBottomPlate)
        Return drawContours
    End Function

End Class
