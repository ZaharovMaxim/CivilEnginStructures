Imports System.ComponentModel
Imports CivilEnginStructures.RackContour
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class RigelContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _type As StructureElement.typeObject
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _type = StructureElement.typeObject.OtherElement
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, TypeContour As StructureElement.typeObject)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
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
    <Description("Тип объекта")>
    <Category("Свойства")>
    <DisplayName("Тип объекта")>
    <[ReadOnly](True)>
    Public Property TypeRigelContour() As StructureElement.typeObject
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
        elementPillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPillar.ClassObject = StructureElement.classStructure.RigelPillar
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterRigelBottom Then
            elementPillar.Description = "Контур ригеля по низу"
        ElseIf type = StructureElement.typeObject.counterRigelTop Then
            elementPillar.Description = "Контур ригеля по верху"
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
                            Dim userObject As RigelContour = tempData.getRigelContourPillar
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
    'рисование контура 
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userRigel As RigelPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContour = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userRigel) Then Return drawContour
        If IsNothing(userRigel._elementBridgePoint.StartAxisPoint) = True Then Return drawContour
        If IsNothing(userRigel._elementBridgePoint.EndAxisPoint) = True Then Return drawContour
        'вспомогательные построения
        Dim layerRigel As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorRigel As CadColor = New CadColor(7)
        Dim nameTypeLineRigel As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineRigel As Integer = 1
        Dim widthTypeLineRigel As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleTopCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleTopCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Ригель (верх контура)")
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userRigel._elementBridgePoint.ListPointModel.Count > 3 Then
            For i As Integer = 0 To userRigel._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userRigel._elementBridgePoint.ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next i
        End If
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = RigelContour.getContour(dictionaryObjectsBridge, userRigel.NumberPillar, StructureElement.typeObject.counterRigelTop, userRigel.Number)
        Dim poly3dCounterTop As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = RigelContour.createContour(idBridge, StructureElement.typeObject.counterRigelTop)
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As RigelContour = New RigelContour(userRigel.NumberPillar, userRigel.Number, StructureElement.typeObject.counterRigelTop)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTop) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTop)
        End If
        styleTopCounter.setObjectStyle(poly3dCounterTop)
        drawContour.Add(StructureElement.typeObject.counterRigelTop, poly3dCounterTop)
        '============================================================================================================================
        'находим старый контур по низу
        Dim styleBottomCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleBottomCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ригель (низ контура)")
        Dim dataBottomCounter As StructureElement = RigelContour.getContour(dictionaryObjectsBridge, userRigel.NumberPillar, StructureElement.typeObject.counterRigelBottom, userRigel.Number)
        Dim poly3dCounterBottom As DwgPolyline3D = New DwgPolyline3D
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = RigelContour.createContour(idBridge, StructureElement.typeObject.counterRigelBottom)
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As RigelContour = New RigelContour(userRigel.NumberPillar, userRigel.Number, StructureElement.typeObject.counterRigelBottom)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottom) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottom)
        End If
        styleBottomCounter.setObjectStyle(poly3dCounterBottom)
        drawContour.Add(StructureElement.typeObject.counterRigelBottom, poly3dCounterBottom)
        '============================================================================================================================
        'линия слива
        Dim middlePointLeft As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        Dim middlePointRight As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userRigel._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userRigel._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "middlePt1" Then
                middlePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "middlePt2" Then
                middlePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If middlePointLeft.X <> 0 And middlePointLeft.Y <> 0 And middlePointRight.X <> 0 And middlePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {middlePointLeft, middlePointRight}
            Dim dataLineDrain As StructureElement = RigelContour.getContour(dictionaryObjectsBridge, userRigel.NumberPillar, StructureElement.typeObject.contourRigelCenter, userRigel.Number)
            Dim poly3dDrainLine As DwgPolyline3D = Nothing
            If IsNothing(dataLineDrain) = True Then
                dataLineDrain = RigelContour.createContour(idBridge, StructureElement.typeObject.contourRigelCenter)
                poly3dDrainLine = dataLineDrain.DWGEntity
            Else
                poly3dDrainLine = dataLineDrain.DWGEntity
            End If
            If poly3dDrainLine.Count = 0 Then
                poly3dDrainLine = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As RigelContour = New RigelContour(userRigel.NumberPillar, userRigel.Number, StructureElement.typeObject.contourRigelCenter)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineDrain.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dDrainLine, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineDrain)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dDrainLine, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dDrainLine) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dDrainLine)
            End If
            styleTopCounter.setObjectStyle(poly3dDrainLine)
        End If
        '============================================================================================================================
        'левая консоль
        Dim consolePointLeft As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        Dim consolePointRight As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userRigel._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userRigel._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "leftConsol1" Then
                consolePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "rightConsol1" Then
                consolePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If consolePointLeft.X <> 0 And consolePointLeft.Y <> 0 And consolePointRight.X <> 0 And consolePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {consolePointLeft, consolePointRight}
            Dim dataLineLeftConsole As StructureElement = RigelContour.getContour(dictionaryObjectsBridge, userRigel.NumberPillar, StructureElement.typeObject.contourRigelLeftConsole, userRigel.Number)
            Dim poly3dLeftConsole As DwgPolyline3D = Nothing
            If IsNothing(dataLineLeftConsole) = True Then
                dataLineLeftConsole = RigelContour.createContour(idBridge, StructureElement.typeObject.contourRigelLeftConsole)
                poly3dLeftConsole = dataLineLeftConsole.DWGEntity
            Else
                poly3dLeftConsole = dataLineLeftConsole.DWGEntity
            End If
            If poly3dLeftConsole.Count = 0 Then
                poly3dLeftConsole = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As RigelContour = New RigelContour(userRigel.NumberPillar, userRigel.Number, StructureElement.typeObject.contourRigelLeftConsole)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineLeftConsole.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dLeftConsole, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineLeftConsole)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dLeftConsole, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dLeftConsole) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dLeftConsole)
            End If
            styleBottomCounter.setObjectStyle(poly3dLeftConsole)
        End If
        '============================================================================================================================
        'правая консоль
        consolePointLeft = New Cad.Foundation.Vector3D
        consolePointRight = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userRigel._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userRigel._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "leftConsol2" Then
                consolePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "rightConsol2" Then
                consolePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If consolePointLeft.X <> 0 And consolePointLeft.Y <> 0 And consolePointRight.X <> 0 And consolePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {consolePointLeft, consolePointRight}
            Dim dataLineRightConsole As StructureElement = RigelContour.getContour(dictionaryObjectsBridge, userRigel.NumberPillar, StructureElement.typeObject.contourRigelRightConsole, userRigel.Number)
            Dim poly3dRightConsole As DwgPolyline3D = Nothing
            If IsNothing(dataLineRightConsole) = True Then
                dataLineRightConsole = RigelContour.createContour(idBridge, StructureElement.typeObject.contourRigelRightConsole)
                poly3dRightConsole = dataLineRightConsole.DWGEntity
            Else
                poly3dRightConsole = dataLineRightConsole.DWGEntity
            End If
            If poly3dRightConsole.Count = 0 Then
                poly3dRightConsole = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As RigelContour = New RigelContour(userRigel.NumberPillar, userRigel.Number, StructureElement.typeObject.contourRigelRightConsole)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineRightConsole.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dRightConsole, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineRightConsole)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dRightConsole, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dRightConsole) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dRightConsole)
            End If
            styleBottomCounter.setObjectStyle(poly3dRightConsole)
        End If
        Return drawContour
    End Function
End Class
