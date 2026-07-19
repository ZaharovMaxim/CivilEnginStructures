Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public Class NozzleContour
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
    Public Shared Function createContoursNozzle(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementNozzlePillar As StructureElement = New StructureElement()
        elementNozzlePillar.Label = "Мосты и путепроводы"
        elementNozzlePillar.ClassObject = StructureElement.classStructure.NozzlePillar
        elementNozzlePillar.Name = type
        If type = StructureElement.typeObject.contourNozzleBottom Then
            elementNozzlePillar.Description = "Контур насадки по низу"
        ElseIf type = StructureElement.typeObject.contourNozzleTop Then
            elementNozzlePillar.Description = "Контур насадки по верху"
        ElseIf type = StructureElement.typeObject.contourNozzleLeftConsole Then
            elementNozzlePillar.Description = "Левая консоль насадки"
        ElseIf type = StructureElement.typeObject.contourNozzleRightConsole Then
            elementNozzlePillar.Description = "Правая консоль насадки"
        ElseIf type = StructureElement.typeObject.contourNozzleCabinetWall Then
            elementNozzlePillar.Description = "Линия начала шкафной стенки"
        End If
        elementNozzlePillar.KeyParameter = ""
        elementNozzlePillar.IdElement = Guid.NewGuid.ToString
        elementNozzlePillar.IdStructure = idBridge
        elementNozzlePillar.Note = ""
        elementNozzlePillar.DWGEntity = New DwgPolyline3D
        Return elementNozzlePillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующую насадку
    Public Shared Function getContoursNozzle(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataContourNozzle As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(typeContour) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(typeContour)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As NozzleContour = tempData.getNozzleCounterPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        dataContourNozzle = tempData
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataContourNozzle
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование контура насадки
    Public Shared Function drawContoursNozzle(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userNozzle As NozzlePillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContoursNozzle = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userNozzle) Then Return drawContoursNozzle
        If IsNothing(userNozzle._elementBridgePoint.StartAxisPoint) = True Then Return drawContoursNozzle
        If IsNothing(userNozzle._elementBridgePoint.EndAxisPoint) = True Then Return drawContoursNozzle
        'вспомогательные построения
        Dim layerNozzle As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorNozzle As CadColor = New CadColor(7)
        Dim nameTypeLineNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineNozzle As Integer = 1
        Dim widthTypeLineNozzle As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounterNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounterNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Насадка (контур)")
        Dim listPointTopCounterNozzle As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounterNozzle As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userNozzle._elementBridgePoint.ListPointModel.Count > 3 Then
            Dim listPointLeftConsole As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
            Dim listPointRightConsole As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
            For i As Integer = 0 To userNozzle._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userNozzle._elementBridgePoint.ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounterNozzle.Add(topPoint)
                listPointBottomCounterNozzle.Add(bottomPoint)
            Next i
        End If
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = NozzleContour.getContoursNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, StructureElement.typeObject.contourNozzleTop, userNozzle.Number)
        Dim poly3dCounterTop As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounter) = True Then
            dataTopCounter = NozzleContour.createContoursNozzle(idBridge, StructureElement.typeObject.contourNozzleTop)
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounterNozzle, True)
            Dim userCounter As NozzleContour = New NozzleContour(userNozzle.NumberPillar, userNozzle.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounterNozzle)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTop) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTop)
            styleCounterNozzle.setObjectStyle(poly3dCounterTop)
        End If
        drawContoursNozzle.Add(StructureElement.typeObject.contourNozzleTop, poly3dCounterTop)
        '============================================================================================================================
        'находим старый контур по низу
        Dim dataBottomCounter As StructureElement = NozzleContour.getContoursNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, StructureElement.typeObject.contourNozzleBottom, userNozzle.Number)
        Dim poly3dCounterBottom As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounter) = True Then
            dataBottomCounter = NozzleContour.createContoursNozzle(idBridge, StructureElement.typeObject.contourNozzleBottom)
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounterNozzle, True)
            Dim userCounter As NozzleContour = New NozzleContour(userNozzle.NumberPillar, userNozzle.Number)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounter.KeyParameter = strGSON
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounterNozzle)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottom) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottom)
            styleCounterNozzle.setObjectStyle(poly3dCounterBottom)
        End If
        drawContoursNozzle.Add(StructureElement.typeObject.contourNozzleBottom, poly3dCounterBottom)
        '============================================================================================================================
        'линия шкафной стенки
        Dim middlePointLeft As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        Dim middlePointRight As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userNozzle._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userNozzle._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "middlePt1" Then
                middlePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "middlePt2" Then
                middlePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If middlePointLeft.X <> 0 And middlePointLeft.Y <> 0 And middlePointRight.X <> 0 And middlePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {middlePointLeft, middlePointRight}
            Dim dataLineCabinetWall As StructureElement = NozzleContour.getContoursNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, StructureElement.typeObject.contourNozzleCabinetWall, userNozzle.Number)
            Dim poly3dCabinetWall As DwgPolyline3D = Nothing
            If IsNothing(dataLineCabinetWall) = True Then
                dataLineCabinetWall = NozzleContour.createContoursNozzle(idBridge, StructureElement.typeObject.contourNozzleCabinetWall)
                poly3dCabinetWall = dataLineCabinetWall.DWGEntity
            Else
                poly3dCabinetWall = dataLineCabinetWall.DWGEntity
            End If
            If poly3dCabinetWall.Count = 0 Then
                poly3dCabinetWall = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As NozzleContour = New NozzleContour(userNozzle.NumberPillar, userNozzle.Number)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineCabinetWall.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCabinetWall, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineCabinetWall)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCabinetWall, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCabinetWall) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dCabinetWall)
                styleCounterNozzle.setObjectStyle(poly3dCabinetWall)
            End If
        End If
        '============================================================================================================================
        'левая консоль
        Dim consolePointLeft As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        Dim consolePointRight As Cad.Foundation.Vector3D = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userNozzle._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userNozzle._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "leftConsol1" Then
                consolePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "leftConsol2" Then
                consolePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If consolePointLeft.X <> 0 And consolePointLeft.Y <> 0 And consolePointRight.X <> 0 And consolePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {consolePointLeft, consolePointRight}
            Dim dataLineLeftConsole As StructureElement = NozzleContour.getContoursNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, StructureElement.typeObject.contourNozzleLeftConsole, userNozzle.Number)
            Dim poly3dLeftConsole As DwgPolyline3D = Nothing
            If IsNothing(dataLineLeftConsole) = True Then
                dataLineLeftConsole = NozzleContour.createContoursNozzle(idBridge, StructureElement.typeObject.contourNozzleLeftConsole)
                poly3dLeftConsole = dataLineLeftConsole.DWGEntity
            Else
                poly3dLeftConsole = dataLineLeftConsole.DWGEntity
            End If
            If poly3dLeftConsole.Count = 0 Then
                poly3dLeftConsole = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As NozzleContour = New NozzleContour(userNozzle.NumberPillar, userNozzle.Number)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineLeftConsole.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dLeftConsole, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineLeftConsole)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dLeftConsole, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dLeftConsole) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dLeftConsole)
                styleCounterNozzle.setObjectStyle(poly3dLeftConsole)
            End If
        End If
        '============================================================================================================================
        'правая консоль
        consolePointLeft = New Cad.Foundation.Vector3D
        consolePointRight = New Cad.Foundation.Vector3D
        For i As Integer = 0 To userNozzle._elementBridgePoint.ListPointModel.Count - 1
            Dim pointBridge As PointStructure = userNozzle._elementBridgePoint.ListPointModel.ElementAt(i).Value
            If pointBridge.Code Like "rightConsol1" Then
                consolePointLeft = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            ElseIf pointBridge.Code Like "rightConsol2" Then
                consolePointRight = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
            End If
        Next i
        If consolePointLeft.X <> 0 And consolePointLeft.Y <> 0 And consolePointRight.X <> 0 And consolePointRight.Y <> 0 Then
            Dim listPoint As List(Of Vector3D) = New List(Of Vector3D) From {consolePointLeft, consolePointRight}
            Dim dataLineRightConsole As StructureElement = NozzleContour.getContoursNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, StructureElement.typeObject.contourNozzleRightConsole, userNozzle.Number)
            Dim poly3dRightConsole As DwgPolyline3D = Nothing
            If IsNothing(dataLineRightConsole) = True Then
                dataLineRightConsole = NozzleContour.createContoursNozzle(idBridge, StructureElement.typeObject.contourNozzleRightConsole)
                poly3dRightConsole = dataLineRightConsole.DWGEntity
            Else
                poly3dRightConsole = dataLineRightConsole.DWGEntity
            End If
            If poly3dRightConsole.Count = 0 Then
                poly3dRightConsole = drawClass.createPolyline3D(listPoint, False)
                Dim userCounter As NozzleContour = New NozzleContour(userNozzle.NumberPillar, userNozzle.Number)
                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
                dataLineRightConsole.KeyParameter = strGSON
                Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dRightConsole, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataLineRightConsole)
            Else
                Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dRightConsole, listPoint)
            End If
            If activProjectDocument.ActiveSpace.Entities.Contains(poly3dRightConsole) = False Then
                activProjectDocument.ActiveSpace.Entities.Add(poly3dRightConsole)
                styleCounterNozzle.setObjectStyle(poly3dRightConsole)
            End If
        End If
    End Function
End Class
