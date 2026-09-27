Imports System.ComponentModel
Imports CivilEnginStructures.Pillar
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class HandContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _type As SidePillarElement
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _type = SidePillarElement.None
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, SideHand As SidePillarElement)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _type = SideHand
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
    <Description("Тип элемента")>
    <Category("Свойства")>
    <DisplayName("Тип элемента")>
    <[ReadOnly](True)>
    Public Property SideHand() As SidePillarElement
        Get
            Return _type
        End Get
        Set(value As SidePillarElement)
            _type = value
        End Set
    End Property
    'создать контур откосного крыла
    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementCounter.Name = type
        If type = StructureElement.typeObject.counterLeftHandBottom Then
            elementCounter.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.counterLeftHandTop Then
            elementCounter.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.counterLeftHandCorniceBottom Then
            elementCounter.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.counterLeftHandCorniceTop Then
            elementCounter.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.counterRightHandBottom Then
            elementCounter.ClassObject = StructureElement.classStructure.HandRightPillar
        ElseIf type = StructureElement.typeObject.counterRightHandTop Then
            elementCounter.ClassObject = StructureElement.classStructure.HandRightPillar
        ElseIf type = StructureElement.typeObject.counterRightHandCorniceBottom Then
            elementCounter.ClassObject = StructureElement.classStructure.HandRightPillar
        ElseIf type = StructureElement.typeObject.counterRightHandCorniceTop Then
            elementCounter.ClassObject = StructureElement.classStructure.HandRightPillar
        End If
        elementCounter.Description = StructureElement.GetDescription(elementCounter.Name)
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgPolyline3D
        Return elementCounter
    End Function
    'функция ищет существующий контур шкафной стенки
    Public Shared Function getContour(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal sideUserHand As SidePillarElement, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
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
                            Dim userObject As HandContour = tempData.getHandCounterPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If userObject.SideHand = sideUserHand Then
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
    'рисование контура
    Public Shared Function drawContours(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userHand As HandPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContours = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userHand) Then Return drawContours
        'контур обратного открылка
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userHand._elementBridgePoint.ListPointModel.Count > 3 Then
            For i As Integer = 0 To userHand._elementBridgePoint.ListPointModel.Count - 1
                Dim pointBridge As PointStructure = userHand._elementBridgePoint.ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                listPointTopCounter.Add(topPoint)
                If pointBridge.dz2 <> 0 Then
                    If pointBridge.Code.IndexOf("leftPt") > -1 Then
                        Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                        listPointBottomCounter.Add(bottomPoint)
                        Dim bottomPoint2 As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz2)
                        listPointBottomCounter.Add(bottomPoint2)
                    ElseIf pointBridge.Code.IndexOf("rightPt") > -1 Then
                        Dim bottomPoint2 As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz2)
                        listPointBottomCounter.Add(bottomPoint2)
                        Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                        listPointBottomCounter.Add(bottomPoint)
                    End If
                Else
                    Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                    listPointBottomCounter.Add(bottomPoint)
                End If
            Next i
        End If
        'контур обратного открылка
        Dim listPointTopCornice As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCornice As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If userHand._elementBridgePoint.ListPointSecondModel.Count > 3 Then
            For i As Integer = 0 To userHand._elementBridgePoint.ListPointSecondModel.Count - 1
                Dim pointBridge As PointStructure = userHand._elementBridgePoint.ListPointSecondModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCornice.Add(topPoint)
                listPointBottomCornice.Add(bottomPoint)
            Next i
        End If
        '=====================================================================================================================
        'стиль по верху
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Откосное крыло (верх контура)")
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = Nothing
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            dataTopCounter = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterLeftHandTop, userHand.NumberSubPillar)
        Else
            dataTopCounter = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterRightHandTop, userHand.NumberSubPillar)
        End If
        Dim poly3dCounterTop As DwgPolyline3D = New DwgPolyline3D
        If IsNothing(dataTopCounter) = True Then
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                dataTopCounter = HandContour.createContour(idBridge, StructureElement.typeObject.counterLeftHandTop)
            Else
                dataTopCounter = HandContour.createContour(idBridge, StructureElement.typeObject.counterRightHandTop)
            End If
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As HandContour = New HandContour(userHand.NumberPillar, userHand.NumberSubPillar, SidePillarElement.None)
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                userCounter.SideHand = SidePillarElement.Left
            ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
                userCounter.SideHand = SidePillarElement.Right
            End If
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounter.KeyParameter = strGSON
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTop, listPointTopCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTop) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTop)
        End If
        styleCounter.setObjectStyle(poly3dCounterTop)
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(poly3dCounterTop, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounter)
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            drawContours.Add(StructureElement.typeObject.counterLeftHandTop, poly3dCounterTop)
        ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
            drawContours.Add(StructureElement.typeObject.counterRightHandTop, poly3dCounterTop)
        End If
        '============================================================================================================================
        'верх карниза
        Dim dataTopCounterCornice As StructureElement = New StructureElement()
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            dataTopCounterCornice = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterLeftHandCorniceTop, userHand.NumberSubPillar)
        Else
            dataTopCounterCornice = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterRightHandCorniceTop, userHand.NumberSubPillar)
        End If
        Dim poly3dCounterTopCornice As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounterCornice) = True Then
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                dataTopCounterCornice = HandContour.createContour(idBridge, StructureElement.typeObject.counterLeftHandCorniceTop)
            Else
                dataTopCounterCornice = HandContour.createContour(idBridge, StructureElement.typeObject.counterRightHandCorniceTop)
            End If
            poly3dCounterTopCornice = dataTopCounterCornice.DWGEntity
        Else
            poly3dCounterTopCornice = dataTopCounterCornice.DWGEntity
        End If
        If poly3dCounterTopCornice.Count = 0 Then
            poly3dCounterTopCornice = drawClass.createPolyline3D(listPointTopCornice, True)
            Dim userCounter As HandContour = New HandContour(userHand.NumberPillar, userHand.NumberSubPillar, SidePillarElement.None)
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                userCounter.SideHand = SidePillarElement.Left
            ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
                userCounter.SideHand = SidePillarElement.Right
            End If
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataTopCounterCornice.KeyParameter = strGSON
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterTopCornice, listPointTopCornice)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterTopCornice) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterTopCornice)
        End If
        styleCounter.setObjectStyle(poly3dCounterTopCornice)
        boolRecData = FuncXRecords.setXRecords(poly3dCounterTopCornice, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataTopCounterCornice)
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            drawContours.Add(StructureElement.typeObject.counterLeftHandCorniceTop, poly3dCounterTopCornice)
        ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
            drawContours.Add(StructureElement.typeObject.counterRightHandCorniceTop, poly3dCounterTopCornice)
        End If
        '============================================================================================================================
        'находим старый контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Откосное крыло (низ контура)")
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataBottomCounter As StructureElement = Nothing
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            dataBottomCounter = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterLeftHandBottom, userHand.NumberSubPillar)
        Else
            dataBottomCounter = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterRightHandBottom, userHand.NumberSubPillar)
        End If
        Dim poly3dCounterBottom As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounter) = True Then
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                dataBottomCounter = HandContour.createContour(idBridge, StructureElement.typeObject.counterLeftHandBottom)
            Else
                dataBottomCounter = HandContour.createContour(idBridge, StructureElement.typeObject.counterRightHandBottom)
            End If
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As HandContour = New HandContour(userHand.NumberPillar, userHand.NumberSubPillar, SidePillarElement.None)
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                userCounter.SideHand = SidePillarElement.Left
            ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
                userCounter.SideHand = SidePillarElement.Right
            End If
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounter.KeyParameter = strGSON
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottom, listPointBottomCounter)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottom) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottom)
        End If
        styleCounter.setObjectStyle(poly3dCounterBottom)
        boolRecData = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            drawContours.Add(StructureElement.typeObject.counterLeftHandBottom, poly3dCounterBottom)
        ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
            drawContours.Add(StructureElement.typeObject.counterRightHandBottom, poly3dCounterBottom)
        End If
        '============================================================================================================================
        'верх карниза
        Dim dataBottomCounterCornice As StructureElement = Nothing
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            dataBottomCounterCornice = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterLeftHandCorniceBottom, userHand.NumberSubPillar)
        Else
            dataBottomCounterCornice = HandContour.getContour(dictionaryObjectsBridge, userHand.NumberPillar, userHand.SideHand, StructureElement.typeObject.counterRightHandCorniceBottom, userHand.NumberSubPillar)
        End If
        Dim poly3dCounterBottomCornice As DwgPolyline3D = New DwgPolyline3D
        If IsNothing(dataBottomCounterCornice) = True Then
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                dataBottomCounterCornice = HandContour.createContour(idBridge, StructureElement.typeObject.counterLeftHandCorniceBottom)
            Else
                dataBottomCounterCornice = HandContour.createContour(idBridge, StructureElement.typeObject.counterRightHandCorniceBottom)
            End If
            poly3dCounterBottomCornice = dataBottomCounterCornice.DWGEntity
        Else
            poly3dCounterBottomCornice = dataBottomCounterCornice.DWGEntity
        End If
        If poly3dCounterBottomCornice.Count = 0 Then
            poly3dCounterBottomCornice = drawClass.createPolyline3D(listPointBottomCornice, True)
            Dim userCounter As HandContour = New HandContour(userHand.NumberPillar, userHand.NumberSubPillar, SidePillarElement.None)
            If userHand.SideHand = Pillar.SidePillarElement.Left Then
                userCounter.SideHand = SidePillarElement.Left
            ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
                userCounter.SideHand = SidePillarElement.Right
            End If
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounter)
            dataBottomCounterCornice.KeyParameter = strGSON
        Else
            Dim boolRedrawPline As Boolean = drawClass.reDrawPolyline3D(poly3dCounterBottomCornice, listPointBottomCornice)
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(poly3dCounterBottomCornice) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(poly3dCounterBottomCornice)
        End If
        styleCounter.setObjectStyle(poly3dCounterBottomCornice)
        boolRecData = FuncXRecords.setXRecords(poly3dCounterBottomCornice, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounterCornice)
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            drawContours.Add(StructureElement.typeObject.counterLeftHandCorniceBottom, poly3dCounterBottomCornice)
        ElseIf userHand.SideHand = Pillar.SidePillarElement.Right Then
            drawContours.Add(StructureElement.typeObject.counterRightHandCorniceBottom, poly3dCounterBottomCornice)
        End If
        Return drawContours
    End Function
End Class
