Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class HandModel
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер подопоры
    Private _sideElement As Pillar.SidePillarElement 'право или лево

    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _sideElement = Pillar.SidePillarElement.None
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, SideElement As Pillar.SidePillarElement)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _sideElement = SideElement
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
    Public Property SideModel() As Pillar.SidePillarElement
        Get
            Return _sideElement
        End Get
        Set(value As Pillar.SidePillarElement)
            _sideElement = value
        End Set
    End Property
    Public Shared Function createModel(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementModel As StructureElement = New StructureElement()
        elementModel.Label = "Мосты и путепроводы"
        elementModel.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementModel.Name = type
        If type = StructureElement.typeObject.modelLeftHand Then
            elementModel.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.modelLeftHandCornice Then
            elementModel.ClassObject = StructureElement.classStructure.HandLeftPillar
        ElseIf type = StructureElement.typeObject.modelRightHand Then
            elementModel.ClassObject = StructureElement.classStructure.HandRightPillar
        ElseIf type = StructureElement.typeObject.modelRightHandCornice Then
            elementModel.ClassObject = StructureElement.classStructure.HandRightPillar
        End If
        elementModel.Description = StructureElement.GetDescription(type)
        elementModel.KeyParameter = ""
        elementModel.IdElement = Guid.NewGuid.ToString
        elementModel.IdStructure = idBridge
        elementModel.Note = ""
        elementModel.DWGEntity = New DwgModel3DElement
        Return elementModel
    End Function

    'функция ищет существующий контур
    Public Shared Function getModel(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal typeModel As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataModel As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(typeModel) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(typeModel)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As HandModel = tempData.getHandModelPillar()
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        dataModel = tempData
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataModel
    End Function

    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userHand As HandPillar, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userHand) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterHand As DwgPolyline3D = Nothing
        Dim bottomCounterHand As DwgPolyline3D = Nothing
        Dim topCounterHandCornice As DwgPolyline3D = Nothing
        Dim bottomCounterCornice As DwgPolyline3D = Nothing
        If userHand.SideHand = Pillar.SidePillarElement.Left Then
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftHandTop) = True Then
                topCounterHand = dictPolyline.Item(StructureElement.typeObject.counterLeftHandTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftHandBottom) = True Then
                bottomCounterHand = dictPolyline.Item(StructureElement.typeObject.counterLeftHandBottom)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftHandCorniceTop) = True Then
                topCounterHandCornice = dictPolyline.Item(StructureElement.typeObject.counterLeftHandCorniceTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftHandCorniceBottom) = True Then
                bottomCounterCornice = dictPolyline.Item(StructureElement.typeObject.counterLeftHandCorniceBottom)
            End If
        Else
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightHandTop) = True Then
                topCounterHand = dictPolyline.Item(StructureElement.typeObject.counterRightHandTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightHandBottom) = True Then
                bottomCounterHand = dictPolyline.Item(StructureElement.typeObject.counterRightHandBottom)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightHandCorniceTop) = True Then
                topCounterHandCornice = dictPolyline.Item(StructureElement.typeObject.counterRightHandCorniceTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightHandCorniceBottom) = True Then
                bottomCounterCornice = dictPolyline.Item(StructureElement.typeObject.counterRightHandCorniceBottom)
            End If
        End If
        If IsNothing(topCounterHand) = False Then
            If IsNothing(bottomCounterHand) = False Then
                If topCounterHand.Count > 3 And bottomCounterHand.Count > 3 Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Откосное крыло (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim dataModel As StructureElement = Nothing
                    Dim modelHand As DwgModel3DElement = New DwgModel3DElement()
                    If userHand.SideHand = Pillar.SidePillarElement.Left Then
                        dataModel = HandModel.getModel(dictionaryObjectsBridge, userHand.NumberPillar, StructureElement.typeObject.modelLeftHand, userHand.NumberSubPillar)
                    Else
                        dataModel = HandModel.getModel(dictionaryObjectsBridge, userHand.NumberPillar, StructureElement.typeObject.modelRightHand, userHand.NumberSubPillar)
                    End If
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            modelHand = dwgModel
                        End If
                    Else
                        If userHand.SideHand = Pillar.SidePillarElement.Left Then
                            dataModel = createModel(idBridge, StructureElement.typeObject.modelLeftHand)
                        Else
                            dataModel = createModel(idBridge, StructureElement.typeObject.modelRightHand)
                        End If
                    End If
                    Dim axisStartPoint As Vector3D = userHand._elementBridgePoint.StartAxisPoint
                    Dim axisEndPoint As Vector3D = userHand._elementBridgePoint.EndAxisPoint
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterHand, bottomCounterHand)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim centreAxisModel As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(axisStartPoint, axisEndPoint)
                            Dim elementModel = New StaticSolidElement("Откосное крыло (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = centreAxisModel
                            modelHand.Position = elementModel.Origin
                            modelHand.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelHand) = False Then
                                activProjectDocument.ActiveSpace.Add(modelHand)
                                Dim userModel As HandModel = New HandModel(userHand.NumberPillar, userHand.NumberSubPillar, userHand.SideHand)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = modelHand
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelHand, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(modelHand)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If IsNothing(topCounterHandCornice) = False Then
            If IsNothing(bottomCounterCornice) = False Then
                If topCounterHandCornice.Count > 3 And bottomCounterCornice.Count > 3 And topCounterHandCornice.Count = bottomCounterCornice.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Карниз откосного крыла (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim modelCornice As DwgModel3DElement = New DwgModel3DElement()
                    Dim dataModelCornice As StructureElement = Nothing
                    If userHand.SideHand = Pillar.SidePillarElement.Left Then
                        dataModelCornice = CabinetWallModel.getModel(dictionaryObjectsBridge, userHand.NumberPillar, StructureElement.typeObject.modelLeftHandCornice, userHand.NumberSubPillar)
                    Else
                        dataModelCornice = CabinetWallModel.getModel(dictionaryObjectsBridge, userHand.NumberPillar, StructureElement.typeObject.modelRightHandCornice, userHand.NumberSubPillar)
                    End If
                    If IsNothing(dataModelCornice) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModelCornice.DWGEntity
                        modelCornice = dwgModel
                    Else
                        If userHand.SideHand = Pillar.SidePillarElement.Left Then
                            dataModelCornice = createModel(idBridge, StructureElement.typeObject.modelLeftHandCornice)
                        Else
                            dataModelCornice = createModel(idBridge, StructureElement.typeObject.modelRightHandCornice)
                        End If
                    End If
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterHandCornice, bottomCounterCornice)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim elementModel = New StaticSolidElement("Карниз откосного крыла (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = topCounterHandCornice.Item(0)
                            modelCornice.Position = elementModel.Origin
                            modelCornice.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelCornice) = False Then
                                activProjectDocument.ActiveSpace.Add(modelCornice)
                                Dim userModel As HandModel = New HandModel(userHand.NumberPillar, userHand.NumberSubPillar, userHand.SideHand)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModelCornice.KeyParameter = strGSON
                                dataModelCornice.DWGEntity = modelCornice
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelCornice, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModelCornice)
                                styleModel.setObjectStyle(modelCornice)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
End Class
