Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class CabinetWallModel
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
    Public Shared Function createModel(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassObject = StructureElement.classStructure.CabinetWallPillar
        elementCounter.Name = type
        If type = StructureElement.typeObject.modelCabinetWall Then
            elementCounter.Description = "Шкафная стенка (модель)"
        ElseIf type = StructureElement.typeObject.modelCabinetWallPlate Then
            elementCounter.Description = "Зуб упора (модель)"
        End If
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgPolyline3D
        Return elementCounter
    End Function

    'функция ищет существующий контур шкафной стенки
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
                            Dim userObject As CabinetWallModel = tempData.getCabinetWallModelPillar
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

    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userCabinetWall As CabinetWallPillar, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userCabinetWall) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterCabinetWall As DwgPolyline3D = Nothing
        Dim bottomCounterCabinetWall As DwgPolyline3D = Nothing
        Dim topCounterCabinetWallPlate As DwgPolyline3D = Nothing
        Dim bottomCounterCabinetWallPlate As DwgPolyline3D = Nothing
        If dictPolyline.ContainsKey(StructureElement.typeObject.contourCabinetWallTop) = True Then
            topCounterCabinetWall = dictPolyline.Item(StructureElement.typeObject.contourCabinetWallTop)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.contourCabinetWallBottom) = True Then
            bottomCounterCabinetWall = dictPolyline.Item(StructureElement.typeObject.contourCabinetWallBottom)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.contourCabinetWallPlateTop) = True Then
            topCounterCabinetWallPlate = dictPolyline.Item(StructureElement.typeObject.contourCabinetWallPlateTop)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.contourCabinetWallPlateBottom) = True Then
            bottomCounterCabinetWallPlate = dictPolyline.Item(StructureElement.typeObject.contourCabinetWallPlateBottom)
        End If
        If IsNothing(topCounterCabinetWall) = False Then
            If IsNothing(bottomCounterCabinetWall) = False Then
                If topCounterCabinetWall.Count > 3 And bottomCounterCabinetWall.Count > 3 And topCounterCabinetWall.Count = bottomCounterCabinetWall.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Шкафная стенка (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim modelCabinetWall As DwgModel3DElement = New DwgModel3DElement()
                    Dim dataModel As StructureElement = CabinetWallModel.getModel(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.modelCabinetWall, userCabinetWall.Number)
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            modelCabinetWall = dwgModel
                        End If
                    Else
                        dataModel = createModel(idBridge, StructureElement.typeObject.modelCabinetWall)
                    End If
                    Dim axisStartPoint As Vector3D = userCabinetWall._elementBridgePoint.StartAxisPoint
                    Dim axisEndPoint As Vector3D = userCabinetWall._elementBridgePoint.EndAxisPoint
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterCabinetWall, bottomCounterCabinetWall)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim centreAxisModel As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(axisStartPoint, axisEndPoint)
                            Dim elementModel = New StaticSolidElement("Шкафная стенка (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = userCabinetWall._elementBridgePoint.CenterTopPoint
                            modelCabinetWall.Position = elementModel.Origin
                            modelCabinetWall.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelCabinetWall) = False Then
                                activProjectDocument.ActiveSpace.Add(modelCabinetWall)
                                Dim userModel As CabinetWallModel = New CabinetWallModel(userCabinetWall.NumberPillar, userCabinetWall.Number)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = modelCabinetWall
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelCabinetWall, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(modelCabinetWall)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If IsNothing(topCounterCabinetWallPlate) = False Then
            If IsNothing(bottomCounterCabinetWallPlate) = False Then
                If topCounterCabinetWallPlate.Count > 3 And bottomCounterCabinetWallPlate.Count > 3 And topCounterCabinetWallPlate.Count = bottomCounterCabinetWallPlate.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Зуб упора (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim modelCabinetWall As DwgModel3DElement = New DwgModel3DElement()
                    Dim dataModel As StructureElement = CabinetWallModel.getModel(dictionaryObjectsBridge, userCabinetWall.NumberPillar, StructureElement.typeObject.modelCabinetWallPlate, userCabinetWall.Number)
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            modelCabinetWall = dwgModel
                        End If
                    Else
                        dataModel = createModel(idBridge, StructureElement.typeObject.modelCabinetWallPlate)
                    End If
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterCabinetWallPlate, bottomCounterCabinetWallPlate)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim elementModel = New StaticSolidElement("Зуб упора (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = topCounterCabinetWallPlate.Item(0)
                            modelCabinetWall.Position = elementModel.Origin
                            modelCabinetWall.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelCabinetWall) = False Then
                                activProjectDocument.ActiveSpace.Add(modelCabinetWall)
                                Dim userModel As CabinetWallModel = New CabinetWallModel(userCabinetWall.NumberPillar, userCabinetWall.Number)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = modelCabinetWall
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelCabinetWall, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(modelCabinetWall)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
End Class
