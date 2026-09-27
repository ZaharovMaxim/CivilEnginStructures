Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class PreparationModel
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
        elementCounter.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementCounter.ClassObject = StructureElement.classStructure.PreparationPillar
        elementCounter.Name = type
        elementCounter.Description = "Подготовка (модель)"
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgModel3DElement
        Return elementCounter
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
                            Dim userObject As PreparationModel = tempData.getPreparationModelPillar()
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

    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPreparation As PreparationPillar, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userPreparation) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterPreparation As DwgPolyline3D = Nothing
        Dim bottomCounterPreparation As DwgPolyline3D = Nothing

        If dictPolyline.ContainsKey(StructureElement.typeObject.counterPreparationTop) = True Then
            topCounterPreparation = dictPolyline.Item(StructureElement.typeObject.counterPreparationTop)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.counterPreparationBottom) = True Then
            bottomCounterPreparation = dictPolyline.Item(StructureElement.typeObject.counterPreparationBottom)
        End If
        If IsNothing(topCounterPreparation) = False Then
            If IsNothing(bottomCounterPreparation) = False Then
                If topCounterPreparation.Count > 3 And bottomCounterPreparation.Count > 3 And topCounterPreparation.Count = bottomCounterPreparation.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Подготовка (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim modelPreparation As DwgModel3DElement = New DwgModel3DElement()
                    Dim dataModel As StructureElement = GrillageModel.getModel(dictionaryObjectsBridge, userPreparation.NumberPillar, StructureElement.typeObject.modelPreparation, userPreparation.Number)
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            modelPreparation = dwgModel
                        End If
                    Else
                        dataModel = createModel(idBridge, StructureElement.typeObject.modelPreparation)
                    End If
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterPreparation, bottomCounterPreparation)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim centreAxisModel As Cad.Foundation.Vector3D = userPreparation._elementBridgePoint.StartAxisPoint
                            Dim elementModel = New StaticSolidElement("Подготовка (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = centreAxisModel
                            modelPreparation.Position = elementModel.Origin
                            modelPreparation.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelPreparation) = False Then
                                activProjectDocument.ActiveSpace.Add(modelPreparation)
                                Dim userModel As PreparationModel = New PreparationModel(userPreparation.NumberPillar, userPreparation.Number)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = modelPreparation
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelPreparation, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(modelPreparation)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
End Class
