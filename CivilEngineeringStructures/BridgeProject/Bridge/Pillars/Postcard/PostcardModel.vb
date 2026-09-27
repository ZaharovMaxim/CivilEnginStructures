Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class PostcardModel
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
        elementCounter.Name = type
        If type = StructureElement.typeObject.modelLeftPostcard Then
            elementCounter.ClassObject = StructureElement.classStructure.PostcardLeftPillar
        ElseIf type = StructureElement.typeObject.modelRightPostcard Then
            elementCounter.ClassObject = StructureElement.classStructure.PostcardRightPillar
        End If
        elementCounter.Description = StructureElement.GetDescription(type)
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
                            Dim userObject As PostcardModel = tempData.getPostcardModelPillar()
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

    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPostcard As PostcardPillar, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userPostcard) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterPostcard As DwgPolyline3D = Nothing
        Dim bottomCounterPostcard As DwgPolyline3D = Nothing
        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftPostcardTop) = True Then
                topCounterPostcard = dictPolyline.Item(StructureElement.typeObject.counterLeftPostcardTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterLeftPostcardBottom) = True Then
                bottomCounterPostcard = dictPolyline.Item(StructureElement.typeObject.counterLeftPostcardBottom)
            End If
        Else
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightPostcardTop) = True Then
                topCounterPostcard = dictPolyline.Item(StructureElement.typeObject.counterRightPostcardTop)
            End If
            If dictPolyline.ContainsKey(StructureElement.typeObject.counterRightPostcardBottom) = True Then
                bottomCounterPostcard = dictPolyline.Item(StructureElement.typeObject.counterRightPostcardBottom)
            End If
        End If
        If IsNothing(topCounterPostcard) = False Then
            If IsNothing(bottomCounterPostcard) = False Then
                If topCounterPostcard.Count > 3 And bottomCounterPostcard.Count > 3 And topCounterPostcard.Count = bottomCounterPostcard.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Откосное крыло (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim dataModel As StructureElement = Nothing
                    Dim modelPostcard As DwgModel3DElement = New DwgModel3DElement()
                    If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                        dataModel = PostcardModel.getModel(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.modelLeftPostcard, userPostcard.NumberSubPillar)
                    Else
                        dataModel = PostcardModel.getModel(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.modelRightPostcard, userPostcard.NumberSubPillar)
                    End If
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            modelPostcard = dwgModel
                        End If
                    Else
                        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                            dataModel = createModel(idBridge, StructureElement.typeObject.modelLeftPostcard)
                        Else
                            dataModel = createModel(idBridge, StructureElement.typeObject.modelRightPostcard)
                        End If
                    End If
                    Dim startPoint As Vector3D = userPostcard._elementBridgePoint.StartAxisPoint
                    Dim endPoint As Vector3D = userPostcard._elementBridgePoint.EndAxisPoint
                    Dim axisStartPoint As Vector3D = New Vector3D(startPoint.X, startPoint.Y, startPoint.Z)
                    Dim axisEndPoint As Vector3D = New Vector3D(endPoint.X, endPoint.Y, endPoint.Z)
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterPostcard, bottomCounterPostcard)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim centreAxisModel As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(axisStartPoint, axisEndPoint)
                            Dim elementModel = New StaticSolidElement("Откосное крыло (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = centreAxisModel
                            modelPostcard.Position = elementModel.Origin
                            modelPostcard.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(modelPostcard) = False Then
                                activProjectDocument.ActiveSpace.Add(modelPostcard)
                                Dim userModel As PostcardModel = New PostcardModel(userPostcard.NumberPillar, userPostcard.NumberSubPillar)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = modelPostcard
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelPostcard, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(modelPostcard)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
End Class
