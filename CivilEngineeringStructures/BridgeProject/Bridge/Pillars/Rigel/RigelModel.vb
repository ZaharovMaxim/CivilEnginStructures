Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Alg.Bridges
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class RigelModel
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
    'создать новую пустую насадку
    Public Shared Function createModel(ByVal idBridge As String) As StructureElement
        Dim elementNozzlePillar As StructureElement = New StructureElement()
        elementNozzlePillar.Label = "Мосты и путепроводы"
        elementNozzlePillar.ClassObject = StructureElement.classStructure.RigelPillar
        elementNozzlePillar.Name = StructureElement.typeObject.modelRigel
        elementNozzlePillar.Description = "Ригель (модель)"
        elementNozzlePillar.KeyParameter = ""
        elementNozzlePillar.IdElement = Guid.NewGuid.ToString
        elementNozzlePillar.IdStructure = idBridge
        elementNozzlePillar.Note = ""
        elementNozzlePillar.DWGEntity = New DwgModel3DElement
        Return elementNozzlePillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий ригель
    Public Shared Function getModel(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataModel As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelRigel) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelRigel)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As RigelModel = tempData.getRigelModelPillar
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование контура ригеля
    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userRigel As RigelPillar, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userRigel) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterRigel As DwgPolyline3D = Nothing
        Dim bottomCounterRigel As DwgPolyline3D = Nothing
        Dim centerLine As DwgLine = Nothing
        If dictPolyline.ContainsKey(StructureElement.typeObject.counterRigelTop) = True Then
            topCounterRigel = dictPolyline.Item(StructureElement.typeObject.counterRigelTop)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.counterRigelBottom) = True Then
            bottomCounterRigel = dictPolyline.Item(StructureElement.typeObject.counterRigelBottom)
        End If
        If IsNothing(userRigel._elementBridgePoint.StartAxisPoint) = False Then
            If IsNothing(userRigel._elementBridgePoint.EndAxisPoint) = False Then
                centerLine = New DwgLine()
                centerLine.StartPoint = userRigel._elementBridgePoint.StartAxisPoint
                centerLine.EndPoint = userRigel._elementBridgePoint.EndAxisPoint
            End If
        End If
        If IsNothing(topCounterRigel) = False Then
            If IsNothing(bottomCounterRigel) = False Then
                If topCounterRigel.Count > 3 And bottomCounterRigel.Count > 3 And topCounterRigel.Count = bottomCounterRigel.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Ригель (модель)")
                    '============================================================================================================================
                    'находим старый контур по верху
                    Dim dataModel As StructureElement = RigelModel.getModel(dictionaryObjectsBridge, userRigel.NumberPillar, userRigel.Number)
                    If IsNothing(dataModel) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            If activProjectDocument.ActiveSpace.Entities.Contains(dwgModel) = False Then
                                activProjectDocument.ActiveSpace.Entities.Remove(dwgModel)
                            End If
                        End If
                    Else
                        dataModel = createModel(idBridge)
                    End If
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterRigel, bottomCounterRigel, centerLine)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim elementModel = New StaticSolidElement("Ригель (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = userRigel._elementBridgePoint.StartAxisPoint
                            Dim model As DwgModel3DElement = New DwgModel3DElement()
                            model.Position = elementModel.Origin
                            model.Element = elementModel
                            activProjectDocument.ActiveSpace.Add(model)
                            If activProjectDocument.ActiveSpace.Entities.Contains(model) = True Then
                                Dim userModel As RigelModel = New RigelModel(userRigel.NumberPillar, userRigel.Number)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModel.KeyParameter = strGSON
                                dataModel.DWGEntity = model
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(model, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                styleModel.setObjectStyle(model)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
End Class
