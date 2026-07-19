Imports System.ComponentModel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports CivilEnginStructures.PileContour
Imports CivilEnginStructures.RackContour
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class RackModel
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberRack As Integer 'номер стойки
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberRack = 0
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberRack As Integer)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberRack = NumberRack
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
    <Description("Номер стойки")>
    <Category("Свойства")>
    <DisplayName("Номер стойки")>
    <[ReadOnly](True)>
    Public Property NumberRack() As Integer
        Get
            Return _numberRack
        End Get
        Set(value As Integer)
            _numberRack = value
        End Set
    End Property

    Public Shared Function createModel(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassObject = StructureElement.classStructure.RackPillar
        elementCounter.Name = type
        elementCounter.Description = "Стойка (модель)"
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgModel3DElement
        Return elementCounter
    End Function

    'функция ищет существующий контур
    Public Shared Function getModel(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberRack As Integer, Optional ByVal numberSubPillar As Integer = 0, Optional removeDict As Boolean = False) As StructureElement
        Dim dataModel As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelRack) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelRack)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As RackModel = tempData.getRackModelPillar()
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If numberRack = userObject.NumberRack Then
                                            dataModel = tempData
                                            If removeDict = True Then
                                                listObject.Item(k) = Nothing
                                            End If
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
        Return dataModel
    End Function

    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userRack As RackPillar, ByVal dictEntity As Dictionary(Of StructureElement.typeObject, DwgEntity), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String, Optional removeDict As Boolean = True) As Boolean
        If IsNothing(userRack) Then Return False
        If IsNothing(dictEntity) = True Then Return False
        If dictEntity.Count = 0 Then Return False
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Стойка (модель)")

        Dim numberRack As Integer = userRack.Number
        'рисуем модель
        Dim topElement As DwgEntity = Nothing
        If dictEntity.ContainsKey(StructureElement.typeObject.counterRackTop) = True Then
            topElement = dictEntity.Item(StructureElement.typeObject.counterRackTop)
        End If
        Dim BottomElement As DwgEntity = Nothing
        If dictEntity.ContainsKey(StructureElement.typeObject.counterRackBottom) = True Then
            BottomElement = dictEntity.Item(StructureElement.typeObject.counterRackBottom)
        End If
        If IsNothing(topElement) = True Then Return False
        If IsNothing(BottomElement) = True Then Return False

        '============================================================================================================================
        'находим существующую модель
        Dim dataModel As StructureElement = RackModel.getModel(dictionaryObjectsBridge, userRack.NumberPillar, userRack.Number, userRack.NumberSubPillars, removeDict)
        Dim modelRack As DwgModel3DElement = New DwgModel3DElement()
        If IsNothing(dataModel) = False Then
            Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
            If IsNothing(dwgModel) = False Then
                modelRack = dwgModel
            Else
                dataModel = createModel(idBridge, StructureElement.typeObject.modelRack)
            End If
        Else
            dataModel = createModel(idBridge, StructureElement.typeObject.modelRack)
        End If
        'рисуем контура
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If userRack.RackType = RackPillar.TypeRack.Trapezoidal Then
            Dim topCounterRack As DwgPolyline3D = topElement
            Dim bottomCounterRack As DwgPolyline3D = BottomElement
            If IsNothing(topCounterRack) = False Then
                If IsNothing(bottomCounterRack) = False Then
                    If topCounterRack.Count > 3 And bottomCounterRack.Count > 3 And topCounterRack.Count = bottomCounterRack.Count Then
                        Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterRack, bottomCounterRack)
                        If IsNothing(shellModel) = False Then
                            If shellModel.Vertices.Count > 3 Then
                                Dim elementModel = New StaticSolidElement("Стойка (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                                elementModel.Origin = userRack._elementBridgePoint.StartAxisPoint
                                modelRack.Position = elementModel.Origin
                                modelRack.Element = elementModel
                                If activProjectDocument.ActiveSpace.Entities.Contains(modelRack) = False Then
                                    activProjectDocument.ActiveSpace.Add(modelRack)
                                    Dim userModel As RackModel = New RackModel(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number)
                                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                    dataModel.KeyParameter = strGSON
                                    dataModel.DWGEntity = modelRack
                                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelRack, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                                    styleModel.setObjectStyle(modelRack)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            'рисуем круглую стойку
            Dim shellRack As Shell = Nothing
            If userRack.RackType = RackPillar.TypeRack.Octagonal Then
                shellRack = Tools.Cylinder(userRack.Diameter / 2, userRack.Diameter / 2, userRack.Height + userRack.TopSeal, 8)
            Else
                shellRack = Tools.Cylinder(userRack.Diameter / 2, userRack.Diameter / 2, userRack.Height + userRack.TopSeal, 20)
            End If
            If IsNothing(shellRack) = False Then
                Dim elementModelRack = New StaticSolidElement("Стойка (модель)", "SmdxElement", New ImProperties(), shellRack, New ImDocuments())
                'elementModelRack.Origin = userRack._elementBridgePoint.StartAxisPoint
                modelRack.Position = userRack._elementBridgePoint.CenterBottomPoint
                modelRack.Element = elementModelRack
                If activProjectDocument.ActiveSpace.Entities.Contains(modelRack) = False Then
                    activProjectDocument.ActiveSpace.Add(modelRack)
                    Dim userModel As RackModel = New RackModel(userRack.NumberPillar, userRack.NumberSubPillars, userRack.Number)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                    dataModel.KeyParameter = strGSON
                    dataModel.DWGEntity = modelRack
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelRack, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                    styleModel.setObjectStyle(modelRack)
                End If
            End If
        End If
        If removeDict = True Then
            'Dim boolremoveModel As Boolean = Pillar.removeModel(activProjectDocument, dictionaryObjectsBridge)
        End If
        Return True
    End Function
End Class
