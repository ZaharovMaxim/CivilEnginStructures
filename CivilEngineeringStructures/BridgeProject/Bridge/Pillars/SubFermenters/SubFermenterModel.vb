Imports System.ComponentModel
Imports CivilEnginStructures.SubFermenterContour
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class SubFermenterModel
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberProlet As Integer
    Private _numberRow As Integer
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberProlet = 0
        _numberRow = 0
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberColumn As Integer, NumberRow As Integer)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberProlet = NumberColumn
        _numberRow = NumberRow
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
    <Description("Номер пролета")>
    <Category("Свойства")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property NumberProlet() As Integer
        Get
            Return _numberProlet
        End Get
        Set(value As Integer)
            _numberProlet = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер ряда балок")>
    <Category("Свойства")>
    <DisplayName("Номер ряда балок")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    'создать новый пустой подферменник
    Public Shared Function createModel(ByVal idBridge As String) As StructureElement
        Dim elementNozzlePillar As StructureElement = New StructureElement()
        elementNozzlePillar.Label = "Мосты и путепроводы"
        elementNozzlePillar.ClassObject = StructureElement.classStructure.SubFermenters
        elementNozzlePillar.Name = StructureElement.typeObject.modelSubFermenters
        elementNozzlePillar.Description = "Подферменник (модель)"
        elementNozzlePillar.KeyParameter = ""
        elementNozzlePillar.IdElement = Guid.NewGuid.ToString
        elementNozzlePillar.IdStructure = idBridge
        elementNozzlePillar.Note = ""
        elementNozzlePillar.DWGEntity = New DwgModel3DElement
        Return elementNozzlePillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующий подферменник
    Public Shared Function getModel(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberRow As Integer, ByVal numberProlet As Integer, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataModel As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberProlet < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelSubFermenters) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelSubFermenters)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As SubFermenterModel = tempData.getSubFermentersModel()
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If numberProlet = userObject.NumberProlet Then
                                            If numberRow = userObject.NumberRow Then
                                                dataModel = tempData
                                                Exit For
                                            End If
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
    'функция создает модель
    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userSubFerm As SubFermenters, ByVal dictPolyline As Dictionary(Of StructureElement.typeObject, DwgPolyline3D), ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userSubFerm) Then Return False
        If IsNothing(dictPolyline) = True Then Return False
        Dim topCounterSubFerm As DwgPolyline3D = Nothing
        Dim bottomCounterSubFerm As DwgPolyline3D = Nothing
        If dictPolyline.ContainsKey(StructureElement.typeObject.counterSubFermentersTop) = True Then
            topCounterSubFerm = dictPolyline.Item(StructureElement.typeObject.counterSubFermentersTop)
        End If
        If dictPolyline.ContainsKey(StructureElement.typeObject.counterSubFermentersBottom) = True Then
            bottomCounterSubFerm = dictPolyline.Item(StructureElement.typeObject.counterSubFermentersBottom)
        End If
        If IsNothing(topCounterSubFerm) = False Then
            If IsNothing(bottomCounterSubFerm) = False Then
                If topCounterSubFerm.Count > 3 And bottomCounterSubFerm.Count > 3 And topCounterSubFerm.Count = bottomCounterSubFerm.Count Then
                    'стиль
                    Dim categoryTables As String = "Искусственные сооружения"
                    Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
                    styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Откосное крыло (модель)")
                    '============================================================================================================================
                    'находим старую модель и ее удаляем
                    Dim model As DwgModel3DElement = New DwgModel3DElement()
                    Dim dataModelSubFerm As StructureElement = getModel(dictionaryObjectsBridge, userSubFerm.NumberPillar, userSubFerm.NumberRow, userSubFerm.NumberProlet, userSubFerm.NumberSubPillar)
                    If IsNothing(dataModelSubFerm) = False Then
                        Dim dwgModel As DwgModel3DElement = dataModelSubFerm.DWGEntity
                        If IsNothing(dwgModel) = False Then
                            model = dwgModel
                        End If
                    End If
                    Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
                    Dim shellModel As Shell = drawClass.createSolid3DByTwoPolylines3d(topCounterSubFerm, bottomCounterSubFerm)
                    If IsNothing(shellModel) = False Then
                        If shellModel.Vertices.Count > 3 Then
                            Dim elementModel = New StaticSolidElement("Подферменник (модель)", "SmdxElement", New ImProperties(), shellModel, New ImDocuments())
                            elementModel.Origin = userSubFerm._elementBridgePoint.CenterTopPoint
                            model.Position = elementModel.Origin
                            model.Element = elementModel
                            If activProjectDocument.ActiveSpace.Entities.Contains(model) = False Then
                                activProjectDocument.ActiveSpace.Add(model)
                                dataModelSubFerm = createModel(idBridge)
                                Dim userModel As SubFermenterModel = New SubFermenterModel(userSubFerm.NumberPillar, userSubFerm.NumberSubPillar, userSubFerm.NumberProlet, userSubFerm.NumberRow)
                                Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                                dataModelSubFerm.KeyParameter = strGSON
                                dataModelSubFerm.DWGEntity = model
                                Dim boolRecData As Boolean = FuncXRecords.setXRecords(model, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModelSubFerm)
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
