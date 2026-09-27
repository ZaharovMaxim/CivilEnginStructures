Imports System.ComponentModel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Public Class NozzleModel
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
        elementNozzlePillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementNozzlePillar.ClassObject = StructureElement.classStructure.NozzlePillar
        elementNozzlePillar.Name = StructureElement.typeObject.modelNozzle
        elementNozzlePillar.Description = "Насадка (модель)"
        elementNozzlePillar.KeyParameter = ""
        elementNozzlePillar.IdElement = Guid.NewGuid.ToString
        elementNozzlePillar.IdStructure = idBridge
        elementNozzlePillar.Note = ""
        elementNozzlePillar.DWGEntity = New DwgModel3DElement
        Return elementNozzlePillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет существующую насадку
    Public Shared Function getModelNozzle(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
        Dim dataModelNozzle As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelNozzle) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelNozzle)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As NozzleModel = tempData.getNozzleModelPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        dataModelNozzle = tempData
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataModelNozzle
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование модели насадки
    Public Shared Function drawModelNozzle(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal topPolyline As DwgPolyline3D, ByVal bottomPolyline As DwgPolyline3D, ByVal lineCabinetWall As DwgLine, ByVal dataNozzle As StructureElement, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As DwgModel3DElement
        Dim result As DwgModel3DElement = New DwgModel3DElement()
        If IsNothing(dataNozzle) Then Return result
        Dim idBridge As String = dataNozzle.IdStructure
        If IsNothing(idBridge) = True Then Return result
        If idBridge.Trim.Length = 0 Then Return result
        If IsNothing(topPolyline) = True Then Return result
        If topPolyline.Count < 4 Then
            Return result
        End If
        If IsNothing(bottomPolyline) = True Then Return result
        If bottomPolyline.Count < 4 Then
            Return result
        End If
        If topPolyline.Count <> bottomPolyline.Count Then
            Return result
        End If
        If IsNothing(lineCabinetWall) = True Then
            Return result
        End If
        If lineCabinetWall.Length = 0 Then
            Return result
        End If
        Dim userNozzle As NozzlePillar = dataNozzle.getNozzlePillar()
        Dim axisStartPoint As Vector3D = userNozzle._elementBridgePoint.StartAxisPoint
        Dim axisEndPoint As Vector3D = userNozzle._elementBridgePoint.EndAxisPoint
        'модель
        Dim layerModelNozzle As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorModelNozzle As CadColor = New CadColor(7)
        Dim nameTypeLineModelNozzle As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineModelNozzle As Integer = 1
        Dim widthTypeLineModelNozzle As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleModelNozzle As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleModelNozzle.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Насадка (модель)")
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataModel As StructureElement = NozzleModel.getModelNozzle(dictionaryObjectsBridge, userNozzle.NumberPillar, userNozzle.Number)
        If IsNothing(dataModel) = False Then
            Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
            If IsNothing(dwgModel) = False Then
                result = dwgModel
            End If
        Else
            dataModel = createModel(idBridge)
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If topPolyline.Count > 3 And bottomPolyline.Count > 3 Then
            Dim shellNozzle As Shell = drawClass.createSolid3DByTwoPolylines3d(topPolyline, bottomPolyline, lineCabinetWall)
            If IsNothing(shellNozzle) = False Then
                If shellNozzle.Vertices.Count > 3 Then
                    Dim centreAxisNozzle As Cad.Foundation.Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(axisStartPoint, axisEndPoint)
                    Dim elementModelNozzle = New StaticSolidElement("Насадка (модель)", "SmdxElement", New ImProperties(), shellNozzle, New ImDocuments())
                    elementModelNozzle.Origin = centreAxisNozzle
                    result.Position = elementModelNozzle.Origin
                    result.Element = elementModelNozzle
                    If activProjectDocument.ActiveSpace.Entities.Contains(result) = False Then
                        activProjectDocument.ActiveSpace.Add(result)
                        Dim userModel As NozzleModel = New NozzleModel(userNozzle.NumberPillar, userNozzle.Number)
                        Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                        dataModel.KeyParameter = strGSON
                        dataModel.DWGEntity = result
                        Dim boolRecData As Boolean = FuncXRecords.setXRecords(result, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                        styleModelNozzle.setObjectStyle(result)
                    End If
                End If
            End If
        End If
        Return result
    End Function
End Class
