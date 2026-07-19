Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Runtime
Public Class PileModel
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _numberColumn As Integer
    Private _numberRow As Integer
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberColumn = 0
        _numberRow = 0
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, NumberColumn As Integer, NumberRow As Integer)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _numberColumn = NumberColumn
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
    <Description("Номер столбца")>
    <Category("Свойства")>
    <DisplayName("Номер столбца")>
    <[ReadOnly](True)>
    Public Property NumberColumn() As Integer
        Get
            Return _numberColumn
        End Get
        Set(value As Integer)
            _numberColumn = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property NumberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    Public Shared Function createModel(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassObject = StructureElement.classStructure.PilePillar
        elementCounter.Name = type
        elementCounter.Description = "Свая (модель)"
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgModel3DElement
        Return elementCounter
    End Function

    'функция ищет существующий контур
    Public Shared Function getModel(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberSubPillar As Integer, ByVal numberColumn As Integer, ByVal numberRow As Integer, Optional removeDict As Boolean = False) As StructureElement
        Dim dataModel As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberPillar < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelPile) = True Then
            Dim listObject As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelPile)
            If IsNothing(listObject) = False Then
                If listObject.Count > 0 Then
                    For k As Integer = 0 To listObject.Count - 1
                        Dim tempData As StructureElement = listObject.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userObject As PileModel = tempData.getPileModelPillar()
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        If numberColumn = userObject.NumberColumn Then
                                            If numberRow = userObject.NumberRow Then
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
                        End If
                    Next k
                End If
            End If
        End If
        Return dataModel
    End Function

    'рисует Сваю (ТЛС объект)
    Public Shared Function drawModel(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPile As PilePillar, ByVal docPileTLC As ConstructionDocument, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Boolean
        If IsNothing(userPile) Then Return False
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Свая (модель)")

        '============================================================================================================================
        'находим старый контур по верху
        Dim dataModel As StructureElement = PileModel.getModel(dictionaryObjectsBridge, userPile.NumberPillar, userPile.NumberSubPillars, userPile.NumberColumn, userPile.NumberRow, True)
        If IsNothing(dataModel) = False Then
            Dim dwgModel As DwgModel3DElement = dataModel.DWGEntity
            If IsNothing(dwgModel) = False Then
                If activProjectDocument.ActiveSpace.Entities.Contains(dwgModel) = False Then
                    activProjectDocument.ActiveSpace.Entities.Remove(dwgModel)
                End If
            End If
        Else
            dataModel = createModel(idBridge, StructureElement.typeObject.modelPile)
        End If

        Dim centerPoint As Cad.Foundation.Vector3D = userPile._elementBridgePoint.CenterTopPoint
        Dim bottomPointPile As Topomatic.Cad.Foundation.Vector3D = userPile._elementBridgePoint.CenterBottomPoint
        Dim axisLinePile As DwgLine = New DwgLine()
        axisLinePile.StartPoint = centerPoint
        axisLinePile.EndPoint = bottomPointPile
        'вставляем модель
        Dim acModelPile As DwgModel3DElement = New DwgModel3DElement()
        Dim elDefault As ImProperties = New ImProperties()
        acModelPile.Position = New Cad.Foundation.Vector3D(centerPoint.Pos, centerPoint.Z + userPile.TopSeal)
        acModelPile.Rotation = userPile.Rotation
        activProjectDocument.ActiveSpace.Add(acModelPile)
        Dim elementPile As ConstructedModel3dElement = New ConstructedModel3dElement(docPileTLC, elDefault)
        acModelPile.Element = elementPile
        Dim elementPropertiesObject As ImElement = acModelPile.Element
        Dim generalPropertiesObject As ImProperties = elementPropertiesObject.GetProperties()
        If generalPropertiesObject.Count > 0 Then
            acModelPile.BeginUpdate()
            For k3 As Integer = 0 To generalPropertiesObject.Count - 1
                Dim generalPropLevel1 As ImProperty = generalPropertiesObject.Item(k3)
                If generalPropLevel1.Name Like "Тип" Then
                    generalPropLevel1.Value = userPile.Type
                    Exit For
                End If
            Next k3
            acModelPile.EndUpdate()
            acModelPile.BeginUpdate()
            generalPropertiesObject = elementPropertiesObject.GetProperties()
            For k3 As Integer = 0 To generalPropertiesObject.Count - 1
                Dim generalPropLevel1 As ImProperty = generalPropertiesObject.Item(k3)
                If generalPropLevel1.Name Like "Длина сваи" Then
                    generalPropLevel1.Value = userPile.Height + userPile.TopSeal
                ElseIf generalPropLevel1.Name Like "Высота заделки сваи" Then
                    generalPropLevel1.Value = userPile.TopSeal
                ElseIf generalPropLevel1.Name Like "Угол сваи по X" Then
                    generalPropLevel1.Value = userPile.AngleX
                ElseIf generalPropLevel1.Name Like "Угол сваи по Y" Then
                    generalPropLevel1.Value = userPile.AngleY
                ElseIf generalPropLevel1.Name Like "Сторона сваи, мм" Then
                    If elementPile.Name Like "Призматическая свая" Then
                        If userPile.Width = 0.3 Then
                            generalPropLevel1.Value = "300"
                        ElseIf userPile.Width = 0.35 Then
                            generalPropLevel1.Value = "350"
                        Else
                            generalPropLevel1.Value = "400"
                        End If
                    Else
                        generalPropLevel1.Value = userPile.Diameter
                    End If
                ElseIf generalPropLevel1.Name Like "Диаметр сваи" Then
                    If elementPile.Name Like "Буровая свая" Then
                        generalPropLevel1.Value = userPile.Diameter
                    End If
                ElseIf generalPropLevel1.Name Like "Наличие уширения" Then
                    If elementPile.Name Like "Буровая свая" Then
                        If userPile.Expand = True Then
                            generalPropLevel1.Value = "Да"
                        Else
                            generalPropLevel1.Value = "Нет"
                        End If
                    End If
                End If
            Next k3
            acModelPile.EndUpdate()
            If userPile.Expand = True Then
                generalPropertiesObject = elementPropertiesObject.GetProperties()
                For k3 As Integer = 0 To generalPropertiesObject.Count - 1
                    Dim generalPropLevel1 As ImProperty = generalPropertiesObject.Item(k3)
                    If generalPropLevel1.Name Like "Высота уширения по низу" Then
                        generalPropLevel1.Value = userPile.HeightExpand
                    ElseIf generalPropLevel1.Name Like "Диаметр уширения" Then
                        generalPropLevel1.Value = userPile.WidthExpand
                    ElseIf generalPropLevel1.Name Like "Высота от низа сваи до уширения" Then
                        generalPropLevel1.Value = userPile.HeightDownExpand
                    ElseIf generalPropLevel1.Name Like "Угол" Then
                        generalPropLevel1.Value = userPile.DegExpand
                    End If
                Next k3
            End If
        End If
        acModelPile.Prepare(activProjectDocument)
        If activProjectDocument.ActiveSpace.Entities.Contains(acModelPile) = True Then
            Dim userModel As PileModel = New PileModel(userPile.NumberPillar, userPile.NumberSubPillars, userPile.NumberColumn, userPile.NumberRow)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
            dataModel.KeyParameter = strGSON
            dataModel.DWGEntity = acModelPile
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(acModelPile, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
            styleModel.setObjectStyle(acModelPile)
        End If
        'Dim boolremoveModel As Boolean = Pillar.removeModel(activProjectDocument, dictionaryObjectsBridge)
        Return True
    End Function
    '
    Public Shared Function getConstructionElement(ByVal generalDirectory As String, ByVal nameModel As String) As ConstructionDocument
        Dim result As ConstructionDocument = New ConstructionDocument()
        If IO.Directory.Exists(generalDirectory) = False Then
            Dim directoryPileTLC As String = generalDirectory & "\TopomaticRobur\DesignBridge\Pillars\Сваи\"
            If Directory.Exists(directoryPileTLC) = True Then
                Dim arrayTLCFiles As String() = Directory.GetFiles(directoryPileTLC, "*.tlc")
                If IsArray(arrayTLCFiles) = True Then
                    For i As Integer = 0 To arrayTLCFiles.Length - 1
                        Dim fInfo As FileInfo = New FileInfo(arrayTLCFiles(i))
                        Dim nameFile As String = fInfo.Name
                        If nameFile Like nameModel Then
                            result.LoadFromFile(fInfo.FullName)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return result
    End Function

End Class
