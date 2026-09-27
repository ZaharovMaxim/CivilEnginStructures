Imports System.ComponentModel
Imports CivilEnginStructures.Pillar
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class PostcardContour
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер насадки (0 если насадка одна в опоре)
    Private _type As Pillar.SidePillarElement
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _type = Pillar.SidePillarElement.None
    End Sub
    Public Sub New(NumberPillar As Integer, NumberSubPillar As Integer, TypeCounter As Pillar.SidePillarElement)
        _numberPillar = NumberPillar
        _numberSubPillar = NumberSubPillar
        _type = TypeCounter
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
    Public Property SideContour() As Pillar.SidePillarElement
        Get
            Return _type
        End Get
        Set(value As Pillar.SidePillarElement)
            _type = value
        End Set
    End Property
    Public Shared Function createContour(ByVal idBridge As String, ByVal type As StructureElement.typeObject) As StructureElement
        Dim elementPillar As StructureElement = New StructureElement()
        elementPillar.Label = "Мосты и путепроводы"
        elementPillar.ClassBridgeObject = StructureElement.classBridge.Pillars
        elementPillar.Name = type
        If type = StructureElement.typeObject.counterLeftPostcardTop Then
            elementPillar.ClassObject = StructureElement.classStructure.PostcardLeftPillar
        ElseIf type = StructureElement.typeObject.counterLeftPostcardBottom Then
            elementPillar.ClassObject = StructureElement.classStructure.PostcardLeftPillar
        ElseIf type = StructureElement.typeObject.counterRightPostcardTop Then
            elementPillar.ClassObject = StructureElement.classStructure.PostcardRightPillar
        ElseIf type = StructureElement.typeObject.counterRightPostcardBottom Then
            elementPillar.ClassObject = StructureElement.classStructure.PostcardRightPillar
        End If
        elementPillar.Description = StructureElement.GetDescription(elementPillar.Name)
        elementPillar.KeyParameter = ""
        elementPillar.IdElement = Guid.NewGuid.ToString
        elementPillar.IdStructure = idBridge
        elementPillar.Note = ""
        elementPillar.DWGEntity = New DwgPolyline3D
        Return elementPillar
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет откосное крыло
    Public Shared Function getContour(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal typeContour As StructureElement.typeObject, Optional ByVal numberSubPillar As Integer = 0) As StructureElement
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
                            Dim userObject As PostcardContour = tempData.getPostcardCounterPillar
                            If IsNothing(userObject) = False Then
                                If numberPillar = userObject.NumberPillar Then
                                    If numberSubPillar = 0 Then numberSubPillar = userObject.NumberSubPillar
                                    If numberSubPillar = userObject.NumberSubPillar Then
                                        dataContour = tempData
                                        Exit For
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование контура откосного крыла
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userPostcard As PostcardPillar, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal templateXML As String) As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        Dim result As New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userPostcard) Then Return result
        Dim ListPointModel As Dictionary(Of Integer, PointStructure) = userPostcard._elementBridgePoint.ListPointModel
        If IsNothing(ListPointModel) = True Then Return result
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleCounter As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Обратный открылок (верх контура)")
        Dim listPointTopCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        Dim listPointBottomCounter As List(Of Cad.Foundation.Vector3D) = New List(Of Cad.Foundation.Vector3D)
        If ListPointModel.Count > 3 Then
            For i As Integer = 0 To ListPointModel.Count - 1
                Dim pointBridge As PointStructure = ListPointModel.ElementAt(i).Value
                Dim topPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X, pointBridge.Y, pointBridge.Z)
                Dim bottomPoint As Vector3D = New Cad.Foundation.Vector3D(pointBridge.X - pointBridge.dx, pointBridge.Y - pointBridge.dy, pointBridge.Z + pointBridge.dz)
                listPointTopCounter.Add(topPoint)
                listPointBottomCounter.Add(bottomPoint)
            Next i
        End If
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataTopCounter As StructureElement = Nothing
        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
            dataTopCounter = PostcardContour.getContour(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.counterLeftPostcardTop, userPostcard.NumberSubPillar)
        Else
            dataTopCounter = PostcardContour.getContour(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.counterRightPostcardTop, userPostcard.NumberSubPillar)
        End If
        Dim poly3dCounterTop As DwgPolyline3D = Nothing
        If IsNothing(dataTopCounter) = True Then
            If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                dataTopCounter = PostcardContour.createContour(idBridge, StructureElement.typeObject.counterLeftPostcardTop)
            Else
                dataTopCounter = PostcardContour.createContour(idBridge, StructureElement.typeObject.counterRightPostcardTop)
            End If
            poly3dCounterTop = dataTopCounter.DWGEntity
        Else
            poly3dCounterTop = dataTopCounter.DWGEntity
        End If
        Dim drawClass As CreateDwgObject = New CreateDwgObject(activProjectDocument)
        If poly3dCounterTop.Count = 0 Then
            poly3dCounterTop = drawClass.createPolyline3D(listPointTopCounter, True)
            Dim userCounter As PostcardContour = New PostcardContour(userPostcard.NumberPillar, userPostcard.NumberSubPillar, SidePillarElement.None)
            If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                userCounter.SideContour = SidePillarElement.Left
            ElseIf userPostcard.SidePostcard = Pillar.SidePillarElement.Right Then
                userCounter.SideContour = SidePillarElement.Right
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
        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
            result.Add(StructureElement.typeObject.counterLeftPostcardTop, poly3dCounterTop)
        ElseIf userPostcard.SidePostcard = Pillar.SidePillarElement.Right Then
            result.Add(StructureElement.typeObject.counterRightPostcardTop, poly3dCounterTop)
        End If
        '============================================================================================================================
        'находим старый контур по низу
        styleCounter.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Обратный открылок (низ контура)")
        '============================================================================================================================
        'находим старый контур по верху
        Dim dataBottomCounter As StructureElement = Nothing
        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
            dataBottomCounter = PostcardContour.getContour(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.counterLeftPostcardBottom, userPostcard.NumberSubPillar)
        Else
            dataBottomCounter = PostcardContour.getContour(dictionaryObjectsBridge, userPostcard.NumberPillar, StructureElement.typeObject.counterRightPostcardBottom, userPostcard.NumberSubPillar)
        End If
        Dim poly3dCounterBottom As DwgPolyline3D = Nothing
        If IsNothing(dataBottomCounter) = True Then
            If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                dataBottomCounter = PostcardContour.createContour(idBridge, StructureElement.typeObject.counterLeftPostcardBottom)
            Else
                dataBottomCounter = PostcardContour.createContour(idBridge, StructureElement.typeObject.counterRightPostcardBottom)
            End If
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        Else
            poly3dCounterBottom = dataBottomCounter.DWGEntity
        End If
        If poly3dCounterBottom.Count = 0 Then
            poly3dCounterBottom = drawClass.createPolyline3D(listPointBottomCounter, True)
            Dim userCounter As PostcardContour = New PostcardContour(userPostcard.NumberPillar, userPostcard.NumberSubPillar, SidePillarElement.None)
            If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
                userCounter.SideContour = SidePillarElement.Left
            ElseIf userPostcard.SidePostcard = Pillar.SidePillarElement.Right Then
                userCounter.SideContour = SidePillarElement.Right
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
        Dim boolRecData2 As Boolean = FuncXRecords.setXRecords(poly3dCounterBottom, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataBottomCounter)
        If userPostcard.SidePostcard = Pillar.SidePillarElement.Left Then
            result.Add(StructureElement.typeObject.counterLeftPostcardBottom, poly3dCounterBottom)
        ElseIf userPostcard.SidePostcard = Pillar.SidePillarElement.Right Then
            result.Add(StructureElement.typeObject.counterRightPostcardBottom, poly3dCounterBottom)
        End If
        Return result
    End Function
End Class
