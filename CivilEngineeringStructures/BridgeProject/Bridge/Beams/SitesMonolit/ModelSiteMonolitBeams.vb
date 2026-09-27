Imports System.ComponentModel
Imports CivilEnginStructures.CounterBeam
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization.Runtime

Public Class ModelSiteMonolitBeams
    Private _numberProlet As Integer 'номер пролета
    Private _numberRow As Integer 'номер ряда
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
    End Sub
    Public Sub New(numberProlet As Integer, numberRow As Integer)
        _numberProlet = numberProlet
        _numberRow = numberRow
    End Sub
    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property numberProlet() As Integer
        Get
            Return _numberProlet
        End Get
        Set(value As Integer)
            _numberProlet = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property numberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    'создать класс участок омоноличивания балок
    Public Shared Function createModel(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.SitesBeamsMonolit
        elementAxis.Name = StructureElement.typeObject.modelSiteMonolitBeams
        Dim deskObject As String = StructureElement.GetDescription(elementAxis.Name)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgModel3DElement
        Return elementAxis
    End Function
    Public Shared Function deleteAllModelSitesMonolit(ByRef drawing As Drawing, ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Integer
        Dim countSiteMonolit As Integer = 0
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelSiteMonolitBeams) = True Then
            Dim listTopCountersBeam = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelSiteMonolitBeams)
            If IsNothing(listTopCountersBeam) = False Then
                If listTopCountersBeam.Count > 0 Then
                    For k As Integer = 0 To listTopCountersBeam.Count - 1
                        Dim tempData As StructureElement = listTopCountersBeam.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim modelEntity As DwgModel3DElement = tempData.DWGEntity
                            If IsNothing(modelEntity) = False Then
                                If drawing.ActiveSpace.Entities.Contains(modelEntity) = True Then
                                    drawing.ActiveSpace.Entities.Remove(modelEntity)
                                    countSiteMonolit += 1
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return countSiteMonolit
    End Function

    Public Shared Function getModelMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As StructureElement
        Dim result As StructureElement = Nothing
        If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.modelSiteMonolitBeams) = True Then
            Dim listModelSiteMonolit As List(Of StructureElement) = dictinaryAllObjectBridge.Item(StructureElement.typeObject.modelSiteMonolitBeams)
            If listModelSiteMonolit.Count > 0 Then
                For i As Integer = 0 To listModelSiteMonolit.Count - 1
                    Dim dataStructure As StructureElement = listModelSiteMonolit.Item(i)
                    Dim tempModelSiteMonolit As ModelSiteMonolitBeams = dataStructure.getModelSiteMonolitBeams
                    If IsNothing(tempModelSiteMonolit) = False Then
                        If tempModelSiteMonolit.numberProlet = numberProlet Then
                            If tempModelSiteMonolit.numberRow = numberRow Then
                                result = dataStructure
                                Exit For
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return result
    End Function


End Class
