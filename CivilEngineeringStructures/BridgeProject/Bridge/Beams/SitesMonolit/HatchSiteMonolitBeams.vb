Imports System.ComponentModel
Imports CivilEnginStructures.CounterSiteMonolitBeams
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization.Runtime
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Public Class HatchSiteMonolitBeams
    Private _numberProlet As Integer 'номер пролета
    Private _numberRow As Integer 'номер ряда
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
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
    Public Shared Function createAxis(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.SitesBeamsMonolit
        elementAxis.Name = StructureElement.typeObject.hatchSiteMonolitBeams
        Dim deskObject As String = StructureElement.GetDescription(elementAxis.Name)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgHatch
        Return elementAxis
    End Function
    Public Shared Function getHatchMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As StructureElement
        Dim result As StructureElement = Nothing
        If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.hatchSiteMonolitBeams) = True Then
            Dim listHatchMonolitSitesBeam As List(Of StructureElement) = dictinaryAllObjectBridge.Item(StructureElement.typeObject.hatchSiteMonolitBeams)
            If listHatchMonolitSitesBeam.Count > 0 Then
                For i As Integer = 0 To listHatchMonolitSitesBeam.Count - 1
                    Dim dataStructure As StructureElement = listHatchMonolitSitesBeam.Item(i)
                    Dim tempSiteMonolit As HatchSiteMonolitBeams = dataStructure.getHatchMonolitSiteBeams
                    If IsNothing(tempSiteMonolit) = False Then
                        If tempSiteMonolit._numberProlet = numberProlet Then
                            If tempSiteMonolit.numberRow = numberRow Then
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
