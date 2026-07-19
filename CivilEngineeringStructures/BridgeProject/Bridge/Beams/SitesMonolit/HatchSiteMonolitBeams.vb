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

    Public Shared Function getHatchMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As DwgHatch
        Dim result As DwgHatch = New DwgHatch
        Dim listHatchMonolitSitesBeam As List(Of StructureElement) = New List(Of StructureElement)
        If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.hatchSiteMonolitPillar) = True Then
            listHatchMonolitSitesBeam = dictinaryAllObjectBridge.Item(StructureElement.typeObject.hatchSiteMonolitPillar)
            If listHatchMonolitSitesBeam.Count > 0 Then
                For i As Integer = 0 To listHatchMonolitSitesBeam.Count - 1
                    Dim dataStructure As StructureElement = listHatchMonolitSitesBeam.Item(i)
                    Dim tempSiteMonolit As HatchSiteMonolitBeams = dataStructure.getHatchMonolitSiteBeams
                    If IsNothing(tempSiteMonolit) = False Then
                        If tempSiteMonolit._numberProlet = numberProlet Then
                            If tempSiteMonolit.numberRow = numberRow Then
                                result = dataStructure.DWGEntity
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
