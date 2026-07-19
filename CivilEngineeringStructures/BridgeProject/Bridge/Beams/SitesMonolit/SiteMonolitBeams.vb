Imports System.ComponentModel
Imports Topomatic.Dwg.Entities
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Public Class SiteMonolitBeams
    Private _numberProlet As Integer           'Номер пролета
    Private _numberRow As Integer              'Номер ряда
    Private _numberLeftBeam As Integer         'номер левой балки
    Private _numberRightBeam As Integer        'номер правой балки
    Private _thickness As Double               'толщина
    Public _elementBridgePoint As PointsCollections
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
        _numberLeftBeam = 0
        _numberRightBeam = 0
        _thickness = 0
        _elementBridgePoint = New PointsCollections
    End Sub
    Public Sub New(numberProlet As Integer, numberRow As Integer, thickness As Double)
        _numberProlet = numberProlet
        _numberRow = numberRow
        _thickness = thickness
    End Sub
    ' Номер пролета
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
    <Browsable(True)>
    <Description("Номер левой балки")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер левой балки")>
    <[ReadOnly](True)>
    Public Property numberLeftBeam() As Integer
        Get
            Return _numberLeftBeam
        End Get
        Set(value As Integer)
            _numberLeftBeam = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер правой балки")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер правой балки")>
    <[ReadOnly](True)>
    Public Property numberRightBeam() As Integer
        Get
            Return _numberRightBeam
        End Get
        Set(value As Integer)
            _numberRightBeam = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Толщина участка омоноличивания, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Толщина")>
    <[ReadOnly](True)>
    Public Property thickness() As Double
        Get
            Return _thickness
        End Get
        Set(value As Double)
            _thickness = value
        End Set
    End Property
    Public Shared Function getAxisMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As DwgLine
        Dim result As DwgLine = New DwgLine
        Dim listAxisSiteMonolit As List(Of StructureElement) = New List(Of StructureElement)
        If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.axisSiteMonolitBeams) = True Then
            listAxisSiteMonolit = dictinaryAllObjectBridge.Item(StructureElement.typeObject.axisSiteMonolitBeams)
            If listAxisSiteMonolit.Count > 0 Then
                For i As Integer = 0 To listAxisSiteMonolit.Count - 1
                    Dim dataStructure As StructureElement = listAxisSiteMonolit.Item(i)
                    Dim tempSiteMonolit As SiteMonolitBeams = dataStructure.getMonolitSiteBeams
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
