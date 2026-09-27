Imports System.ComponentModel
Imports Topomatic.Dwg.Entities
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Public Class SiteMonolitBeams

    Private _numberProlet As Integer           'Номер пролета
    Private _numberRow As Integer              'Номер ряда
    Private _numberLeftBeam As Integer         'номер левой балки
    Private _numberRightBeam As Integer        'номер правой балки
    Private _thickness As Double               'толщина
    Private _fullMonolit As Boolean            '
    Public _elementBridgePoint As PointsCollections
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
        _numberLeftBeam = 0
        _numberRightBeam = 0
        _thickness = 0
        _fullMonolit = False
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
    Public Property thickness() As Double
        Get
            Return _thickness
        End Get
        Set(value As Double)
            _thickness = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Тип омоноличивания (перпендикуляр от края балки\полное - от края до края), м")>
    <Category("Свойства сооружения")>
    <DisplayName("Полное омоноличивание")>
    Public Property FullMonolit() As Boolean
        Get
            Return _fullMonolit
        End Get
        Set(value As Boolean)
            _fullMonolit = value
        End Set
    End Property
    'создать класс участок омоноличивания балок
    Public Shared Function createAxis(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.SitesBeamsMonolit
        elementAxis.Name = StructureElement.typeObject.axisSiteMonolitBeams
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.axisSiteMonolitBeams)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgLine
        Return elementAxis
    End Function
    'ищем ось участка омоноличивания
    Public Shared Function getAxisMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberLeftBeam As Integer, ByVal numberRightBeam As Integer) As StructureElement
        Dim result As StructureElement = Nothing
        If dictinaryAllObjectBridge.ContainsKey(StructureElement.typeObject.axisSiteMonolitBeams) = True Then
            Dim listAxisSiteMonolit As List(Of StructureElement) = dictinaryAllObjectBridge.Item(StructureElement.typeObject.axisSiteMonolitBeams)
            If listAxisSiteMonolit.Count > 0 Then
                For i As Integer = 0 To listAxisSiteMonolit.Count - 1
                    Dim dataStructure As StructureElement = listAxisSiteMonolit.Item(i)
                    Dim tempSiteMonolit As SiteMonolitBeams = dataStructure.getMonolitSiteBeams
                    If IsNothing(tempSiteMonolit) = False Then
                        If tempSiteMonolit.numberProlet = numberProlet Then
                            If tempSiteMonolit.numberLeftBeam = numberLeftBeam And tempSiteMonolit.numberRightBeam = numberRightBeam Then
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
