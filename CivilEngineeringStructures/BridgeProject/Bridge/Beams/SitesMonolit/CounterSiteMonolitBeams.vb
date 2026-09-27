Imports System.ComponentModel
Imports Topomatic.Dwg.Entities
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Public Class CounterSiteMonolitBeams
    Private _numberProlet As Integer 'номер пролета
    Private _numberRow As Integer 'номер ряда
    Private _type As StructureElement.typeObject 'тип (низ балки, верх плиты балки)
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
        _type = StructureElement.typeObject.OtherElement
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
    <Browsable(True)>
    <Description("Тип контура")>
    <Category("Свойства сооружения")>
    <DisplayName("Тип контура")>
    <[ReadOnly](True)>
    Public Property TypeCounter() As StructureElement.typeObject
        Get
            Return _type
        End Get
        Set(value As StructureElement.typeObject)
            _type = value
        End Set
    End Property

    'создать класс участок омоноличивания балок
    Public Shared Function createAxis(ByVal idBridge As String, ByVal typeCounter As StructureElement.typeObject) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.SitesBeamsMonolit
        If typeCounter = StructureElement.typeObject.counterSiteMonolitBeamsTop Then
            elementAxis.Name = StructureElement.typeObject.counterSiteMonolitBeamsTop
        Else
            elementAxis.Name = StructureElement.typeObject.counterSiteMonolitBeamsBottom
        End If
        Dim deskObject As String = StructureElement.GetDescription(elementAxis.Name)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgPolyline3D
        Return elementAxis
    End Function

    Public Function setCounterPoint(ByVal polyCounter As DwgPolyline3D, ByVal heightPlate As Double) As List(Of PointStructure)
        Dim result As List(Of PointStructure) = New List(Of PointStructure)
        If polyCounter.Count = 4 Then
            Dim pointStartLeft As Vector3D = polyCounter.Item(0)
            Dim pointEndLeft As Vector3D = polyCounter.Item(1)
            Dim pointDownStartLeft As Vector3D = BridgeGeometry.calculatePointP(pointStartLeft, pointEndLeft, 0, -1 * heightPlate)
            Dim pointDownEndLeft As Vector3D = BridgeGeometry.calculatePointP(pointEndLeft, pointStartLeft, 0, -1 * heightPlate)

            Dim pointStartRight As Vector3D = polyCounter.Item(2)
            Dim pointEndRight As Vector3D = polyCounter.Item(3)
            Dim pointDownStartRight As Vector3D = BridgeGeometry.calculatePointP(pointStartRight, pointEndRight, 0, -1 * heightPlate)
            Dim pointDownEndRight As Vector3D = BridgeGeometry.calculatePointP(pointEndRight, pointStartRight, 0, -1 * heightPlate)
            result.Add(New PointStructure(pointStartLeft.X, pointStartLeft.Y, pointStartLeft.Z, pointDownStartLeft.X, pointDownStartLeft.Y, pointDownStartLeft.Z, 0, "leftPt1"))
            result.Add(New PointStructure(pointEndLeft.X, pointEndLeft.Y, pointEndLeft.Z, pointDownEndLeft.X, pointDownEndLeft.Y, pointDownEndLeft.Z, 0, "leftPt2"))

            result.Add(New PointStructure(pointStartRight.X, pointStartRight.Y, pointStartRight.Z, pointDownStartRight.X, pointDownStartRight.Y, pointDownStartRight.Z, 0, "rightPt2"))
            result.Add(New PointStructure(pointEndRight.X, pointEndRight.Y, pointEndRight.Z, pointDownEndRight.X, pointDownEndRight.Y, pointDownEndRight.Z, 0, "rightPt1"))
        End If
        Return result
    End Function
    Public Function setCounterPoint(ByVal polyCounter As DwgPolyline3D) As List(Of PointStructure)
        Dim result As List(Of PointStructure) = New List(Of PointStructure)
        If polyCounter.Count = 4 Then
            Dim pointStartLeft As Vector3D = polyCounter.Item(0)
            Dim pointEndLeft As Vector3D = polyCounter.Item(1)

            Dim pointStartRight As Vector3D = polyCounter.Item(2)
            Dim pointEndRight As Vector3D = polyCounter.Item(3)

            result.Add(New PointStructure(pointStartLeft.X, pointStartLeft.Y, pointStartLeft.Z, 0, 0, 0, 0, "leftPt1"))
            result.Add(New PointStructure(pointEndLeft.X, pointEndLeft.Y, pointEndLeft.Z, 0, 0, 0, 0, "leftPt2"))

            result.Add(New PointStructure(pointStartRight.X, pointStartRight.Y, pointStartRight.Z, 0, 0, 0, 0, "rightPt2"))
            result.Add(New PointStructure(pointEndRight.X, pointEndRight.Y, pointEndRight.Z, 0, 0, 0, 0, "rightPt1"))
        End If
        Return result
    End Function

    Public Shared Function getCounterMonolitSitesBeam(ByRef dictinaryAllObjectBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer, Optional typeCounter As StructureElement.typeObject = StructureElement.typeObject.counterSiteMonolitBeamsTop) As StructureElement
        Dim result As StructureElement = Nothing
        If dictinaryAllObjectBridge.ContainsKey(typeCounter) = True Then
            Dim listCounterMonolitSitesBeam As List(Of StructureElement) = dictinaryAllObjectBridge.Item(typeCounter)
            If listCounterMonolitSitesBeam.Count > 0 Then
                For i As Integer = 0 To listCounterMonolitSitesBeam.Count - 1
                    Dim dataStructure As StructureElement = listCounterMonolitSitesBeam.Item(i)
                    Dim tempCounterSiteMonolit As CounterSiteMonolitBeams = dataStructure.getCounterMonolitSiteBeams
                    If tempCounterSiteMonolit.TypeCounter = typeCounter Then
                        If IsNothing(tempCounterSiteMonolit) = False Then
                            If tempCounterSiteMonolit.numberProlet = numberProlet Then
                                If tempCounterSiteMonolit.numberRow = numberRow Then
                                    result = dataStructure
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        Return result
    End Function
End Class
