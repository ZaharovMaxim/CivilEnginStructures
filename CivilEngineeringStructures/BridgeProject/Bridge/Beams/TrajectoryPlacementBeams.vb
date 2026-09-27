'КЛАСС ДЛЯ ТРАЕКТОРИЙ РАСКЛАДКИ БАЛОК
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath

Public Class TrajectoryPlacementBeams
    Public Enum TypeTrajectoryPlacementBeams
        <Description("Ось трассы")> ProjectAlignment = 0
        <Description("Пользовательская")> UserPolyline = 1
        <Description("Не определено")> None = 2
    End Enum
    Private _numberRow As Integer 'номер ряда
    Private _offsetProjectSurface As Double 'смещение от проектной поверхности
    Private _offsetProjectAlignment As Double 'смещение от проектной оси
    Private _typeTrajectoryPlacementBeams As TypeTrajectoryPlacementBeams 'тип оси раскладки балок
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberRow = 0
        _offsetProjectAlignment = 0
        _offsetProjectSurface = 0
        _typeTrajectoryPlacementBeams = TypeTrajectoryPlacementBeams.None
    End Sub
    Public Sub New(NumberRow As Integer, OffsetProjectSurface As Double, OffsetProjectAlignment As Double, TypeTrajectory As TypeTrajectoryPlacementBeams)
        _numberRow = NumberRow
        _offsetProjectAlignment = OffsetProjectAlignment
        _offsetProjectSurface = OffsetProjectSurface
        _typeTrajectoryPlacementBeams = TypeTrajectory
    End Sub

    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property NumberRows() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение относительно проектной поверхности, м")>
    <Category("Свойства")>
    <DisplayName("Вертикальное смещение")>
    Public Property OffsetProjectSurface() As Double
        Get
            Return _offsetProjectSurface
        End Get
        Set(value As Double)
            _offsetProjectSurface = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Смещение относительно проектной трассы, м")>
    <Category("Свойства")>
    <DisplayName("Горизонтальное смещение")>
    Public Property OffsetProjectAlignment() As Double
        Get
            Return _offsetProjectAlignment
        End Get
        Set(value As Double)
            _offsetProjectAlignment = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Тип траектории (пользовательская\на основе трассы")>
    <Category("Свойства")>
    <DisplayName("Тип траектории")>
    Public Property TypeTrajectory() As TypeTrajectoryPlacementBeams
        Get
            Return _typeTrajectoryPlacementBeams
        End Get
        Set(value As TypeTrajectoryPlacementBeams)
            _typeTrajectoryPlacementBeams = value
        End Set
    End Property
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новую пустую насадку
    Public Shared Function createTrajectoryPlacementBeams(ByVal idBridge As String) As StructureElement
        Dim element As StructureElement = New StructureElement()
        element.Label = "Мосты и путепроводы"
        element.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        element.ClassObject = StructureElement.classStructure.trajectoryPlacementBeams
        element.Name = StructureElement.typeObject.axisTrajectoryPlacementBeams
        element.Description = StructureElement.GetDescription(element.Name)
        element.KeyParameter = ""
        element.IdElement = Guid.NewGuid.ToString
        element.IdStructure = idBridge
        element.Note = ""
        element.DWGEntity = New DwgPolyline
        Return element
    End Function

    'ищет ось раскладки балок
    Public Shared Function getTrajectoryPlacementBeams(ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberRow As Integer) As StructureElement
        Dim dataAxisPlacementBeams As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisTrajectoryPlacementBeams) = True Then
            Dim listAxisPlacementBeams = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisTrajectoryPlacementBeams)
            If IsNothing(listAxisPlacementBeams) = False Then
                If listAxisPlacementBeams.Count > 0 Then
                    For k As Integer = 0 To listAxisPlacementBeams.Count - 1
                        Dim tempData As StructureElement = listAxisPlacementBeams.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisPlacementBeams As TrajectoryPlacementBeams = tempData.getAxisPlacementBeams()
                            If IsNothing(userAxisPlacementBeams) = False Then
                                If numberRow = userAxisPlacementBeams.numberRows Then
                                    dataAxisPlacementBeams = tempData
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataAxisPlacementBeams
    End Function
End Class
