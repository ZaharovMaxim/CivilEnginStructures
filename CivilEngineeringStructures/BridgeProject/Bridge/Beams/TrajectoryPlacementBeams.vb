'КЛАСС ДЛЯ ТРАЕКТОРИЙ РАСКЛАДКИ БАЛОК
Imports System.ComponentModel

Public Class TrajectoryPlacementBeams
    Private _numberRow As Integer 'номер ряда
    Private _points As List(Of PointStructure)
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberRow = 0
        _points = New List(Of PointStructure)
    End Sub
    Public Sub New(numberRow As Integer, Points As List(Of PointStructure))
        _numberRow = numberRow
        _points = Points
    End Sub

    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property numberRows() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    Public Property Points() As List(Of PointStructure)
        Get
            Return _points
        End Get
        Set(value As List(Of PointStructure))
            _points = value
        End Set
    End Property

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
