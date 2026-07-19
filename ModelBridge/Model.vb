Imports System.Collections.Generic
Imports Topomatic.Cad.Foundation
Imports Topomatic.FoundationClasses
Imports Topomatic.Stg
Public Class Model
    Inherits StateControllerObject
    Implements IStgSerializable

    Private m_ReadOnly As Boolean = False

    Public Points As New List(Of Vector2D)()

    ' Флаг только для чтения
    Public Overrides Property [ReadOnly] As Boolean
        Get
            Return m_ReadOnly
        End Get
        Set(value As Boolean)
            m_ReadOnly = value
        End Set
    End Property

    ' Загрузка из узла
    Public Sub LoadFromStg(node As StgNode) Implements IStgSerializable.LoadFromStg
        Points.Clear()
        ' При загрузке массива указывается тип составляющих массив значений
        Dim array As IStgArray = node.GetArray("Points", StgType.Node)
        For i As Integer = 0 To array.Count - 1
            Points.Add(Vector2D.LoadFromStg(array.GetNode(i)))
        Next
    End Sub

    ' Сохранение в узел
    Public Sub SaveToStg(node As StgNode) Implements IStgSerializable.SaveToStg
        ' Сохраняем значения в узел
        ' Сохраняем массив с указанием типа значений
        Dim array As IStgArray = node.AddArray("Points", StgType.Node)
        For i As Integer = 0 To Points.Count - 1
            Points(i).SaveToStg(array.AddNode())
        Next
    End Sub
End Class
