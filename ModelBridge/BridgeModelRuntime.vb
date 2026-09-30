Imports System.Runtime.CompilerServices
Imports Topomatic.Arrangements
Imports Topomatic.Dwg

Public NotInheritable Class BridgeModelRuntime
    Private Shared ReadOnly Roots As New ConditionalWeakTable(Of ArrangementModel, InfrastradaBridgesModel)()

    Private Sub New()
    End Sub

    Public Shared Function GetArrangement(model As Object) As ArrangementModel
        Dim arrangement As ArrangementModel = TryCast(model, ArrangementModel)
        If arrangement IsNot Nothing Then Return arrangement
        Dim root As InfrastradaBridgesModel = TryCast(model, InfrastradaBridgesModel)
        If root IsNot Nothing Then Return root.Arrangement
        Return Nothing
    End Function

    Public Shared Function GetDrawing(arrangement As ArrangementModel) As Drawing
        If arrangement Is Nothing Then Return Nothing
        Dim root As InfrastradaBridgesModel = GetRoot(arrangement)
        Return If(root Is Nothing, arrangement.Drawing, root.Drawing)
    End Function

    Public Shared Function GetRoot(arrangement As ArrangementModel) As InfrastradaBridgesModel
        If arrangement Is Nothing Then Return Nothing
        Dim root As InfrastradaBridgesModel = Nothing
        Roots.TryGetValue(arrangement, root)
        Return root
    End Function

    Public Shared Function GetModelObject(arrangement As ArrangementModel) As Object
        If arrangement Is Nothing Then Return Nothing
        Dim root As InfrastradaBridgesModel = GetRoot(arrangement)
        Return If(root Is Nothing, DirectCast(arrangement, Object), DirectCast(root, Object))
    End Function

    Friend Shared Sub Register(root As InfrastradaBridgesModel, arrangement As ArrangementModel)
        If root Is Nothing OrElse arrangement Is Nothing Then Return
        Roots.Remove(arrangement)
        Roots.Add(arrangement, root)
    End Sub

    Friend Shared Sub Unregister(arrangement As ArrangementModel)
        If arrangement IsNot Nothing Then Roots.Remove(arrangement)
    End Sub
End Class
