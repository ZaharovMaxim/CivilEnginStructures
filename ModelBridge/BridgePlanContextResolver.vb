Imports Topomatic.ApplicationPlatform
Imports Topomatic.Cad.View

Friend NotInheritable Class BridgePlanContextResolver
    Private Sub New()
    End Sub

    Friend Shared Function Resolve(cadView As CadView) As BridgePlanCompoundLayer
        If cadView Is Nothing Then Return Nothing

        Dim direct As BridgePlanCompoundLayer = TryCast(
            cadView(BridgePlanCompoundLayer.Guid),
            BridgePlanCompoundLayer)
        If direct IsNot Nothing Then Return EnabledBridge(direct)

        Dim models As MultiLayer = TryCast(cadView(Consts.ModelsLayer), MultiLayer)
        If models Is Nothing Then Return Nothing
        Return ResolveLayer(models.ResolveActive())
    End Function

    Friend Shared Function ResolveLayer(rootLayer As CadViewLayer) As BridgePlanCompoundLayer
        If rootLayer Is Nothing Then Return Nothing

        Dim bridge As BridgePlanCompoundLayer = TryCast(rootLayer, BridgePlanCompoundLayer)
        If bridge IsNot Nothing Then Return EnabledBridge(bridge)

        Dim multi As MultiLayer = TryCast(rootLayer, MultiLayer)
        If multi IsNot Nothing Then
            Dim active As CadViewLayer = multi.ResolveActive()
            If Object.ReferenceEquals(active, rootLayer) Then Return Nothing
            Return ResolveLayer(active)
        End If

        Dim compound As CompoundLayer = TryCast(rootLayer, CompoundLayer)
        If compound Is Nothing Then Return Nothing

        Dim candidate As CadViewLayer = compound(BridgePlanCompoundLayer.Guid)
        If candidate Is Nothing OrElse Object.ReferenceEquals(candidate, rootLayer) Then Return Nothing
        Return ResolveLayer(candidate)
    End Function

    Private Shared Function EnabledBridge(layer As BridgePlanCompoundLayer) As BridgePlanCompoundLayer
        If layer Is Nothing OrElse Not layer.ResolveEnable() Then Return Nothing
        Return layer
    End Function
End Class
