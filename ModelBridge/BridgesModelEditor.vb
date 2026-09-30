Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.ApplicationPlatform.ServiceClasses
Imports Topomatic.Arrangements
Imports Topomatic.Cad.View
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.Sfc
Imports Newtonsoft.Json

Public Class BridgesModelEditor
    Inherits PlanModelEditor

    Public Overrides Function LoadFromFile(fullpath As String) As Object
        Dim bridgeModel As New InfrastradaBridgesModel()
        Try
            bridgeModel.ConfigureModelFinder(New ModelFinder(bridgeModel))
            If Not String.IsNullOrWhiteSpace(fullpath) Then
                bridgeModel.NativeLoadFromFile(fullpath)
            End If
            BridgeDrawingGroupManager.Attach(bridgeModel.Drawing)
            Return bridgeModel
        Catch
            bridgeModel.Dispose()
            Throw
        End Try
    End Function

    Public Overrides Sub SaveToFile(model As Object, fullpath As String)
        Dim bridgeModel As InfrastradaBridgesModel = TryCast(model, InfrastradaBridgesModel)
        If bridgeModel Is Nothing OrElse String.IsNullOrWhiteSpace(fullpath) Then Return
        bridgeModel.CreateLegacyBackup(fullpath)
        bridgeModel.NativeSaveToFile(fullpath)
    End Sub

    Protected Overrides Function CreatePlanLayer(model As IProjectModel) As CadViewLayer
        Dim bridgeModel As InfrastradaBridgesModel = PluginCoreOps.LockReadContainer(Of InfrastradaBridgesModel)(model)
        If bridgeModel Is Nothing Then Return Nothing
        BridgeDrawingGroupManager.Attach(bridgeModel.Drawing)
        ConnectEvents(bridgeModel)
        Dim layer As New BridgePlanCompoundLayer(PluginCoreOps.FindModelPathId(model)) With {
            .BridgeModel = bridgeModel
        }
        CreateCadView3d(model.Project)
        ApplicationHost.Current.Plugins.Broadcast(Consts.BroadcastAddLayer,
                                                  New String() {model.ModelType, Consts.PlanWindow},
                                                  New Object() {model, layer})
        Return layer
    End Function

    Protected Overrides Sub ReloadModel(model As IProjectModel, editorResult As EditorResult)
        Dim layer As BridgePlanCompoundLayer = TryCast(editorResult.PlanLayer, BridgePlanCompoundLayer)
        If layer Is Nothing Then Return
        DisconnectEvents(layer.BridgeModel)
        Dim bridgeModel As InfrastradaBridgesModel = PluginCoreOps.LockReadContainer(Of InfrastradaBridgesModel)(model)
        If bridgeModel Is Nothing Then Return
        bridgeModel.ConfigureModelFinder(New ModelFinder(bridgeModel))
        bridgeModel.Drawing.Filename = FileManager.LocalDatabasePath(model.Uri.AsAbsoluteUri)
        BridgeDrawingGroupManager.Attach(bridgeModel.Drawing)
        ConnectEvents(bridgeModel)
        layer.BridgeModel = bridgeModel
    End Sub

    Protected Overrides Sub RemovePlanLayer(model As IProjectModel, layer As CadViewLayer)
        Dim bridgeLayer As BridgePlanCompoundLayer = TryCast(layer, BridgePlanCompoundLayer)
        If bridgeLayer Is Nothing Then Return
        RemoveCadView3d(model.Project)
        ApplicationHost.Current.Plugins.Broadcast(Consts.BroadcastRemoveLayer,
                                                  New String() {model.ModelType, Consts.PlanWindow},
                                                  New Object() {bridgeLayer})
        DisconnectEvents(bridgeLayer.BridgeModel)
        bridgeLayer.BridgeModel = Nothing
    End Sub

    Public Overrides Function GetHardReferences(model As IProjectModel) As ModelHardReference()
        Dim result As New List(Of ModelHardReference)()
        Dim root As InfrastradaBridgesModel = PluginCoreOps.LockReadContainer(Of InfrastradaBridgesModel)(model)
        Dim bridgeModel As ArrangementModel = If(root Is Nothing, PluginCoreOps.LockReadContainer(Of ArrangementModel)(model), root.Arrangement)
        If bridgeModel Is Nothing Then Return result.ToArray()
        For Each path As String In bridgeModel.ProjectSurfacesRelativePaths
            AddReference(result, path)
        Next
        If bridgeModel.HasBasisCurve Then AddReference(result, bridgeModel.BasisCurveRelativePath)
        Dim settings As BridgeModelSettings = BridgeModelSettingsStore.GetSettings(bridgeModel)
        AddReference(result, settings.ProjectSurfaceRelativePath)
        AddReference(result, settings.EarthSurfaceRelativePath)
        AddBridgeSurfaceReferences(result, bridgeModel)
        Return result.ToArray()
    End Function

    Public Overrides Sub ReplaceObjectHardReference(model As Object, oldValue As URI, newValue As URI)
        Dim bridgeModel As ArrangementModel = BridgeModelRuntime.GetArrangement(model)
        If bridgeModel Is Nothing Then Return
        Dim root As InfrastradaBridgesModel = TryCast(model, InfrastradaBridgesModel)
        If root Is Nothing Then root = BridgeModelRuntime.GetRoot(bridgeModel)
        Dim changed As Boolean = False
        If root IsNot Nothing Then
            root.BeginUpdate()
        Else
            bridgeModel.BeginUpdate()
        End If
        Try
            For i As Integer = 0 To bridgeModel.ProjectSurfacesRelativePaths.Count - 1
                If String.Equals(bridgeModel.ProjectSurfacesRelativePaths(i), oldValue.AsAbsoluteUri, StringComparison.Ordinal) Then
                    bridgeModel.ProjectSurfacesRelativePaths(i) = newValue.AsAbsoluteUri
                    changed = True
                End If
            Next
            If bridgeModel.HasBasisCurve AndAlso String.Equals(bridgeModel.BasisCurveRelativePath, oldValue.AsAbsoluteUri, StringComparison.Ordinal) Then
                bridgeModel.BasisCurveRelativePath = newValue.AsAbsoluteUri
                changed = True
            End If
            Dim settings As BridgeModelSettings = BridgeModelSettingsStore.GetSettings(bridgeModel)
            If String.Equals(settings.ProjectSurfaceRelativePath, oldValue.AsAbsoluteUri, StringComparison.Ordinal) OrElse
               String.Equals(settings.EarthSurfaceRelativePath, oldValue.AsAbsoluteUri, StringComparison.Ordinal) Then changed = True
            settings.ReplaceReference(oldValue.AsAbsoluteUri, newValue.AsAbsoluteUri)
            If ReplaceBridgeSurfaceReferences(bridgeModel, oldValue.AsAbsoluteUri, newValue.AsAbsoluteUri) Then changed = True
        Finally
            If root IsNot Nothing Then
                If changed Then root.Modified = True
                root.EndUpdate()
            Else
                bridgeModel.EndUpdate()
            End If
        End Try
    End Sub

    Private Shared Sub AddReference(result As List(Of ModelHardReference), path As String)
        If String.IsNullOrWhiteSpace(path) Then Return
        For Each item As ModelHardReference In result
            If String.Equals(item.Uri.AsAbsoluteUri, path, StringComparison.OrdinalIgnoreCase) Then Return
        Next
        result.Add(New ModelHardReference(New URI(path)))
    End Sub

    Private Shared Function ReplaceBridgeSurfaceReferences(model As ArrangementModel, oldPath As String, newPath As String) As Boolean
        Dim drawing = BridgeModelRuntime.GetDrawing(model)
        If drawing Is Nothing Then Return False
        Dim result As Boolean = False
        For Each entity As DwgEntity In drawing.ActiveSpace.Entities
            Dim data As New StructureElement()
            If Not FuncXRecords.getXRecords(entity, data) OrElse data.Name <> StructureElement.typeObject.axisBridge Then Continue For
            Dim bridge As Bridges = data.getBridge()
            If bridge Is Nothing Then Continue For
            Dim changed As Boolean = False
            If String.Equals(bridge.projectSurfaceName, oldPath, StringComparison.Ordinal) Then
                bridge.projectSurfaceName = newPath
                changed = True
            End If
            If String.Equals(bridge.EarthSurfaceName, oldPath, StringComparison.Ordinal) Then
                bridge.EarthSurfaceName = newPath
                changed = True
            End If
            If changed Then
                data.KeyParameter = JsonConvert.SerializeObject(bridge)
                If FuncXRecords.setXRecords(entity, StructureElement.tableXRecords.PROJECT_STRUCTURES, data) Then result = True
            End If
        Next
        Return result
    End Function

    Private Shared Sub AddBridgeSurfaceReferences(result As List(Of ModelHardReference), model As ArrangementModel)
        Dim drawing = BridgeModelRuntime.GetDrawing(model)
        If drawing Is Nothing Then Return
        For Each entity As DwgEntity In drawing.ActiveSpace.Entities
            Dim data As New StructureElement()
            If Not FuncXRecords.getXRecords(entity, data) OrElse data.Name <> StructureElement.typeObject.axisBridge Then Continue For
            Dim bridge As Bridges = data.getBridge()
            If bridge Is Nothing Then Continue For
            AddStoredSurfaceReference(result, bridge.projectSurfaceName)
            AddStoredSurfaceReference(result, bridge.EarthSurfaceName)
        Next
    End Sub

    Private Shared Sub AddStoredSurfaceReference(result As List(Of ModelHardReference), path As String)
        If String.IsNullOrWhiteSpace(path) OrElse
           (path.IndexOf("\"c) < 0 AndAlso path.IndexOf("/"c) < 0 AndAlso String.IsNullOrWhiteSpace(System.IO.Path.GetExtension(path))) Then Return
        AddReference(result, path)
    End Sub

    Private Shared Sub ConnectEvents(model As InfrastradaBridgesModel)
        If model Is Nothing OrElse model.Surface Is Nothing Then Return
        RemoveHandler model.Surface.Changed, AddressOf SurfaceChanged
        RemoveHandler model.Surface.Undo, AddressOf SurfaceChanged
        AddHandler model.Surface.Changed, AddressOf SurfaceChanged
        AddHandler model.Surface.Undo, AddressOf SurfaceChanged
    End Sub

    Private Shared Sub DisconnectEvents(model As InfrastradaBridgesModel)
        If model Is Nothing OrElse model.Surface Is Nothing Then Return
        RemoveHandler model.Surface.Changed, AddressOf SurfaceChanged
        RemoveHandler model.Surface.Undo, AddressOf SurfaceChanged
    End Sub

    Private Shared Sub SurfaceChanged(sender As Object, e As EventArgs)
        Dim surface As Surface = TryCast(sender, Surface)
        If surface Is Nothing Then Return
        Dim root As InfrastradaBridgesModel = TryCast(surface.Owner, InfrastradaBridgesModel)
        If root IsNot Nothing Then root.Modified = True
        ApplicationHost.Current.Plugins.Broadcast("surface_changed",
                                                  New String() {"infrastrada_bridges"},
                                                  New Object() {sender})
    End Sub
End Class
