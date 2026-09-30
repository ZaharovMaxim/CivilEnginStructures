Imports System.Runtime.CompilerServices
Imports System.IO
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Sfc
Imports Topomatic.Stg

Public NotInheritable Class BridgeModelSettings
    Public Property ProjectSurfaceRelativePath As String
    Public Property EarthSurfaceRelativePath As String

    Public Sub LoadFromStg(node As StgNode)
        ProjectSurfaceRelativePath = node.GetString("ProjectSurfaceRelativePath", String.Empty)
        EarthSurfaceRelativePath = node.GetString("EarthSurfaceRelativePath", String.Empty)
    End Sub

    Public Sub SaveToStg(node As StgNode)
        node.AddString("ProjectSurfaceRelativePath", If(ProjectSurfaceRelativePath, String.Empty))
        node.AddString("EarthSurfaceRelativePath", If(EarthSurfaceRelativePath, String.Empty))
    End Sub

    Public Sub ReplaceReference(oldPath As String, newPath As String)
        If String.Equals(ProjectSurfaceRelativePath, oldPath, StringComparison.Ordinal) Then
            ProjectSurfaceRelativePath = newPath
        End If
        If String.Equals(EarthSurfaceRelativePath, oldPath, StringComparison.Ordinal) Then
            EarthSurfaceRelativePath = newPath
        End If
    End Sub

    Friend Function TryInitializeProjectSurfaceReference(path As String) As Boolean
        If String.IsNullOrWhiteSpace(path) OrElse Not String.IsNullOrWhiteSpace(ProjectSurfaceRelativePath) Then Return False
        ProjectSurfaceRelativePath = path
        Return True
    End Function
End Class

Public NotInheritable Class BridgeModelSettingsStore
    Public Const NodeName As String = "InfrastradaBridgeSettings"

    Private Shared ReadOnly SettingsByModel As New ConditionalWeakTable(Of ArrangementModel, BridgeModelSettings)()

    Private Sub New()
    End Sub

    Public Shared Function GetSettings(model As ArrangementModel) As BridgeModelSettings
        If model Is Nothing Then Return Nothing
        Return SettingsByModel.GetValue(model, Function(key) New BridgeModelSettings())
    End Function

    Public Shared Sub Load(model As ArrangementModel, body As StgNode)
        Dim settings As BridgeModelSettings = GetSettings(model)
        settings.LoadFromStg(body.GetNode(NodeName))
        MirrorSurfaceReferences(model, settings)
    End Sub

    Public Shared Sub Save(model As ArrangementModel, body As StgNode)
        Dim settings As BridgeModelSettings = GetSettings(model)
        MirrorSurfaceReferences(model, settings)
        settings.SaveToStg(body.AddNode(NodeName))
    End Sub

    Public Shared Sub MirrorSurfaceReferences(model As ArrangementModel, settings As BridgeModelSettings)
        If model Is Nothing OrElse settings Is Nothing Then Return
        AddReference(model, settings.ProjectSurfaceRelativePath)
        AddReference(model, settings.EarthSurfaceRelativePath)
    End Sub

    Public Shared Sub UpdateSurfaceReferences(model As ArrangementModel, projectPath As String, earthPath As String)
        Dim settings As BridgeModelSettings = GetSettings(model)
        Dim oldProjectPath As String = settings.ProjectSurfaceRelativePath
        Dim oldEarthPath As String = settings.EarthSurfaceRelativePath
        settings.ProjectSurfaceRelativePath = If(projectPath, String.Empty)
        settings.EarthSurfaceRelativePath = If(earthPath, String.Empty)
        RemoveOldRoleReference(model, oldProjectPath, settings)
        RemoveOldRoleReference(model, oldEarthPath, settings)
        MirrorSurfaceReferences(model, settings)
    End Sub

    Friend Shared Function TryInitializeProjectSurfaceReference(model As ArrangementModel, path As String) As Boolean
        If model Is Nothing Then Return False
        Dim settings As BridgeModelSettings = GetSettings(model)
        If settings Is Nothing OrElse Not settings.TryInitializeProjectSurfaceReference(path) Then Return False
        MirrorSurfaceReferences(model, settings)
        Return True
    End Function

    Public Shared Function ResolveSurface(model As ArrangementModel, relativePath As String) As Surface
        If model Is Nothing OrElse model.ModelFinder Is Nothing OrElse String.IsNullOrWhiteSpace(relativePath) Then Return Nothing
        Dim container As ISurfaceContainer = model.ModelFinder.ReadModelFromPath(Of ISurfaceContainer)(relativePath)
        If container Is Nothing Then Return Nothing
        Return container.Surface
    End Function

    Public Shared Function ResolveReferenceName(model As ArrangementModel, relativePath As String) As String
        If model Is Nothing OrElse String.IsNullOrWhiteSpace(relativePath) Then Return String.Empty
        Dim owner = PluginCoreOps.FindModel(BridgeModelRuntime.GetModelObject(model))
        Dim reference = PluginCoreOps.FindModel(owner, relativePath)
        If reference Is Nothing Then Return String.Empty
        Return Path.GetFileNameWithoutExtension(reference.Uri.LastPathComponent)
    End Function

    Private Shared Sub AddReference(model As ArrangementModel, relativePath As String)
        If String.IsNullOrWhiteSpace(relativePath) Then Return
        For Each item As String In model.ProjectSurfacesRelativePaths
            If String.Equals(item, relativePath, StringComparison.OrdinalIgnoreCase) Then Return
        Next
        model.ProjectSurfacesRelativePaths.Add(relativePath)
    End Sub

    Private Shared Sub RemoveOldRoleReference(model As ArrangementModel, relativePath As String, settings As BridgeModelSettings)
        If String.IsNullOrWhiteSpace(relativePath) OrElse
           String.Equals(relativePath, settings.ProjectSurfaceRelativePath, StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(relativePath, settings.EarthSurfaceRelativePath, StringComparison.OrdinalIgnoreCase) Then Return
        For i As Integer = model.ProjectSurfacesRelativePaths.Count - 1 To 0 Step -1
            If String.Equals(model.ProjectSurfacesRelativePaths(i), relativePath, StringComparison.OrdinalIgnoreCase) Then
                model.ProjectSurfacesRelativePaths.RemoveAt(i)
            End If
        Next
    End Sub
End Class
