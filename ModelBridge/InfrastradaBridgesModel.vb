Imports System.IO
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Sfc
Imports Topomatic.Stg

Public NotInheritable Class InfrastradaBridgesModel
    Inherits StateControllerObject
    Implements IDisposable, IDrawingContainer, ITerrainModel, ISurfaceContainer, IStateElevationProviderFactory

    Private Const MetadataNodeName As String = "InfrastradaBridgesModel"
    Private Const ArrangementNodeName As String = "Arrangement"
    Private Const CurrentVersion As Integer = 1

    Private _surface As Surface
    Private _arrangement As ArrangementModel
    Private _loadArrangement As ArrangementModel
    Private _metadataLoaded As Boolean
    Private _disposed As Boolean
    Private _loadedLegacy As Boolean

    Public Sub New()
        BridgeBeamPlanEntity.RegisterActivator()
        _surface = CreateSurface()
        _arrangement = New ArrangementModel()
        BridgeModelRuntime.Register(Me, _arrangement)
        BridgeDrawingGroupManager.Attach(Drawing)
    End Sub

    Public ReadOnly Property Arrangement As ArrangementModel
        Get
            Return _arrangement
        End Get
    End Property

    Public ReadOnly Property Surface As Surface Implements ISurfaceContainer.Surface
        Get
            Return _surface
        End Get
    End Property

    Public ReadOnly Property Drawing As Drawing Implements IDrawingContainer.Drawing
        Get
            Return If(_surface Is Nothing, Nothing, _surface.Situation)
        End Get
    End Property

    Friend ReadOnly Property LoadedLegacyArrangement As Boolean
        Get
            Return _loadedLegacy
        End Get
    End Property

    Public Sub ConfigureModelFinder(finder As IModelFinder)
        ModelFinder = finder
        If _arrangement IsNot Nothing Then _arrangement.ModelFinder = finder
    End Sub

    Public Sub NativeSaveToFile(path As String)
        Dim fullPath As String = System.IO.Path.GetFullPath(path)
        Dim directory As String = System.IO.Path.GetDirectoryName(fullPath)
        Dim temporaryPath As String = System.IO.Path.Combine(
            directory,
            System.IO.Path.GetFileName(fullPath) & "." & System.Guid.NewGuid().ToString("N") & ".tmp")
        Try
            Using stream As New FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)
                NativeSaveToStream(stream)
            End Using
            If File.Exists(fullPath) Then
                File.Replace(temporaryPath, fullPath, Nothing)
            Else
                File.Move(temporaryPath, fullPath)
            End If
            _loadedLegacy = False
        Finally
            If File.Exists(temporaryPath) Then File.Delete(temporaryPath)
        End Try
    End Sub

    Public Sub NativeLoadFromFile(path As String)
        Using stream As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)
            NativeLoadFromStream(stream)
        End Using
        Drawing.Filename = path
    End Sub

    Public Sub NativeSaveToStream(stream As Stream)
        If stream Is Nothing Then Throw New ArgumentNullException(NameOf(stream))
        If _arrangement Is Nothing OrElse _surface Is Nothing Then Throw New ObjectDisposedException(NameOf(InfrastradaBridgesModel))

        If _arrangement.Drawing.ActiveSpace.Entities.Count > 0 Then
            Throw New InvalidOperationException("Во внутреннем хранилище обнаружены объекты. Сохранение остановлено, чтобы не потерять геометрию моста.")
        End If
        AddHandler _surface.SaveToStg, AddressOf SaveMetadata
        Try
            _surface.SaveToStreamSfcx(stream)
        Finally
            RemoveHandler _surface.SaveToStg, AddressOf SaveMetadata
        End Try
    End Sub

    Public Sub NativeLoadFromStream(stream As Stream)
        If stream Is Nothing Then Throw New ArgumentNullException(NameOf(stream))

        Using payload As New MemoryStream()
            stream.CopyTo(payload)
            If payload.Length = 0 Then Throw New InvalidDataException("Файл модели мостов пуст.")

            Dim transactionManager As ITransactionManager = Me.TransactionManager
            SetTransactionManager(Nothing)
            Try
                payload.Position = 0
                If TryLoadNative(payload) Then
                    _loadedLegacy = False
                    Return
                End If

                payload.Position = 0
                LoadLegacyArrangement(payload)
                _loadedLegacy = True
            Finally
                SetTransactionManager(transactionManager)
            End Try
        End Using
    End Sub

    Friend Sub CreateLegacyBackup(path As String)
        If Not _loadedLegacy OrElse String.IsNullOrWhiteSpace(path) OrElse Not File.Exists(path) Then Return
        Dim backupPath As String = path & ".legacy-arrangement.bak"
        If Not File.Exists(backupPath) Then File.Copy(path, backupPath, False)
    End Sub

    Public Function GetReferences() As IEnumerable(Of String) Implements IReferenceHolder.GetReferences
        Dim result As New List(Of String)()
        If _arrangement Is Nothing Then Return result
        For Each path As String In _arrangement.ProjectSurfacesRelativePaths
            AddReference(result, path)
        Next
        Dim settings As BridgeModelSettings = BridgeModelSettingsStore.GetSettings(_arrangement)
        If settings IsNot Nothing Then
            AddReference(result, settings.ProjectSurfaceRelativePath)
            AddReference(result, settings.EarthSurfaceRelativePath)
        End If
        Return result
    End Function

    Private Function CreateProvider() As StateElevationProvider Implements IStateElevationProviderFactory.CreateProvider
        Return New BridgeElevationProvider(Me)
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True
        BridgeDrawingGroupManager.Detach(Drawing)
        BridgeModelRuntime.Unregister(_arrangement)
        If _surface IsNot Nothing Then _surface.Dispose()
        If _arrangement IsNot Nothing Then _arrangement.Dispose()
        _surface = Nothing
        _arrangement = Nothing
    End Sub

    Private Function CreateSurface() As Surface
        Dim result As New Surface(Me, False) With {
            .Code = 400,
            .Designed = True
        }
        result.Style.Dynamic = False
        result.Situation.Owner = Me
        Return result
    End Function

    Private Function TryLoadNative(payload As Stream) As Boolean
        Dim stagedSurface As Surface = CreateSurface()
        Dim stagedArrangement As New ArrangementModel()
        stagedArrangement.ModelFinder = ModelFinder
        _loadArrangement = stagedArrangement
        _metadataLoaded = False
        AddHandler stagedSurface.LoadFromStg, AddressOf LoadMetadata
        Dim parsed As Boolean = False
        Try
            stagedSurface.LoadFromStream(payload)
            parsed = _metadataLoaded AndAlso _loadArrangement IsNot Nothing
        Catch ex As Exception
        Finally
            RemoveHandler stagedSurface.LoadFromStg, AddressOf LoadMetadata
            _loadArrangement = Nothing
            _metadataLoaded = False
        End Try

        If Not parsed Then
            stagedSurface.Dispose()
            stagedArrangement.Dispose()
            Return False
        End If

        If stagedArrangement.Drawing.ActiveSpace.Entities.Count > 0 Then
            stagedSurface.Dispose()
            stagedArrangement.Dispose()
            Throw New InvalidDataException("В служебной части файла обнаружена геометрия моста. Загрузка остановлена, чтобы не потерять объекты.")
        End If

        Try
            ReplaceArrangement(stagedArrangement)
            stagedArrangement = Nothing
            ReplaceSurface(stagedSurface)
            stagedSurface = Nothing
            Return True
        Finally
            If stagedSurface IsNot Nothing Then stagedSurface.Dispose()
            If stagedArrangement IsNot Nothing Then stagedArrangement.Dispose()
        End Try
    End Function

    Private Sub LoadLegacyArrangement(payload As Stream)
        Dim document As New StgDocument()
        Try
            document.LoadFromStreamAsBinary(payload)
        Catch ex As Exception
            Try
                payload.Position = 0
                document.LoadFromStreamAsXml(payload)
            Catch xmlException As Exception
                Throw New InvalidDataException("Файл не является моделью мостов поддерживаемого формата.", ex)
            End Try
        End Try
        If Not document.Body.IsExists(BridgeModelSettingsStore.NodeName) Then
            Throw New InvalidDataException("Повреждённый файл собственной модели мостов не может быть открыт как старая модель.")
        End If

        Dim legacy As ArrangementModel = New ArrangementModel()
        Dim migratedSurface As Surface = Nothing
        Try
            legacy.ModelFinder = ModelFinder
            legacy.LoadFromStg(document.Body)
            BridgeModelSettingsStore.Load(legacy, document.Body)

            migratedSurface = CreateSurface()
            migratedSurface.Situation.Assign(legacy.Drawing)
            legacy.Drawing.Clear()
            ReplaceArrangement(legacy)
            legacy = Nothing
            ReplaceSurface(migratedSurface)
            migratedSurface = Nothing
            BridgeDrawingGroupManager.Attach(Drawing)
        Finally
            If migratedSurface IsNot Nothing Then migratedSurface.Dispose()
            If legacy IsNot Nothing Then legacy.Dispose()
        End Try
    End Sub

    Private Sub SaveMetadata(sender As Object, e As StgDocumentOperationEventArgs)
        Dim node As StgNode = e.Document.Body.AddNode(MetadataNodeName)
        node.AddInt32("Version", CurrentVersion)
        _arrangement.SaveToStg(node.AddNode(ArrangementNodeName))
        BridgeModelSettingsStore.Save(_arrangement, node)
    End Sub

    Private Sub LoadMetadata(sender As Object, e As StgDocumentOperationEventArgs)
        If Not e.Document.Body.IsExists(MetadataNodeName) Then
            _loadArrangement = Nothing
            Return
        End If
        Dim node As StgNode = e.Document.Body.GetNode(MetadataNodeName)
        If node.GetInt32("Version", 0) <> CurrentVersion OrElse Not node.IsExists(ArrangementNodeName) Then
            _loadArrangement = Nothing
            Return
        End If
        _loadArrangement.LoadFromStg(node.GetNode(ArrangementNodeName))
        BridgeModelSettingsStore.Load(_loadArrangement, node)
        _metadataLoaded = True
    End Sub

    Private Sub ReplaceArrangement(value As ArrangementModel)
        Dim oldArrangement As ArrangementModel = _arrangement
        If Object.ReferenceEquals(oldArrangement, value) Then Return
        BridgeModelRuntime.Unregister(oldArrangement)
        _arrangement = value
        _arrangement.ModelFinder = ModelFinder
        BridgeModelRuntime.Register(Me, _arrangement)
        If oldArrangement IsNot Nothing Then oldArrangement.Dispose()
    End Sub

    Private Sub ReplaceSurface(value As Surface)
        Dim oldSurface As Surface = _surface
        If Object.ReferenceEquals(oldSurface, value) Then Return
        BridgeDrawingGroupManager.Detach(If(oldSurface Is Nothing, Nothing, oldSurface.Situation))
        _surface = value
        _surface.Owner = Me
        _surface.Situation.Owner = Me
        BridgeDrawingGroupManager.Attach(_surface.Situation)
        If oldSurface IsNot Nothing Then oldSurface.Dispose()
    End Sub

    Private Shared Sub AddReference(result As List(Of String), path As String)
        If String.IsNullOrWhiteSpace(path) OrElse
           result.Any(Function(item) String.Equals(item, path, StringComparison.OrdinalIgnoreCase)) Then Return
        result.Add(path)
    End Sub

    Private NotInheritable Class BridgeElevationProvider
        Inherits StateElevationProvider
        Implements ISurfaceContainer

        Private ReadOnly _surfaces As New List(Of Surface)()

        Public Sub New(root As InfrastradaBridgesModel)
            If root Is Nothing Then Return
            _surfaces.Add(root.Surface)
            If root.ModelFinder Is Nothing Then Return
            For Each path As String In root.GetReferences()
                Try
                    Dim container As ISurfaceContainer = root.ModelFinder.ReadModelFromPath(Of ISurfaceContainer)(path)
                    If container IsNot Nothing AndAlso container.Surface IsNot Nothing AndAlso
                       Not _surfaces.Contains(container.Surface) Then _surfaces.Add(container.Surface)
                Catch
                End Try
            Next
        End Sub

        Public Overrides Function GetElevations(point As Vector2D) As IEnumerable(Of StateElevationValue)
            Dim result As New List(Of StateElevationValue)()
            For Each item As Surface In _surfaces
                Dim elevation As Nullable(Of Double) = item.GetElevation(point)
                If elevation.HasValue Then
                    result.Add(New StateElevationValue With {
                        .Value = elevation.Value,
                        .State = If(item.Designed, StateElevationType.DESIGNED, StateElevationType.EXISTING)
                    })
                End If
            Next
            Return result
        End Function

        Public Overrides Function GetReferences() As IEnumerable(Of IElevationProvider)
            Return Enumerable.Empty(Of IElevationProvider)()
        End Function

        Private ReadOnly Property ProviderSurface As Surface Implements ISurfaceContainer.Surface
            Get
                Return _surfaces.FirstOrDefault()
            End Get
        End Property
    End Class
End Class
