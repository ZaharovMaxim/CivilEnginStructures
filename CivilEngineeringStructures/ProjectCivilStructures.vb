Imports System.IO
Imports System.Windows.Shapes
Imports Newtonsoft.Json.Linq
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Alg.Runtime.Tools
Imports Topomatic.Alg.Survey.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Imports Topomatic.Sites.Controller.Sheet.SummaryExtendedSheet.Frames
Imports Topomatic.Sites.Core

Public Class ProjectCivilStructures
    'Inherits ProjectModelStructures
    Const nameFolder As String = "/Модели/ИССО"
    Private _name As String
    Private _listModelStructures As List(Of ArrangementModel)
    Private _listModelSurfaces As List(Of TerrainModel)
    Private _listModelRoads As List(Of RoadModel)
    Private _listModelSites As List(Of SiteModel)
    Private _listModelSurvey As List(Of SurveyModel)
    Public Sub New()
        _name = ""
        _listModelStructures = New List(Of ArrangementModel)
        _listModelSurfaces = New List(Of TerrainModel)
        _listModelRoads = New List(Of RoadModel)
        _listModelSites = New List(Of SiteModel)
        _listModelSurvey = New List(Of SurveyModel)
        getIModels()
    End Sub
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    Public Property ListModelStructures As List(Of ArrangementModel)
        Get
            Return _listModelStructures
        End Get
        Set(value As List(Of ArrangementModel))
            _listModelStructures = value
        End Set
    End Property
    Public Property ListModelSurfaces As List(Of TerrainModel)
        Get
            Return _listModelSurfaces
        End Get
        Set(value As List(Of TerrainModel))
            _listModelSurfaces = value
        End Set
    End Property
    Public Property ListModelRoads As List(Of RoadModel)
        Get
            Return _listModelRoads
        End Get
        Set(value As List(Of RoadModel))
            _listModelRoads = value
        End Set
    End Property
    Public Property ListModelSites As List(Of SiteModel)
        Get
            Return _listModelSites
        End Get
        Set(value As List(Of SiteModel))
            _listModelSites = value
        End Set
    End Property

    Public Property ListModelSurvwy As List(Of SurveyModel)
        Get
            Return _listModelSurvey
        End Get
        Set(value As List(Of SurveyModel))
            _listModelSurvey = value
        End Set
    End Property
    '==========================================================================================================================
    'получить все модели проектов
    Public Sub getIModels()
        Try
            PluginCoreOps.FilterModels(Function(modelProjectChild As IProjectModel) As Boolean
                                           If IsNothing(modelProjectChild) = False Then
                                               Dim modelProjectUri As Topomatic.FoundationClasses.URI = modelProjectChild.Uri
                                               Dim modelPatch As String = modelProjectUri.AsFilePath
                                               Dim extension As String = IO.Path.GetExtension(modelPatch)
                                               If extension Like ".rbprojx" Then
                                                   _name = modelProjectUri.LastPathComponent
                                               End If
                                               If extension Like ".arrx" Or extension Like ".sfcx" Or extension Like ".roadx" Or extension Like ".site" Or extension Like ".algx" Then
                                                   Dim container As IDrawingContainer = TryCast(modelProjectChild.LockRead(), IDrawingContainer)
                                                   If TypeOf modelProjectChild.Model Is ArrangementModel Then
                                                       Dim boolFindFolder As Boolean = FuncFiles.IsPathContainsFolder(modelProjectUri.AsFilePath, nameFolder)
                                                       If boolFindFolder = True Then
                                                           Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                                                           _listModelStructures.Add(modelProjectChild.Model)
                                                       End If
                                                   ElseIf TypeOf modelProjectChild.Model Is TerrainModel Then
                                                       Dim terrModel As TerrainModel = modelProjectChild.Model
                                                       Dim surf As Surface = terrModel.Surface
                                                       If IsNothing(surf) = False Then
                                                           If surf.Triangles.Count > 0 Then
                                                               Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                                                               _listModelSurfaces.Add(modelProjectChild.Model)
                                                           End If
                                                       End If
                                                   ElseIf TypeOf modelProjectChild.Model Is RoadModel Then
                                                       Dim roadModel As RoadModel = modelProjectChild.Model
                                                       Dim align As Alignment = roadModel.Alignment
                                                       If IsNothing(align) = False Then
                                                           If align.Plan.CompoundLine.Length > 0 Then
                                                               Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                                                               _listModelRoads.Add(modelProjectChild.Model)
                                                           End If
                                                       End If
                                                   ElseIf TypeOf modelProjectChild.Model Is SiteModel Then
                                                       Dim siteModel As SiteModel = modelProjectChild.Model
                                                       If modelProjectUri.ToString.IndexOf(nameFolder) > -1 Then
                                                           _listModelSites.Add(modelProjectChild.Model)
                                                       End If
                                                   ElseIf TypeOf modelProjectChild.Model Is SurveyModel Then
                                                       Dim syrvModel As SurveyModel = modelProjectChild.Model
                                                       Dim align As Alignment = syrvModel.Alignment
                                                       If IsNothing(align) = False Then
                                                           If align.Plan.CompoundLine.Length > 0 Then
                                                               Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                                                               _listModelSurvey.Add(modelProjectChild.Model)
                                                           End If
                                                       End If
                                                   End If
                                               End If
                                           End If
                                           'Return True
                                       End Function)
        Catch ex As System.Exception
        End Try
    End Sub
    '==========================================================================================================================
    'имена проектов в список, для comboBox
    Public Function listNameArrangementModels() As List(Of String)
        Dim userList As List(Of String) = New List(Of String)
        If _listModelStructures.Count > 0 Then
            For Each mList As ArrangementModel In _listModelStructures
                Dim nameModel As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {mList})
                userList.Add(nameModel)
            Next
        End If
        Return userList
    End Function

    Public Function listNameTerrainModels() As List(Of String)
        Dim userList As List(Of String) = New List(Of String)
        If _listModelSurfaces.Count > 0 Then
            For Each mList As TerrainModel In _listModelSurfaces
                Dim nameModel As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {mList})
                userList.Add(nameModel)
            Next
        End If
        Return userList
    End Function

    Public Function listNameRoadModels() As List(Of String)
        Dim userList As List(Of String) = New List(Of String)
        If _listModelRoads.Count > 0 Then
            For Each mList As RoadModel In _listModelRoads
                Dim nameModel As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {mList})
                userList.Add(nameModel)
            Next
        End If
        Return userList
    End Function

    Public Function listNameSitesModels() As List(Of String)
        Dim userList As List(Of String) = New List(Of String)
        If _listModelSites.Count > 0 Then
            For Each mList As SiteModel In _listModelSites
                Dim nameModel As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {mList})
                userList.Add(nameModel)
            Next
        End If
        Return userList
    End Function
    '==========================================================================================================================
    'получить доступ к проекту по его индексу comboBox
    Public Function getArrangementModelByIndex(ByVal index As Integer) As ArrangementModel
        Dim projectModel As ArrangementModel = Nothing
        If _listModelStructures.Count > 0 Then
            If index < _listModelStructures.Count Then
                projectModel = _listModelStructures.Item(index)
            End If
        End If
        Return projectModel
    End Function

    'получить доступ к поверхности
    Public Function getTerrainMidelByIndex(ByVal index As Integer) As TerrainModel
        Dim terrainModel As TerrainModel = Nothing
        If _listModelSurfaces.Count > 0 Then
            If index < _listModelSurfaces.Count Then
                terrainModel = _listModelSurfaces.Item(index)
            End If
        End If
        Return terrainModel
    End Function
    'получить доступ к трассе
    Public Function getRoadModelByIndex(ByVal index As Integer) As RoadModel
        Dim roadModel As RoadModel = Nothing
        If _listModelRoads.Count > 0 Then
            If index < _listModelRoads.Count Then
                roadModel = _listModelRoads.Item(index)
            End If
        End If
        Return roadModel
    End Function

    'получить все площадки
    Public Function getAlignSurveyDrawingByIndex(ByVal index As Integer) As SurveyModel
        Dim surveyModel As SurveyModel = Nothing
        If _listModelRoads.Count > 0 Then
            If index < _listModelRoads.Count Then
                surveyModel = _listModelSurvey.Item(index)
            End If
        End If
        Return surveyModel
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать новый проект arrx
    Public Function createArrangementModel() As ArrangementModel
        Dim folder As String = PluginCoreOps.FindModelPathId(PluginCoreOps.CreateFolder(New String() {"Модели", "ИССО", "Путепроводы"}))
        Try
            Dim iUserProject As IProjectModel = ApplicationHost.Current.Plugins.Execute("mkitem", New Object() {folder, "arr"})
            If IsNothing(iUserProject) = False Then
                Dim boolCreate As Boolean = ApplicationHost.Current.Plugins.Execute("activate", New Object() {iUserProject})
                Dim modelProjectUri As Topomatic.FoundationClasses.URI = iUserProject.Uri
                Dim boolFindFolder As Boolean = FuncFiles.IsPathContainsFolder(modelProjectUri.AsFilePath, nameFolder)
                If boolFindFolder = True Then
                    Dim nameProject As String = IO.Path.GetFileNameWithoutExtension(modelProjectUri.LastPathComponent)
                    _listModelStructures.Add(iUserProject.Model)
                    Return iUserProject.Model
                End If
            End If
        Catch ex As System.OperationCanceledException
        Catch ex As System.Exception
        End Try
        Return Nothing
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'удалить проект
    Public Function removeArrangementModel(ByVal nameModel As String) As Boolean
        If _listModelStructures.Count > 0 Then
            For Each modelStructure As ArrangementModel In _listModelStructures
                Dim nameRemoveProject As String = ApplicationHost.Current.Plugins.Execute("getname", New Object() {modelStructure})
                If nameRemoveProject Like nameModel Then
                    ApplicationHost.Current.Plugins.Execute("rmitem", New Object() {modelStructure})
                End If
            Next
        End If
        Return True
    End Function
End Class
