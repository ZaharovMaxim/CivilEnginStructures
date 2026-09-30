Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Alg.Survey.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
'Imports Topomatic.Sfc
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Cogo
Imports Topomatic.Cad.View
Imports Topomatic.Controls.Dialogs
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
'Imports Topomatic.Srv
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Planchet.Entities
Imports Topomatic.Sfc
Imports Topomatic.Sfc.Layer
Imports Topomatic.Visualization.Geometry
Public Class FuncSurface
    Public Shared Function getSurfaceChoices(ownerModel As ArrangementModel) As List(Of SurfaceModelChoice)
        Dim result As New List(Of SurfaceModelChoice)()
        If ownerModel Is Nothing OrElse ownerModel.ModelFinder Is Nothing Then Return result
        PluginCoreOps.FilterModels(
            Function(child As IProjectModel) As Boolean
                If child Is Nothing OrElse child.Uri Is Nothing OrElse
                   Not SurfaceModelChoice.IsSupportedSurfacePath(child.Uri.LastPathComponent) Then Return False
                If Object.ReferenceEquals(BridgeModelRuntime.GetArrangement(child.Model), ownerModel) Then Return False
                Dim container As ISurfaceContainer = PluginCoreOps.LockReadContainer(Of ISurfaceContainer)(child)
                If container Is Nothing OrElse container.Surface Is Nothing Then Return False
                Dim relativePath As String = ownerModel.ModelFinder.FindRelativePath(container.Surface)
                If String.IsNullOrWhiteSpace(relativePath) Then Return False
                If result.Any(Function(choice) String.Equals(choice.RelativePath, relativePath, StringComparison.OrdinalIgnoreCase)) Then Return False
                Dim modelName As String = Path.GetFileNameWithoutExtension(child.Uri.LastPathComponent)
                result.Add(New SurfaceModelChoice(modelName, child.ModelType, relativePath))
                Return False
            End Function)
        Return result
    End Function

    Public Shared Sub ConfigureSurfaceCombo(combo As ComboBox, ownerModel As ArrangementModel, storedValue As String)
        If combo Is Nothing Then Return
        Dim choices As List(Of SurfaceModelChoice) = getSurfaceChoices(ownerModel)
        choices.Insert(0, New SurfaceModelChoice("(не назначена)", String.Empty, String.Empty))
        Dim selected As SurfaceModelChoice = If(String.IsNullOrWhiteSpace(storedValue),
                                                 choices(0),
                                                 SurfaceModelChoice.FindByStoredValue(choices, storedValue))
        If selected Is Nothing AndAlso Not String.IsNullOrWhiteSpace(storedValue) Then
            selected = New SurfaceModelChoice("(недоступна)", String.Empty, storedValue)
            choices.Add(selected)
        End If
        combo.DataSource = Nothing
        combo.DisplayMember = NameOf(SurfaceModelChoice.DisplayName)
        combo.ValueMember = NameOf(SurfaceModelChoice.RelativePath)
        combo.DataSource = choices
        If selected IsNot Nothing Then combo.SelectedItem = selected
    End Sub

    Public Shared Sub SelectSurfaceChoice(combo As ComboBox, storedValue As String)
        If combo Is Nothing Then Return
        Dim source As IEnumerable(Of SurfaceModelChoice) = TryCast(combo.DataSource, IEnumerable(Of SurfaceModelChoice))
        Dim choices As List(Of SurfaceModelChoice) = If(source Is Nothing,
                                                         New List(Of SurfaceModelChoice)(),
                                                         New List(Of SurfaceModelChoice)(source))
        Dim selected As SurfaceModelChoice = SurfaceModelChoice.FindByStoredValue(choices, storedValue)
        If selected Is Nothing AndAlso String.IsNullOrWhiteSpace(storedValue) Then
            selected = choices.FirstOrDefault(Function(choice) String.IsNullOrWhiteSpace(choice.RelativePath))
        End If
        If selected Is Nothing AndAlso Not String.IsNullOrWhiteSpace(storedValue) Then
            selected = New SurfaceModelChoice("(недоступна)", String.Empty, storedValue)
            choices.Add(selected)
            combo.DataSource = Nothing
            combo.DisplayMember = NameOf(SurfaceModelChoice.DisplayName)
            combo.ValueMember = NameOf(SurfaceModelChoice.RelativePath)
            combo.DataSource = choices
        End If
        combo.SelectedItem = selected
    End Sub

    Public Shared Function getSelectedSurfaceReference(combo As ComboBox) As String
        If combo Is Nothing Then Return String.Empty
        Dim choice As SurfaceModelChoice = TryCast(combo.SelectedItem, SurfaceModelChoice)
        If choice IsNot Nothing Then Return choice.RelativePath
        Return combo.Text
    End Function

    Public Shared Function resolveSurface(ownerModel As ArrangementModel, storedValue As String) As Surface
        If String.IsNullOrWhiteSpace(storedValue) Then Return Nothing
        Dim normalizedValue As String = storedValue.Trim()
        If ownerModel IsNot Nothing AndAlso ownerModel.ModelFinder IsNot Nothing Then
            Try
                Dim container As ISurfaceContainer = ownerModel.ModelFinder.ReadModelFromPath(Of ISurfaceContainer)(normalizedValue)
                If container IsNot Nothing AndAlso container.Surface IsNot Nothing Then Return container.Surface
            Catch ex As System.Exception
            End Try
        End If
        If IsExactSurfaceReference(normalizedValue) Then Return Nothing
        Return getSurfaceByName(normalizedValue)
    End Function

    Public Shared Function requireSurface(ownerModel As ArrangementModel,
                                          storedValue As String,
                                          valueName As String) As Surface
        Dim displayName As String = If(String.IsNullOrWhiteSpace(valueName), "Поверхность", valueName.Trim())
        If String.IsNullOrWhiteSpace(storedValue) Then
            Throw New InvalidOperationException(displayName & " не назначена.")
        End If

        Dim normalizedValue As String = storedValue.Trim()
        If IsExactSurfaceReference(normalizedValue) Then
            If ownerModel Is Nothing OrElse ownerModel.ModelFinder Is Nothing Then
                Throw New InvalidOperationException(displayName & " недоступна: " & normalizedValue)
            End If
            Try
                Dim container As ISurfaceContainer = ownerModel.ModelFinder.ReadModelFromPath(Of ISurfaceContainer)(normalizedValue)
                If container IsNot Nothing AndAlso container.Surface IsNot Nothing Then Return container.Surface
            Catch ex As Exception
                Throw New InvalidOperationException(
                    displayName & " не открыта по сохранённому пути: " & normalizedValue,
                    ex)
            End Try
            Throw New InvalidOperationException(
                displayName & " не содержит доступную поверхность: " & normalizedValue)
        End If

        Dim legacySurface As Surface = getSurfaceByName(normalizedValue)
        If legacySurface Is Nothing Then
            Throw New InvalidOperationException(displayName & " не найдена: " & normalizedValue)
        End If
        Return legacySurface
    End Function

    Private Shared Function IsExactSurfaceReference(value As String) As Boolean
        Return value.IndexOf("\"c) >= 0 OrElse
               value.IndexOf("/"c) >= 0 OrElse
               SurfaceModelChoice.IsSupportedSurfacePath(value)
    End Function

    'функция получает все поверхности проекта
    Public Shared Function getSurfaces() As Dictionary(Of String, Surface)
        Dim result As Dictionary(Of String, Surface) = New Dictionary(Of String, Surface)
        PluginCoreOps.FilterModels(
            Function(child As IProjectModel) As Boolean
                If child IsNot Nothing Then
                    Dim modelUri As URI = child.Uri
                    If modelUri.Extension Like ".sfcx" OrElse modelUri.Extension Like ".algx" OrElse modelUri.Extension Like ".roadx" Then
                        Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                        Dim container As ISurfaceContainer = PluginCoreOps.LockReadContainer(Of ISurfaceContainer)(child)
                        If container IsNot Nothing AndAlso container.Surface IsNot Nothing AndAlso Not result.ContainsKey(fileNameModel) Then
                            result.Add(fileNameModel, container.Surface)
                        End If
                    End If
                End If
                Return False
            End Function)
        Return result
    End Function
    'функция возвращает поверхность по ее имени
    Public Shared Function getSurfaceByName(ByVal nameSurface As String) As Surface
        getSurfaceByName = Nothing
        Try
            Dim result As Surface = Nothing
            PluginCoreOps.FilterModels(
                Function(child As IProjectModel) As Boolean
                    If result Is Nothing AndAlso child IsNot Nothing Then
                        Dim modelUri As URI = child.Uri
                        If modelUri.Extension Like ".sfcx" OrElse modelUri.Extension Like ".algx" OrElse modelUri.Extension Like ".roadx" Then
                            Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                            If nameSurface.Trim Like fileNameModel.Trim Then
                                Dim container As ISurfaceContainer = PluginCoreOps.LockReadContainer(Of ISurfaceContainer)(child)
                                If container IsNot Nothing Then result = container.Surface
                            End If
                        End If
                    End If
                    Return False
                End Function)
            Return result
        Catch ex As System.Exception
        End Try
    End Function
    'функция возвращает высоту точки на поверхности
    Public Shared Function getElevationToSurface(ByVal surface As Surface, ByVal point As Vector2D, ByRef elevation As Double) As Boolean
        If surface Is Nothing Then Return False

        Try
            Dim surfaceElevation As Nullable(Of Double) = surface.GetElevation(point)
            If surfaceElevation.HasValue AndAlso
               Not Double.IsNaN(surfaceElevation.Value) AndAlso
               Not Double.IsInfinity(surfaceElevation.Value) Then
                elevation = surfaceElevation.Value
                Return True
            End If
        Catch ex As System.NullReferenceException
        Catch ex As System.InvalidOperationException
        End Try

        Try
            Dim triangles As SurfaceTriangleArray = surface.Triangles
            Dim points As SurfacePointArray = surface.Points

            If triangles IsNot Nothing AndAlso triangles.Count > 0 Then
                Dim triangleIndex As Integer = -1
                Try
                    triangleIndex = surface.FindTriangle(point)
                Catch ex As System.NullReferenceException
                Catch ex As System.InvalidOperationException
                Catch ex As System.ArgumentOutOfRangeException
                Catch ex As System.IndexOutOfRangeException
                End Try

                Dim a As Vector3D
                Dim b As Vector3D
                Dim c As Vector3D
                If TryGetTriangleVertices(triangles, points, triangleIndex, a, b, c) Then
                    Dim indexedElevation As Double = 0.0
                    Dim indexedSquaredDistance As Double = 0.0
                    If SurfaceElevationGeometry.TryGetElevationOnTriangle(point,
                                                                           a,
                                                                           b,
                                                                           c,
                                                                           False,
                                                                           indexedElevation,
                                                                           indexedSquaredDistance) Then
                        elevation = indexedElevation
                        Return True
                    End If
                End If
            End If

            Dim hasClosestTriangle As Boolean = False
            Dim closestElevation As Double = 0.0
            Dim closestSquaredDistance As Double = Double.MaxValue

            If triangles IsNot Nothing Then
                For triangleIndex As Integer = 0 To triangles.Count - 1
                    Dim a As Vector3D
                    Dim b As Vector3D
                    Dim c As Vector3D
                    If TryGetTriangleVertices(triangles, points, triangleIndex, a, b, c) Then
                        Dim candidateElevation As Double = 0.0
                        Dim candidateSquaredDistance As Double = 0.0

                        If SurfaceElevationGeometry.TryGetElevationOnTriangle(point,
                                                                               a,
                                                                               b,
                                                                               c,
                                                                               True,
                                                                               candidateElevation,
                                                                               candidateSquaredDistance) Then
                            If candidateSquaredDistance = 0.0 Then
                                elevation = candidateElevation
                                Return True
                            End If

                            If Not hasClosestTriangle OrElse candidateSquaredDistance < closestSquaredDistance Then
                                closestElevation = candidateElevation
                                closestSquaredDistance = candidateSquaredDistance
                                hasClosestTriangle = True
                            End If
                        End If
                    End If
                Next
            End If

            If hasClosestTriangle Then
                elevation = closestElevation
                Return True
            End If

            Dim vertices As New System.Collections.Generic.List(Of Vector3D)
            If points IsNot Nothing Then
                For pointIndex As Integer = 0 To points.Count - 1
                    Dim surfacePoint As SurfacePoint = points(pointIndex)
                    If Not surfacePoint.IsRemoved Then vertices.Add(surfacePoint.Vertex)
                Next
            End If

            Dim nearestElevation As Double = 0.0
            If SurfaceElevationGeometry.TryGetNearestSurfacePointElevation(point, vertices, nearestElevation) Then
                elevation = nearestElevation
                Return True
            End If
        Catch ex As System.NullReferenceException
        Catch ex As System.InvalidOperationException
        Catch ex As System.ArgumentOutOfRangeException
        Catch ex As System.IndexOutOfRangeException
        End Try

        Return False
    End Function

    Private Shared Function TryGetTriangleVertices(ByVal triangles As SurfaceTriangleArray,
                                                   ByVal points As SurfacePointArray,
                                                   ByVal triangleIndex As Integer,
                                                   ByRef a As Vector3D,
                                                   ByRef b As Vector3D,
                                                   ByRef c As Vector3D) As Boolean
        If triangles Is Nothing OrElse points Is Nothing OrElse
           triangleIndex < 0 OrElse triangleIndex >= triangles.Count Then
            Return False
        End If

        Dim triangle As SurfaceTriangle = triangles(triangleIndex)
        If triangle.IsRemoved OrElse
           triangle.A < 0 OrElse triangle.A >= points.Count OrElse
           triangle.B < 0 OrElse triangle.B >= points.Count OrElse
           triangle.C < 0 OrElse triangle.C >= points.Count Then
            Return False
        End If

        Dim pointA As SurfacePoint = points(triangle.A)
        Dim pointB As SurfacePoint = points(triangle.B)
        Dim pointC As SurfacePoint = points(triangle.C)
        If pointA.IsRemoved OrElse pointB.IsRemoved OrElse pointC.IsRemoved Then Return False

        a = pointA.Vertex
        b = pointB.Vertex
        c = pointC.Vertex
        Return True
    End Function

End Class
