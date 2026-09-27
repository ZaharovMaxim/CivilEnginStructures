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
    'функция получает все поверхности проекта
    Public Shared Function getSurfaces() As Dictionary(Of String, Surface)
        Dim result As Dictionary(Of String, Surface) = New Dictionary(Of String, Surface)
        Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
        Dim childs As IProjectModel() = Project.Model.GetChilds()
        For Each child As IProjectModel In childs
            Dim modelUri As URI = child.Uri
            If modelUri.Extension Like ".sfcx" Then
                Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                Dim userTerrainModel As TerrainModel = child.Model
                result.Add(fileNameModel, userTerrainModel.Surface)
            ElseIf modelUri.Extension Like ".algx" Then
                Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                Dim survModel As SurveyModel = child.Model
                result.Add(fileNameModel, survModel.Surface)
            End If
        Next
        Return result
    End Function
    'функция возвращает поверхность по ее имени
    Public Shared Function getSurfaceByName(ByVal nameSurface As String) As Surface
        getSurfaceByName = Nothing
        Try
            Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
            Dim childs As IProjectModel() = Project.Model.GetChilds()
            For Each child As IProjectModel In childs
                Dim modelUri As URI = child.Uri
                If modelUri.Extension Like ".sfcx" Then
                    Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If nameSurface.Trim Like fileNameModel.Trim Then
                        Dim userTerrainModel As TerrainModel = child.Model
                        If IsNothing(userTerrainModel) = False Then
                            Return userTerrainModel.Surface
                        End If
                    End If
                ElseIf modelUri.Extension Like ".algx" Then
                    Dim fileNameModel As String = Path.GetFileNameWithoutExtension(modelUri.LastPathComponent)
                    If nameSurface.Trim Like fileNameModel.Trim Then
                        Dim survModel As SurveyModel = child.Model
                        If IsNothing(survModel) = False Then
                            Return survModel.Surface
                        End If
                    End If
                End If
            Next
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
