Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Net.Security
Imports System.Security.AccessControl
Imports System.Security.RightsManagement
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.Alg.Prf
Imports Topomatic.Alg.Road.Core
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Sfc
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D

Public Class Triangle
    Public Property Vertices As Integer()
    Public Property Circumcenter As Vector2D
    Public Property Circumradius As Double

    Public Sub New(v1 As Integer, v2 As Integer, v3 As Integer)
        Vertices = {v1, v2, v3}
        Array.Sort(Vertices)
    End Sub

    Public Function ContainsVertex(vertexIndex As Integer) As Boolean
        Return Vertices.Contains(vertexIndex)
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Not TypeOf obj Is Triangle Then Return False
        Dim other As Triangle = DirectCast(obj, Triangle)
        Return Vertices(0) = other.Vertices(0) AndAlso
               Vertices(1) = other.Vertices(1) AndAlso
               Vertices(2) = other.Vertices(2)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return Vertices(0) Xor Vertices(1) Xor Vertices(2)
    End Function

    Public Overrides Function ToString() As String
        Return $"[{Vertices(0)}, {Vertices(1)}, {Vertices(2)}]"
    End Function
End Class

Public Class Edge
    Public Property StartVertex As Integer
    Public Property EndVertex As Integer

    Public Sub New(start As Integer, [end] As Integer)
        StartVertex = Math.Min(start, [end])
        EndVertex = Math.Max(start, [end])
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing OrElse Not TypeOf obj Is Edge Then Return False
        Dim other As Edge = DirectCast(obj, Edge)
        Return StartVertex = other.StartVertex AndAlso EndVertex = other.EndVertex
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return StartVertex Xor EndVertex
    End Function

    Public Overrides Function ToString() As String
        Return $"[{StartVertex}-{EndVertex}]"
    End Function
End Class

Public Class DelaunayTriangulation
    Private Points As List(Of Vector2D)
    Private Triangles As HashSet(Of Triangle)

    Public Sub New()
        Points = New List(Of Vector2D)()
        Triangles = New HashSet(Of Triangle)()
    End Sub

    ''' <summary>
    ''' Выполняет триангуляцию Делоне по алгоритму Боуера-Ватсона
    ''' </summary>
    ''' <returns>Словарь: ключ - номер треугольника, значение - координаты трех вершин</returns>
    Public Function Triangulate(points As List(Of Vector2D)) As Dictionary(Of Integer, List(Of Vector2D))
        If points Is Nothing OrElse points.Count < 3 Then
            Throw New ArgumentException("Для триангуляции необходимо как минимум 3 точки")
        End If

        Me.Points = New List(Of Vector2D)(points)
        Triangles.Clear()

        ' Создаем супер-треугольник, содержащий все точки
        Dim superTriangle As Triangle = CreateSuperTriangle()
        Triangles.Add(superTriangle)

        ' Добавляем точки по одной
        For pointIndex As Integer = 0 To points.Count - 1
            AddPoint(pointIndex)
        Next

        ' Удаляем треугольники, содержащие вершины супер-треугольника
        RemoveSuperTriangle(superTriangle)

        Return GetTriangulationResult()
    End Function

    ''' <summary>
    ''' Создает супер-треугольник, содержащий все точки
    ''' </summary>
    Private Function CreateSuperTriangle() As Triangle
        Dim minX As Double = Points.Min(Function(p) p.X)
        Dim maxX As Double = Points.Max(Function(p) p.X)
        Dim minY As Double = Points.Min(Function(p) p.Y)
        Dim maxY As Double = Points.Max(Function(p) p.Y)

        Dim dx As Double = maxX - minX
        Dim dy As Double = maxY - minY
        Dim deltaMax As Double = Math.Max(dx, dy) * 2

        Dim p1 As New Vector2D(minX - deltaMax, minY - deltaMax)
        Dim p2 As New Vector2D(maxX + deltaMax, minY - deltaMax)
        Dim p3 As New Vector2D(minX + dx / 2, maxY + deltaMax)

        ' Добавляем вершины супер-треугольника в список точек
        Points.Add(p1)
        Points.Add(p2)
        Points.Add(p3)

        Return New Triangle(Points.Count - 3, Points.Count - 2, Points.Count - 1)
    End Function

    ''' <summary>
    ''' Добавляет точку в триангуляцию
    ''' </summary>
    Private Sub AddPoint(pointIndex As Integer)
        Dim badTriangles As New HashSet(Of Triangle)()
        Dim polygonEdges As New HashSet(Of Edge)()

        ' Находим все треугольники, чья описанная окружность содержит новую точку
        For Each triangle In Triangles
            If IsPointInCircumcircle(Points(pointIndex), triangle) Then
                badTriangles.Add(triangle)

                ' Добавляем ребра плохого треугольника
                For Each edge In GetTriangleEdges(triangle)
                    If polygonEdges.Contains(edge) Then
                        polygonEdges.Remove(edge)
                    Else
                        polygonEdges.Add(edge)
                    End If
                Next
            End If
        Next

        ' Удаляем плохие треугольники
        For Each badTriangle In badTriangles
            Triangles.Remove(badTriangle)
        Next

        ' Создаем новые треугольники из ребер полигона и новой точки
        For Each edge In polygonEdges
            Dim newTriangle As New Triangle(edge.StartVertex, edge.EndVertex, pointIndex)
            CalculateCircumcircle(newTriangle)
            Triangles.Add(newTriangle)
        Next
    End Sub

    ''' <summary>
    ''' Проверяет, находится ли точка внутри описанной окружности треугольника
    ''' </summary>
    Private Function IsPointInCircumcircle(point As Vector2D, triangle As Triangle, Optional includeBoundary As Boolean = True) As Boolean
        If triangle.Circumradius = 0 Then
            CalculateCircumcircle(triangle)
        End If

        Dim dx As Double = point.X - triangle.Circumcenter.X
        Dim dy As Double = point.Y - triangle.Circumcenter.Y
        Dim distanceSquared As Double = dx * dx + dy * dy
        Dim radiusSquared As Double = triangle.Circumradius * triangle.Circumradius

        If includeBoundary Then
            Return distanceSquared <= radiusSquared
        End If

        Dim tolerance As Double = radiusSquared * 0.000000000001
        Return distanceSquared < radiusSquared - tolerance
    End Function

    ''' <summary>
    ''' Вычисляет описанную окружность для треугольника
    ''' </summary>
    Private Sub CalculateCircumcircle(triangle As Triangle)
        Dim A As Vector2D = Points(triangle.Vertices(0))
        Dim B As Vector2D = Points(triangle.Vertices(1))
        Dim C As Vector2D = Points(triangle.Vertices(2))

        Dim bx As Double = B.X - A.X
        Dim by As Double = B.Y - A.Y
        Dim cx As Double = C.X - A.X
        Dim cy As Double = C.Y - A.Y
        Dim D As Double = 2 * (bx * cy - by * cx)
        Dim coordinateScale As Double = Math.Max(Math.Max(Math.Abs(bx), Math.Abs(by)), Math.Max(Math.Abs(cx), Math.Abs(cy)))

        If coordinateScale = 0 OrElse Math.Abs(D) < coordinateScale * coordinateScale * 0.000000000001 Then
            ' Точки коллинеарны
            triangle.Circumcenter = New Vector2D(0, 0)
            triangle.Circumradius = 0
            Return
        End If

        Dim B_sq As Double = bx * bx + by * by
        Dim C_sq As Double = cx * cx + cy * cy

        Dim Ux As Double = A.X + (B_sq * cy - C_sq * by) / D
        Dim Uy As Double = A.Y + (C_sq * bx - B_sq * cx) / D

        triangle.Circumcenter = New Vector2D(Ux, Uy)

        Dim dx As Double = A.X - Ux
        Dim dy As Double = A.Y - Uy
        triangle.Circumradius = Math.Sqrt(dx * dx + dy * dy)
    End Sub

    ''' <summary>
    ''' Получает ребра треугольника
    ''' </summary>
    Private Function GetTriangleEdges(triangle As Triangle) As List(Of Edge)
        Dim edges As New List(Of Edge)()
        edges.Add(New Edge(triangle.Vertices(0), triangle.Vertices(1)))
        edges.Add(New Edge(triangle.Vertices(1), triangle.Vertices(2)))
        edges.Add(New Edge(triangle.Vertices(2), triangle.Vertices(0)))
        Return edges
    End Function

    ''' <summary>
    ''' Удаляет треугольники, связанные с супер-треугольником
    ''' </summary>
    Private Sub RemoveSuperTriangle(superTriangle As Triangle)
        Dim trianglesToRemove As New List(Of Triangle)()

        For Each triangle In Triangles
            For Each vertex In triangle.Vertices
                If superTriangle.ContainsVertex(vertex) Then
                    trianglesToRemove.Add(triangle)
                    Exit For
                End If
            Next
        Next

        For Each triangle In trianglesToRemove
            Triangles.Remove(triangle)
        Next

        ' Удаляем вершины супер-треугольника из списка точек
        Points.RemoveRange(Points.Count - 3, 3)
    End Sub

    ''' <summary>
    ''' Формирует результат триангуляции в требуемом формате
    ''' </summary>
    ''' <returns>Словарь: ключ - номер треугольника, значение - координаты трех вершин</returns>
    Private Function GetTriangulationResult() As Dictionary(Of Integer, List(Of Vector2D))
        Dim result As New Dictionary(Of Integer, List(Of Vector2D))()

        Dim triangleIndex As Integer = 0
        For Each triangle In Triangles
            Dim triangleVertices As New List(Of Vector2D)()

            ' Получаем координаты вершин треугольника
            For Each vertexIndex In triangle.Vertices
                triangleVertices.Add(Points(vertexIndex))
            Next

            result.Add(triangleIndex, triangleVertices)
            triangleIndex += 1
        Next

        Return result
    End Function

    ''' <summary>
    ''' Проверяет корректность триангуляции Делоне
    ''' </summary>
    Public Function VerifyDelaunay() As Boolean
        For Each triangle In Triangles
            For i As Integer = 0 To Points.Count - 1
                If Not triangle.ContainsVertex(i) Then
                    If IsPointInCircumcircle(Points(i), triangle, False) Then
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function
End Class
