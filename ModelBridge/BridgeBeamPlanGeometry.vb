Imports System.Collections.Generic
Imports System.Linq
Imports NetTopologySuite.Geometries
Imports Topomatic.Cad.Foundation

Public NotInheritable Class BridgeBeamPlanGeometry
    Private Sub New()
    End Sub

    Public Shared Function BuildOutline(top As IEnumerable(Of Vector2D),
                                        bottom As IEnumerable(Of Vector2D)) As Vector2D()
        Dim geometryFactory As New GeometryFactory()
        Dim topPolygon As Polygon = CreatePolygon(geometryFactory, top, "верхний")
        Dim bottomPolygon As Polygon = CreatePolygon(geometryFactory, bottom, "нижний")

        Dim unionGeometry As Geometry
        Try
            unionGeometry = topPolygon.Union(bottomPolygon)
        Catch exception As Exception
            Throw New InvalidOperationException("Не удалось объединить контуры балки: " & exception.Message, exception)
        End Try

        Dim unionPolygon As Polygon = TryCast(unionGeometry, Polygon)
        If unionPolygon Is Nothing OrElse unionPolygon.IsEmpty Then
            Throw New InvalidOperationException("Контуры балки должны образовывать один связный полигон.")
        End If
        If unionPolygon.NumInteriorRings <> 0 Then
            Throw New InvalidOperationException("Объединённый контур балки содержит внутренние отверстия.")
        End If
        If Not unionPolygon.IsValid OrElse unionPolygon.Area <= 0.0R Then
            Throw New InvalidOperationException("Объединённый контур балки геометрически некорректен.")
        End If

        Dim outline As List(Of Vector2D) = unionPolygon.ExteriorRing.Coordinates.
            Select(Function(coordinate) New Vector2D(coordinate.X, coordinate.Y)).ToList()
        RemoveRedundantVertices(outline)
        If outline.Count < 4 Then
            Throw New InvalidOperationException("Объединённый контур балки не содержит допустимой границы.")
        End If
        If SignedArea(outline) < 0.0R Then
            outline.Reverse()
        End If
        EnsureClosed(outline)
        Return outline.ToArray()
    End Function

    Private Shared Function CreatePolygon(factory As GeometryFactory,
                                          source As IEnumerable(Of Vector2D),
                                          caption As String) As Polygon
        If source Is Nothing Then
            Throw New ArgumentNullException(NameOf(source), "Не задан " & caption & " контур балки.")
        End If

        Dim points As New List(Of Vector2D)()
        For Each pointValue As Vector2D In source
            If Not IsFinite(pointValue.X) OrElse Not IsFinite(pointValue.Y) Then
                Throw New ArgumentException("Координаты " & caption & " контура балки должны быть конечными.", NameOf(source))
            End If
            If points.Count = 0 OrElse Not SamePoint(points(points.Count - 1), pointValue) Then
                points.Add(pointValue)
            End If
        Next
        If points.Count > 1 AndAlso SamePoint(points(0), points(points.Count - 1)) Then
            points.RemoveAt(points.Count - 1)
        End If
        If points.Count < 3 Then
            Throw New ArgumentException("В " & caption & " контуре балки должно быть не менее трёх вершин.", NameOf(source))
        End If

        Dim coordinates(points.Count) As Coordinate
        For index As Integer = 0 To points.Count - 1
            coordinates(index) = New Coordinate(points(index).X, points(index).Y)
        Next
        coordinates(points.Count) = New Coordinate(points(0).X, points(0).Y)

        Dim polygon As Polygon
        Try
            polygon = factory.CreatePolygon(factory.CreateLinearRing(coordinates))
        Catch exception As Exception
            Throw New ArgumentException("Не удалось создать " & caption & " контур балки: " & exception.Message,
                                        NameOf(source), exception)
        End Try
        If polygon.IsEmpty OrElse polygon.Area <= 0.0R Then
            Throw New ArgumentException(FirstUpper(caption) & " контур балки имеет нулевую площадь.", NameOf(source))
        End If
        If Not polygon.IsValid Then
            Throw New ArgumentException(FirstUpper(caption) & " контур балки самопересекается или геометрически некорректен.",
                                        NameOf(source))
        End If
        Return polygon
    End Function

    Private Shared Sub RemoveRedundantVertices(points As List(Of Vector2D))
        If points.Count > 1 AndAlso SamePoint(points(0), points(points.Count - 1)) Then
            points.RemoveAt(points.Count - 1)
        End If
        Dim changed As Boolean = True
        While changed AndAlso points.Count > 3
            changed = False
            For index As Integer = points.Count - 1 To 0 Step -1
                Dim previousIndex As Integer = (index + points.Count - 1) Mod points.Count
                Dim nextIndex As Integer = (index + 1) Mod points.Count
                If SamePoint(points(previousIndex), points(index)) OrElse
                   IsRedundantCollinear(points(previousIndex), points(index), points(nextIndex)) Then
                    points.RemoveAt(index)
                    changed = True
                End If
            Next
        End While
        EnsureClosed(points)
    End Sub

    Private Shared Function IsRedundantCollinear(previousPoint As Vector2D,
                                                  currentPoint As Vector2D,
                                                  nextPoint As Vector2D) As Boolean
        Dim firstX As Double = currentPoint.X - previousPoint.X
        Dim firstY As Double = currentPoint.Y - previousPoint.Y
        Dim secondX As Double = nextPoint.X - currentPoint.X
        Dim secondY As Double = nextPoint.Y - currentPoint.Y
        Dim cross As Double = firstX * secondY - firstY * secondX
        Dim scale As Double = Math.Max(1.0R,
                                      Math.Sqrt(firstX * firstX + firstY * firstY) *
                                      Math.Sqrt(secondX * secondX + secondY * secondY))
        If Math.Abs(cross) > scale * 0.0000000001R Then Return False
        Return firstX * secondX + firstY * secondY >= 0.0R
    End Function

    Private Shared Sub EnsureClosed(points As List(Of Vector2D))
        If points.Count > 0 AndAlso Not SamePoint(points(0), points(points.Count - 1)) Then
            points.Add(points(0))
        End If
    End Sub

    Private Shared Function SignedArea(points As IList(Of Vector2D)) As Double
        If points.Count < 3 Then Return 0.0R
        Dim origin As Vector2D = points(0)
        Dim twiceArea As Double = 0.0R
        For index As Integer = 0 To points.Count - 2
            Dim currentX As Double = points(index).X - origin.X
            Dim currentY As Double = points(index).Y - origin.Y
            Dim nextX As Double = points(index + 1).X - origin.X
            Dim nextY As Double = points(index + 1).Y - origin.Y
            twiceArea += currentX * nextY - nextX * currentY
        Next
        Return twiceArea / 2.0R
    End Function

    Private Shared Function SamePoint(first As Vector2D, second As Vector2D) As Boolean
        Return first.X = second.X AndAlso first.Y = second.Y
    End Function

    Private Shared Function IsFinite(value As Double) As Boolean
        Return Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value)
    End Function

    Private Shared Function FirstUpper(value As String) As String
        If String.IsNullOrEmpty(value) Then Return value
        Return Char.ToUpperInvariant(value(0)) & value.Substring(1)
    End Function
End Class
