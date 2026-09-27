Imports System.Collections.Generic
Imports Topomatic.Cad.Foundation

Friend NotInheritable Class SurfaceElevationGeometry
    Private Const GeometryTolerance As Double = 0.000000000001

    Private Sub New()
    End Sub

    Friend Shared Function TryGetElevationOnTriangle(ByVal point As Vector2D,
                                                       ByVal a As Vector3D,
                                                       ByVal b As Vector3D,
                                                       ByVal c As Vector3D,
                                                       ByVal clampToTriangle As Boolean,
                                                       ByRef elevation As Double,
                                                       ByRef squaredDistance As Double) As Boolean
        If Not IsFinite(point.X) OrElse Not IsFinite(point.Y) OrElse
           Not IsFinite(a.X) OrElse Not IsFinite(a.Y) OrElse Not IsFinite(a.Z) OrElse
           Not IsFinite(b.X) OrElse Not IsFinite(b.Y) OrElse Not IsFinite(b.Z) OrElse
           Not IsFinite(c.X) OrElse Not IsFinite(c.Y) OrElse Not IsFinite(c.Z) Then
            Return False
        End If

        Dim abX As Double = b.X - a.X
        Dim abY As Double = b.Y - a.Y
        Dim acX As Double = c.X - a.X
        Dim acY As Double = c.Y - a.Y
        Dim bcX As Double = c.X - b.X
        Dim bcY As Double = c.Y - b.Y
        Dim determinant As Double = abX * acY - abY * acX
        Dim maximumEdgeLengthSquared As Double = Math.Max(abX * abX + abY * abY,
                                                          Math.Max(acX * acX + acY * acY,
                                                                   bcX * bcX + bcY * bcY))

        If Not IsFinite(determinant) OrElse Not IsFinite(maximumEdgeLengthSquared) OrElse
           maximumEdgeLengthSquared <= 0.0 OrElse
           Math.Abs(determinant) <= maximumEdgeLengthSquared * GeometryTolerance Then
            Return False
        End If

        Dim apX As Double = point.X - a.X
        Dim apY As Double = point.Y - a.Y
        Dim weightB As Double = (apX * acY - apY * acX) / determinant
        Dim weightC As Double = (abX * apY - abY * apX) / determinant
        Dim weightA As Double = 1.0 - weightB - weightC

        If Not IsFinite(weightA) OrElse Not IsFinite(weightB) OrElse Not IsFinite(weightC) Then
            Return False
        End If

        If weightA >= -GeometryTolerance AndAlso
           weightB >= -GeometryTolerance AndAlso
           weightC >= -GeometryTolerance Then
            weightA = ClampWeight(weightA)
            weightB = ClampWeight(weightB)
            weightC = ClampWeight(weightC)

            Dim weightSum As Double = weightA + weightB + weightC
            If Not IsFinite(weightSum) OrElse weightSum <= 0.0 Then Return False

            weightA /= weightSum
            weightB /= weightSum
            weightC /= weightSum

            Dim interpolatedElevation As Double = a.Z * weightA + b.Z * weightB + c.Z * weightC
            If Not IsFinite(interpolatedElevation) Then Return False

            elevation = interpolatedElevation
            squaredDistance = 0.0
            Return True
        End If

        If Not clampToTriangle Then Return False

        Dim hasClosestPoint As Boolean = False
        Dim closestElevation As Double = 0.0
        Dim closestSquaredDistance As Double = Double.MaxValue
        Dim candidateElevation As Double = 0.0
        Dim candidateSquaredDistance As Double = 0.0

        If TryGetElevationOnSegment(point, a, b, candidateElevation, candidateSquaredDistance) Then
            closestElevation = candidateElevation
            closestSquaredDistance = candidateSquaredDistance
            hasClosestPoint = True
        End If

        If TryGetElevationOnSegment(point, b, c, candidateElevation, candidateSquaredDistance) AndAlso
           (Not hasClosestPoint OrElse candidateSquaredDistance < closestSquaredDistance) Then
            closestElevation = candidateElevation
            closestSquaredDistance = candidateSquaredDistance
            hasClosestPoint = True
        End If

        If TryGetElevationOnSegment(point, c, a, candidateElevation, candidateSquaredDistance) AndAlso
           (Not hasClosestPoint OrElse candidateSquaredDistance < closestSquaredDistance) Then
            closestElevation = candidateElevation
            closestSquaredDistance = candidateSquaredDistance
            hasClosestPoint = True
        End If

        If Not hasClosestPoint Then Return False

        elevation = closestElevation
        squaredDistance = closestSquaredDistance
        Return True
    End Function

    Friend Shared Function TryGetNearestSurfacePointElevation(ByVal point As Vector2D,
                                                               ByVal points As IEnumerable(Of Vector3D),
                                                               ByRef elevation As Double) As Boolean
        If points Is Nothing OrElse Not IsFinite(point.X) OrElse Not IsFinite(point.Y) Then Return False

        Dim found As Boolean = False
        Dim nearestElevation As Double = 0.0
        Dim nearestSquaredDistance As Double = Double.MaxValue

        For Each surfacePoint As Vector3D In points
            If IsFinite(surfacePoint.X) AndAlso IsFinite(surfacePoint.Y) AndAlso IsFinite(surfacePoint.Z) Then
                Dim deltaX As Double = point.X - surfacePoint.X
                Dim deltaY As Double = point.Y - surfacePoint.Y
                Dim pointSquaredDistance As Double = deltaX * deltaX + deltaY * deltaY

                If IsFinite(pointSquaredDistance) AndAlso
                   (Not found OrElse pointSquaredDistance < nearestSquaredDistance) Then
                    nearestElevation = surfacePoint.Z
                    nearestSquaredDistance = pointSquaredDistance
                    found = True
                End If
            End If
        Next

        If Not found Then Return False

        elevation = nearestElevation
        Return True
    End Function

    Private Shared Function TryGetElevationOnSegment(ByVal point As Vector2D,
                                                      ByVal startPoint As Vector3D,
                                                      ByVal endPoint As Vector3D,
                                                      ByRef elevation As Double,
                                                      ByRef squaredDistance As Double) As Boolean
        Dim segmentX As Double = endPoint.X - startPoint.X
        Dim segmentY As Double = endPoint.Y - startPoint.Y
        Dim segmentLengthSquared As Double = segmentX * segmentX + segmentY * segmentY
        If Not IsFinite(segmentLengthSquared) OrElse segmentLengthSquared <= 0.0 Then Return False

        Dim fraction As Double = ((point.X - startPoint.X) * segmentX +
                                  (point.Y - startPoint.Y) * segmentY) / segmentLengthSquared
        If Not IsFinite(fraction) Then Return False

        fraction = Math.Max(0.0, Math.Min(1.0, fraction))

        Dim closestX As Double = startPoint.X + segmentX * fraction
        Dim closestY As Double = startPoint.Y + segmentY * fraction
        Dim deltaX As Double = point.X - closestX
        Dim deltaY As Double = point.Y - closestY
        Dim candidateSquaredDistance As Double = deltaX * deltaX + deltaY * deltaY
        Dim candidateElevation As Double = startPoint.Z * (1.0 - fraction) + endPoint.Z * fraction

        If Not IsFinite(candidateSquaredDistance) OrElse Not IsFinite(candidateElevation) Then Return False

        elevation = candidateElevation
        squaredDistance = candidateSquaredDistance
        Return True
    End Function

    Private Shared Function ClampWeight(ByVal value As Double) As Double
        Return Math.Max(0.0, Math.Min(1.0, value))
    End Function

    Private Shared Function IsFinite(ByVal value As Double) As Boolean
        Return Not Double.IsNaN(value) AndAlso Not Double.IsInfinity(value)
    End Function
End Class
