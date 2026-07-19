Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters

Public Class FuncConvertData
    Public Shared Function FuncConvertVertexPolylineToArrayDbl(ByVal polyline As DwgPolyline, ByRef arrayVertex As Double(,)) As Boolean
        FuncConvertVertexPolylineToArrayDbl = False
        Erase arrayVertex
        Dim countArrayVertex As Integer = 0
        If IsNothing(polyline) = False Then
            Dim vert2d As Vector2D = Nothing
            For i As Integer = 0 To polyline.Count - 1
                Dim bVert As BugleVector2D = polyline.Item(i)
                Dim newVert2d As Vector2D = New Vector2D(bVert.Vertex.X, bVert.Vertex.Y)
                If IsNothing(vert2d) = True Then
                    vert2d = newVert2d
                    ReDim Preserve arrayVertex(1, countArrayVertex)
                    arrayVertex(0, countArrayVertex) = newVert2d.X
                    arrayVertex(1, countArrayVertex) = newVert2d.Y
                    countArrayVertex += 1
                Else
                    If newVert2d = vert2d Then
                        Continue For
                    Else
                        vert2d = newVert2d
                        ReDim Preserve arrayVertex(1, countArrayVertex)
                        arrayVertex(0, countArrayVertex) = newVert2d.X
                        arrayVertex(1, countArrayVertex) = newVert2d.Y
                        countArrayVertex += 1
                    End If
                End If
            Next i
        End If
        If IsArray(arrayVertex) = True Then
            Return True
        End If
    End Function
End Class
