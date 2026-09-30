Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc

Public Class StructuresLines
    '================================================================================================
    'создание структурной линии на основе полилини
    Public Shared Function CreateStructureLineByPolyline2d(ByVal acSurface As Surface, ByVal polyline As DwgPolyline, Optional ByVal code As Integer = 0, Optional ByVal desk As String = "", Optional ByVal offsetElevation As Double = 0) As StructureLine
        CreateStructureLineByPolyline2d = Nothing
        Try
            Dim editor As Topomatic.Sfc.PointEditor = New Topomatic.Sfc.PointEditor(acSurface)
            Dim tempNewStructureLine As Topomatic.Sfc.StructureLine = New Topomatic.Sfc.StructureLine
            Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
            Dim elev As Double = 0
            If IsNothing(polyline) = False Then
                If polyline.Count > 1 Then
                    For i As Integer = 0 To polyline.Count - 1
                        rp.X = polyline.Item(i).Vertex.X
                        rp.Y = polyline.Item(i).Vertex.Y
                        rp.Z = polyline.Elevation
                        Dim point As Topomatic.Sfc.SurfacePoint = New Topomatic.Sfc.SurfacePoint(rp)
                        Dim point_index As Integer = editor.Add(point)
                        tempNewStructureLine.Add(point_index)
                    Next i
                    acSurface.StructureLines.Add(tempNewStructureLine)
                    tempNewStructureLine.LinearCode = code
                    tempNewStructureLine.Description = desk
                    If offsetElevation = 0 Then
                        tempNewStructureLine.IsLimitation = True
                    Else
                        tempNewStructureLine.IsLimitation = False
                    End If
                End If
            End If
            Return tempNewStructureLine
        Catch ex As Exception
        End Try
    End Function
    '================================================================================================
    'создание структурной линии на основе полилини
    Public Shared Function CreateStructureLineByPolyline3d(ByVal acSurface As Surface, ByVal polyline As DwgPolyline3D, Optional ByVal code As Integer = 0, Optional ByVal desk As String = "", Optional ByVal offsetElevation As Double = 0) As StructureLine
        CreateStructureLineByPolyline3d = Nothing
        Try
            Return CreateStructureLineByPolyline3dCore(acSurface, polyline, code, desk, offsetElevation)
        Catch ex As Exception
        End Try
    End Function

    Public Shared Function RequireStructureLineByPolyline3d(ByVal acSurface As Surface, ByVal polyline As DwgPolyline3D, Optional ByVal code As Integer = 0, Optional ByVal desk As String = "", Optional ByVal offsetElevation As Double = 0) As StructureLine
        If acSurface Is Nothing Then Throw New ArgumentNullException(NameOf(acSurface), "Не задана поверхность для структурной линии.")
        If polyline Is Nothing Then Throw New ArgumentNullException(NameOf(polyline), "Не задана полилиния структурной линии.")
        If polyline.Count < 2 Then Throw New ArgumentException("Для структурной линии нужны минимум две точки.", NameOf(polyline))
        Dim result As StructureLine = CreateStructureLineByPolyline3dCore(acSurface, polyline, code, desk, offsetElevation)
        If result Is Nothing Then Throw New InvalidOperationException("Структурная линия не создана.")
        Return result
    End Function

    Private Shared Function CreateStructureLineByPolyline3dCore(ByVal acSurface As Surface, ByVal polyline As DwgPolyline3D, ByVal code As Integer, ByVal desk As String, ByVal offsetElevation As Double) As StructureLine
        Dim editor As Topomatic.Sfc.PointEditor = New Topomatic.Sfc.PointEditor(acSurface)
        Dim tempNewStructureLine As Topomatic.Sfc.StructureLine = New Topomatic.Sfc.StructureLine
        Dim rp As Topomatic.Cad.Foundation.Vector3D = Nothing
        If IsNothing(polyline) = False Then
            If polyline.Count > 1 Then
                For i As Integer = 0 To polyline.Count - 1
                    rp.X = polyline.Item(i).X
                    rp.Y = polyline.Item(i).Y
                    rp.Z = polyline.Item(i).Z + offsetElevation
                    Dim point As Topomatic.Sfc.SurfacePoint = New Topomatic.Sfc.SurfacePoint(rp)
                    Dim point_index As Integer = editor.Add(point)
                    tempNewStructureLine.Add(point_index)
                Next i
                acSurface.StructureLines.Add(tempNewStructureLine)
                tempNewStructureLine.LinearCode = code
                tempNewStructureLine.Description = desk
                If offsetElevation = 0 Then
                    tempNewStructureLine.IsLimitation = True
                Else
                    tempNewStructureLine.IsLimitation = False
                End If
            End If
        End If
        Return tempNewStructureLine
    End Function
End Class
