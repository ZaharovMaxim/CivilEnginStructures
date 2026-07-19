Imports System.ComponentModel
Imports CivilEnginStructures.BridgeGeometry
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities
Public Class ProjectionPoint
    Public Enum projectView
        <Description("Вид сверху")> Top = 0
        <Description("Вид снизу")> Bottom = 1
        <Description("Вид слева")> Left = 2
        <Description("Вид справа")> Right = 3
        <Description("Вид спереди")> Front = 4
        <Description("Вид сзади")> Back = 5
    End Enum
    Public originPoint As Vector3D
    Public projectPoint As Vector3D
End Class
Public Class TransformCivilStructure
    Private _scaleX As Double
    Private _scaleY As Double
    Private _rotation As Double  ' Угол поворота в радианах
    Private _offsetX As Double
    Private _offsetY As Double
    Private _deltaZ As Double
    Public Sub New()
        _scaleX = 1
        _scaleY = 1
        _rotation = 0
        _offsetX = 0
        _offsetY = 0
        _deltaZ = 0
    End Sub
    Public Property ScaleX() As Double
        Get
            Return _scaleX
        End Get
        Set(value As Double)
            _scaleX = value
        End Set
    End Property

    Public Property ScaleY() As Double
        Get
            Return _scaleY
        End Get
        Set(value As Double)
            _scaleY = value
        End Set
    End Property

    Public Property Rotation() As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property

    Public Property OffsetX() As Double
        Get
            Return _offsetX
        End Get
        Set(value As Double)
            _offsetX = value
        End Set
    End Property

    Public Property OffsetY() As Double
        Get
            Return _offsetY
        End Get
        Set(value As Double)
            _offsetY = value
        End Set
    End Property

    Public Property DeltaZ() As Double
        Get
            Return _deltaZ
        End Get
        Set(value As Double)
            _deltaZ = value
        End Set
    End Property
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ' Функция для вычисления параметров трансформации по двум точкам
    Public Function CalculateTransformation(srcPoint1 As Vector3D, destPoint1 As Vector3D, srcPoint2 As Vector3D, destPoint2 As Vector2D, Optional maxElevation As Double = 0) As Boolean
        ' Векторы в исходной и целевой системах
        Dim srcVector As New Vector2D(srcPoint2.X - srcPoint1.X, srcPoint2.Y - srcPoint1.Y)
        Dim destVector As New Vector2D(destPoint2.X - destPoint1.X, destPoint2.Y - destPoint1.Y)
        ' Вычисляем масштаб
        Dim srcLength As Double = Math.Sqrt(srcVector.X * srcVector.X + srcVector.Y * srcVector.Y)
        Dim destLength As Double = Math.Sqrt(destVector.X * destVector.X + destVector.Y * destVector.Y)
        If srcLength > 0 Then
            ScaleX = destLength / srcLength
            ScaleY = destLength / srcLength  ' Изотропный масштаб
        Else
            ScaleX = 1
            ScaleY = 1
        End If
        ' Вычисляем угол поворота
        Dim srcAngle As Double = Math.Atan2(srcVector.Y, srcVector.X)
        Dim destAngle As Double = Math.Atan2(destVector.Y, destVector.X)
        Rotation = destAngle - srcAngle
        ' Вычисляем смещение
        ' Применяем обратную трансформацию к точке назначения
        Dim rotatedSrcPoint1 As Vector2D = RotatePoint(srcPoint1, Rotation)
        Dim scaledSrcPoint1 As New Vector2D(
            rotatedSrcPoint1.X * ScaleX,
            rotatedSrcPoint1.Y * ScaleY
        )
        OffsetX = destPoint1.X - scaledSrcPoint1.X
        OffsetY = destPoint1.Y - scaledSrcPoint1.Y
        If maxElevation = 0 Then
            DeltaZ = (srcPoint1.Z + srcPoint2.Z) / 2
        Else
            DeltaZ = maxElevation
        End If

        Return True
    End Function
    ' Вспомогательная функция для поворота точки
    Private Function RotatePoint(point As Vector2D, angleRadians As Double) As Vector2D
        Dim cosAngle As Double = Math.Cos(angleRadians)
        Dim sinAngle As Double = Math.Sin(angleRadians)

        Return New Vector2D(
            point.X * cosAngle - point.Y * sinAngle,
            point.X * sinAngle + point.Y * cosAngle
        )
    End Function
    '=================================================================================================
    ' Функция для трансформации точки
    Public Function transformPoint(point As Vector2D) As Vector2D
        ' 1. Поворот
        Dim rotated As Vector2D = RotatePoint(point, Rotation)
        ' 2. Масштабирование
        Dim scaled As New Vector2D(rotated.X * ScaleX, rotated.Y * ScaleY)
        ' 3. Смещение
        Dim result As New Vector2D(scaled.X + OffsetX, scaled.Y + OffsetY)
        Return result
    End Function

    '================================================================================================================================
    'функция трансформирует линию
    Public Function transformAxisLine(ByVal axisLine As DwgLine) As List(Of Vector3D)
        Dim result As List(Of Vector3D) = New List(Of Vector3D)
        If IsNothing(axisLine) = True Then
            Return result
        End If
        If axisLine.Length = 0 Then
            Return result
        End If
        'получаем смещение по z
        'If maxTransformElevation = -999999 Then
        '    maxTransformElevation = axisLine.StartPoint.Z
        'End If
        'записываем результат
        Dim transformedStartPoint As Vector2D = transformPoint(axisLine.StartPoint.Pos)
        result.Add(New Vector3D(transformedStartPoint.X, transformedStartPoint.Y, DeltaZ - axisLine.StartPoint.Z))
        Dim transformedEndPoint As Vector2D = transformPoint(axisLine.EndPoint.Pos)
        result.Add(New Vector3D(transformedEndPoint.X, transformedEndPoint.Y, DeltaZ - axisLine.EndPoint.Z))
        Return result
    End Function
    '================================================================================================================================
    'функция трансформирует точку
    Public Function transformUserPoint(ByVal pointTransform As Topomatic.Cad.Foundation.Vector3D) As Topomatic.Cad.Foundation.Vector3D
        Dim result As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        If IsNothing(pointTransform) = True Then
            Return result
        End If
        'записываем результат
        Dim transformedStartPoint As Vector2D = transformPoint(pointTransform.Pos)
        result.X = Math.Round(transformedStartPoint.X, 3)
        result.Y = Math.Round(transformedStartPoint.Y, 3)
        result.Z = Math.Round(DeltaZ - pointTransform.Z, 3)
        Return result
    End Function


End Class
