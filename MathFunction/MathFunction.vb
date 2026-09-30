Imports System.Numerics
Imports System.Runtime.InteropServices
Imports System.Windows.Media.Media3D
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports NetTopologySuite.Utilities
Imports Topomatic.Alg.Bridges
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.Pipes.Layers.UserEntities
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
' Класс для представления точки в трехмерном пространстве
Public Class Tolerance
    Public Shared ReadOnly Property GlobalEqualPoint As Double = 0.000001
    Public Shared ReadOnly Property GlobalEqualVector As Double = 0.000001
    Public Shared ReadOnly Property GlobalAngle As Double = 0.000001 ' радианы
End Class


Public Class UPoint3D
    Public Property X As Double
    Public Property Y As Double
    Public Property Z As Double
    Public Sub New(x As Double, y As Double, z As Double)
        Me.X = x
        Me.Y = y
        Me.Z = z
    End Sub
End Class
' Класс для представления плоскости
Public Class UPlane
    Public Property Normal As Point3D
    Public Property D As Double
    Public Sub New(a As Double, b As Double, c As Double, d As Double)
        Me.Normal = New Point3D(a, b, c)
        Me.D = d
    End Sub
End Class
Public Class UPlane3D
    ' Функция для проверки, лежат ли четыре точки в одной плоскости
    Public Shared Function ArePointsCoplanar(P1 As UPoint3D, P2 As UPoint3D, P3 As UPoint3D, P4 As UPoint3D) As Boolean
        ' Векторы, лежащие на плоскости
        Dim v1 As New UPoint3D(P2.X - P1.X, P2.Y - P1.Y, P2.Z - P1.Z)
        Dim v2 As New UPoint3D(P3.X - P1.X, P3.Y - P1.Y, P3.Z - P1.Z)
        Dim v3 As New UPoint3D(P4.X - P1.X, P4.Y - P1.Y, P4.Z - P1.Z)
        ' Смешанное произведение векторов (объем параллелепипеда)
        Dim volume As Double = v1.X * (v2.Y * v3.Z - v2.Z * v3.Y) -
                               v1.Y * (v2.X * v3.Z - v2.Z * v3.X) +
                               v1.Z * (v2.X * v3.Y - v2.Y * v3.X)
        ' Если объем равен нулю, точки лежат в одной плоскости
        Return Math.Abs(volume) < Double.Epsilon
    End Function
    ' Функция для построения уравнения плоскости по трем точкам
    Public Shared Function GetPlaneFromPoints(P1 As UPoint3D, P2 As UPoint3D, P3 As UPoint3D) As UPlane
        ' Векторы, лежащие на плоскости
        Dim v1 As New UPoint3D(P2.X - P1.X, P2.Y - P1.Y, P2.Z - P1.Z)
        Dim v2 As New UPoint3D(P3.X - P1.X, P3.Y - P1.Y, P3.Z - P1.Z)
        ' Вектор нормали к плоскости через векторное произведение
        Dim normal As New UPoint3D(v1.Y * v2.Z - v1.Z * v2.Y, v1.Z * v2.X - v1.X * v2.Z, v1.X * v2.Y - v1.Y * v2.X)
        Dim D As Double = -(normal.X * P1.X + normal.Y * P1.Y + normal.Z * P1.Z)
        Return New UPlane(normal.X, normal.Y, normal.Z, D)
    End Function
    'пересечение линии и плоскости
    Public Shared Function FindIntersection(A As UPoint3D, B As UPoint3D, plane As UPlane) As UPoint3D
        ' Вектор направления отрезка
        Dim direction As New Point3D(B.X - A.X, B.Y - A.Y, B.Z - A.Z)
        ' Параметрическое уравнение отрезка: P = A + t * direction
        ' Подставляем в уравнение плоскости: plane.Normal.X * (A.X + t * direction.X) + ... + plane.D = 0
        ' Решаем уравнение относительно t
        Dim numerator As Double = plane.Normal.X * A.X + plane.Normal.Y * A.Y + plane.Normal.Z * A.Z + plane.D
        Dim denominator As Double = plane.Normal.X * direction.X + plane.Normal.Y * direction.Y + plane.Normal.Z * direction.Z
        ' Если знаменатель равен нулю, отрезок параллелен плоскости
        If Math.Abs(denominator) < Double.Epsilon Then
            Return Nothing
        End If
        Dim t As Double = -numerator / denominator
        ' Если t находится в диапазоне [0, 1], то точка пересечения лежит на отрезке
        If t >= 0 AndAlso t <= 1 Then
            Return New UPoint3D(A.X + t * direction.X, A.Y + t * direction.Y, A.Z + t * direction.Z)
        Else
            Return Nothing
        End If
    End Function
End Class

Public Class MathFunction
    '===================================================================================================
    'функция переводит градусы в радианы (23.77 - > 
    Public Shared Function FuncConvertDegtoRad(ByVal deg As Double) As Double
        Dim r1 As Double = deg * Math.PI
        r1 = r1 / 180
        Return r1
    End Function
    '===================================================================================================
    'функция округляет чило вниз 3.7777777 - 7.77
    Public Shared Function FuncTrimDigitFloo(ByVal a As Double, ByVal n As Integer) As Double
        Return Math.Floor(a * 10 ^ n) / 10 ^ n
    End Function
    '===================================================================================================
    'функция округляет чило вверх 3.7777777 - 7.78
    Public Shared Function FuncTrimDigitCall(ByVal a As Double, ByVal n As Integer) As Double
        Return Math.Ceiling(a * 10 ^ n) / 10 ^ n
    End Function
    '===================================================================================================
    'слияние двух чисел
    Public Shared Function CombineNumbers(num1 As Integer, num2 As Integer) As Integer
        Return Convert.ToInt32(num1.ToString() & num2.ToString())
    End Function
    '===================================================================================================
    Public Shared Function reverseAngle(ByVal angle As Double) As Double
        Dim reverseNewAngle As Double = angle + Math.PI
        If reverseNewAngle > Math.PI * 2 Then
            reverseNewAngle -= Math.PI * 2
        ElseIf reverseNewAngle < 0 Then
            reverseNewAngle += Math.PI * 2
        End If
        Return reverseNewAngle
    End Function
    Public Shared Function normalAngle(ByVal angle As Double, Optional k As Integer = 1) As Double
        Dim normalNewAngle As Double = angle + k * Math.PI / 2
        If normalNewAngle > Math.PI * 2 Then
            normalNewAngle -= Math.PI * 2
        ElseIf normalNewAngle < 0 Then
            normalNewAngle += Math.PI * 2
        End If
        Return normalNewAngle
    End Function
    '==============================================================================================================================================================
    'функция вычисляет прямоугольные координаты на основе угла и расстояния
    Public Shared Function funcCalcCoordinatesByInsPointAndAngle(ByVal InsPoint As Vector2D, ByVal Angle As Double, ByVal Lenght As Double, Optional roundZn As Integer = 3) As Vector2D
        funcCalcCoordinatesByInsPointAndAngle = New Vector2D(-1, -1)
        Try
            funcCalcCoordinatesByInsPointAndAngle.X = InsPoint.X + Lenght * Math.Cos(Angle)
            funcCalcCoordinatesByInsPointAndAngle.Y = InsPoint.Y + Lenght * Math.Sin(Angle)
            funcCalcCoordinatesByInsPointAndAngle.X = Math.Round(funcCalcCoordinatesByInsPointAndAngle.X, roundZn)
            funcCalcCoordinatesByInsPointAndAngle.Y = Math.Round(funcCalcCoordinatesByInsPointAndAngle.Y, roundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '===============================================================================================================================================================
    'функфия вычисляет расстояние между двумя точками (2d)
    Public Shared Function funcCalcDistanceByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 3) As Double
        funcCalcDistanceByToPoints2d = 0
        Try
            funcCalcDistanceByToPoints2d = Math.Sqrt((InsPoint1.X - InsPoint2.X) ^ 2 + (InsPoint1.Y - InsPoint2.Y) ^ 2)
            funcCalcDistanceByToPoints2d = Math.Round(funcCalcDistanceByToPoints2d, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция вычисляет координаты середины отрезка
    Public Shared Function funcCalcMiddleCoordByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 3) As Vector2D
        funcCalcMiddleCoordByToPoints2d = New Vector2D(-1, -1)
        Try
            funcCalcMiddleCoordByToPoints2d = New Vector2D((InsPoint1.X + InsPoint2.X) / 2, (InsPoint1.Y + InsPoint2.Y) / 2)
            funcCalcMiddleCoordByToPoints2d.X = Math.Round(funcCalcMiddleCoordByToPoints2d.X, RoundZn)
            funcCalcMiddleCoordByToPoints2d.Y = Math.Round(funcCalcMiddleCoordByToPoints2d.Y, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция вычисляет координаты середины отрезка 3d
    Public Shared Function funcCalcMiddleCoordByToPoints3d(ByVal InsPoint1 As Topomatic.Cad.Foundation.Vector3D, ByVal InsPoint2 As Topomatic.Cad.Foundation.Vector3D, Optional ByVal RoundZn As Integer = 3) As Topomatic.Cad.Foundation.Vector3D
        funcCalcMiddleCoordByToPoints3d = New Topomatic.Cad.Foundation.Vector3D(-1, -1, -1)
        Try
            Dim x As Double = Math.Round((InsPoint1.X + InsPoint2.X) / 2, RoundZn)
            Dim y As Double = Math.Round((InsPoint1.Y + InsPoint2.Y) / 2, RoundZn)
            Dim z As Double = Math.Round((InsPoint1.Z + InsPoint2.Z) / 2, RoundZn)
            Return New Topomatic.Cad.Foundation.Vector3D(x, y, z)
        Catch ex As System.Exception
        End Try
    End Function
    '=============================================================================================================================
    'функция возвращает точку на прямой
    Public Shared Function FuncCalcPointInLine(ByVal startPoint As Topomatic.Cad.Foundation.Vector3D, ByVal endPoint As Topomatic.Cad.Foundation.Vector3D, ByVal Lenght As Double, Optional round As Integer = 3) As Topomatic.Cad.Foundation.Vector3D
        FuncCalcPointInLine = Nothing
        Try
            Dim L As Double = (endPoint - startPoint).Length
            Dim k1 As Double = Lenght / L
            Dim X1 As Double = startPoint.X + k1 * (endPoint.X - startPoint.X)
            Dim Y1 As Double = startPoint.Y + k1 * (endPoint.Y - startPoint.Y)
            Dim z1 As Double = startPoint.Z + k1 * (endPoint.Z - startPoint.Z)
            Return New Topomatic.Cad.Foundation.Vector3D(X1, Y1, z1)
        Catch ex As System.Exception
        End Try
    End Function
    '=============================================================================================================================
    'функция возвращает точку на прямой 2d
    Public Shared Function FuncCalcPoint2DInLine(ByVal startPoint As Topomatic.Cad.Foundation.Vector2D, ByVal endPoint As Topomatic.Cad.Foundation.Vector2D, ByVal Lenght As Double, Optional round As Integer = 3) As Topomatic.Cad.Foundation.Vector2D
        FuncCalcPoint2DInLine = Nothing
        Try
            Dim L As Double = (endPoint - startPoint).Length
            Dim k1 As Double = Lenght / L
            Dim X1 As Double = startPoint.X + k1 * (endPoint.X - startPoint.X)
            Dim Y1 As Double = startPoint.Y + k1 * (endPoint.Y - startPoint.Y)
            Return New Topomatic.Cad.Foundation.Vector2D(X1, Y1)
        Catch ex As System.Exception
        End Try
    End Function
    '=============================================================================================================================
    'функция возвращает высоту точки на отрезке
    Public Shared Function FuncCalcElevationByLine(ByVal startPoint As Topomatic.Cad.Foundation.Vector3D, ByVal endPoint As Topomatic.Cad.Foundation.Vector3D, ByVal point As Topomatic.Cad.Foundation.Vector2D, Optional round As Integer = 3) As Double
        Dim fullLenghtLine As Double = (startPoint.Pos - endPoint.Pos).Length
        Dim l As Double = (startPoint.Pos - point).Length
        Dim fullElevationLine As Double = Math.Round(endPoint.Z - startPoint.Z, round)
        If fullElevationLine = 0 Then
            Return startPoint.Z
        Else
            Dim h As Double = Math.Round((l * fullElevationLine) / fullLenghtLine, round)
            Return startPoint.Z + h
        End If
    End Function

    Public Shared Function FuncCalcPositionLineByElevation(ByVal startPoint As Topomatic.Cad.Foundation.Vector3D, ByVal endPoint As Topomatic.Cad.Foundation.Vector3D, ByVal elevstion As Double) As Topomatic.Cad.Foundation.Vector2D
        Dim dz As Double = endPoint.Z - startPoint.Z
        If Math.Abs(dz) < 0.000000001 Then
            Throw New InvalidOperationException("Невозможно определить точку: отрезок имеет постоянную высоту.")
        End If
        Dim t As Double = (elevstion - startPoint.Z) / dz
        If t < 0.0 OrElse t > 1.0 Then
            Throw New ArgumentOutOfRangeException(NameOf(elevstion), "Высотная отметка находится вне отрезка.")
        End If
        Dim x As Double = startPoint.X + (endPoint.X - startPoint.X) * t
        Dim y As Double = startPoint.Y + (endPoint.Y - startPoint.Y) * t
        Return New Topomatic.Cad.Foundation.Vector2D(x, y)
    End Function

    '=================================================================================================================================================================
    'функция вычисляет математическое направление между 2 точками - между вектором [1, 0] и моим вектором
    Public Shared Function funcCalcAngleByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 6) As Double
        funcCalcAngleByToPoints2d = 0
        Try
            Dim dx As Double = InsPoint2.X - InsPoint1.X
            Dim dy As Double = InsPoint2.Y - InsPoint1.Y
            If dx = 0 And dy = 0 Then
                funcCalcAngleByToPoints2d = 0
            Else
                funcCalcAngleByToPoints2d = Math.Atan2(Math.Abs(dy), dx)
            End If
            funcCalcAngleByToPoints2d = Math.Round(funcCalcAngleByToPoints2d, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================
    ' Функция для опускания перпендикуляра из точки на луч
    Public Shared Function FuncPerpendicularToRay(rayStart As Vector2D, rayDirection As Vector2D, point As Vector2D) As Vector2D
        ' Нормализуем направляющий вектор луча
        Dim length As Double = Math.Sqrt(rayDirection.X * rayDirection.X + rayDirection.Y * rayDirection.Y)
        Dim normalizedDirection As New Vector2D(rayDirection.X / length, rayDirection.Y / length)
        ' Вектор от начальной точки луча к заданной точке
        Dim vectorToPoint As New Vector2D(point.X - rayStart.X, point.Y - rayStart.Y)
        ' Скалярное произведение для нахождения проекции
        Dim projectionLength As Double = vectorToPoint.X * normalizedDirection.X + vectorToPoint.Y * normalizedDirection.Y
        ' Находим основание перпендикуляра
        Dim perpendicularBase As New Vector2D(rayStart.X + projectionLength * normalizedDirection.X, rayStart.Y + projectionLength * normalizedDirection.Y)
        Return perpendicularBase
    End Function
    '=================================================================================================================================================================
    'функция вычисляет дирекционное направление между 2 точками
    Public Shared Function funcCalcDirectionAngleByToPoints2d(ByVal InsPoint1 As Vector2D, ByVal InsPoint2 As Vector2D, Optional ByVal RoundZn As Integer = 6) As Double
        funcCalcDirectionAngleByToPoints2d = 0
        Try
            Dim deltaX As Double = InsPoint2.X - InsPoint1.X
            Dim deltaY As Double = InsPoint2.Y - InsPoint1.Y
            If deltaX = 0 And deltaY = 0 Then
                funcCalcDirectionAngleByToPoints2d = 0
            Else
                funcCalcDirectionAngleByToPoints2d = Math.Atan2(deltaX, deltaY) * 180 / Math.PI
                If funcCalcDirectionAngleByToPoints2d < 0 Then
                    funcCalcDirectionAngleByToPoints2d += 360
                End If
            End If
            funcCalcDirectionAngleByToPoints2d = Math.Round(funcCalcDirectionAngleByToPoints2d, RoundZn)
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция определяет с какой стороны находится точка (слева -1, справа 1, на линии - 0, ошибка -2)
    Public Shared Function funcLeftOrRightPointToLinearObject(ByVal acEnt As Object, ByVal InsPoint As Vector2D, Optional RoundZn As Integer = 2) As Integer
        funcLeftOrRightPointToLinearObject = -2
        Try
            Dim linear_object As ILinearObject = acEnt
            Dim polyLine As Polyline3D = New Polyline3D()
            linear_object.GetPolyline(polyLine)
            Dim station As Double = -1
            Dim offset As Double = -1
            polyLine.PosToStaOffset(InsPoint, station, offset)
            offset = Math.Round(offset, RoundZn)
            If offset = 0 Then
                Return 0
            ElseIf offset < 0 Then
                Return -1
            ElseIf offset > 0 Then
                Return 1
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '==================================================================================================================================================================
    'функция сортирует двумерный массив
    '=====================================================================================
    'Функция сортирует числовой двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortDblArray(ByRef GeoArray(,) As Double, ByVal ind As Integer) As Boolean
        Dim Temp1 As Double() = Nothing
        Dim Temp2 As Double() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        If IsArray(GeoArray) = False Then
            FuncSortDblArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortDblArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(0) Then
            FuncSortDblArray = False
            Exit Function
        End If
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st1 = GeoArray(ind, i)
                    st2 = Temp2(ind)
                    If st2 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k
                    End If
                Next j
            Next i
        Catch
            Return False
        End Try
        FuncSortDblArray = True
    End Function
    '=====================================================================================
    'Функция сортирует строковый двухмерный массив по выбранному элементу (для любой размерности)
    Public Shared Function FuncSortStrArray(ByRef GeoArray(,) As String, ByVal ind As Integer) As Boolean
        FuncSortStrArray = False
        If IsArray(GeoArray) = False Then
            FuncSortStrArray = False
            Exit Function
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            FuncSortStrArray = False
            Exit Function
        End If
        If ind > GeoArray.GetUpperBound(0) Then
            FuncSortStrArray = False
            Exit Function
        End If
        Dim Temp1 As String() = Nothing
        Dim Temp2 As String() = Nothing
        Dim st1 As Double = Nothing
        Dim st2 As Double = Nothing
        Dim st3 As String = Nothing
        Dim st4 As String = Nothing
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем из него первую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    If IsNumeric(GeoArray(ind, i)) And IsNumeric(Temp2(ind)) Then
                        st1 = Val(GeoArray(ind, i))
                        st2 = Val(Temp2(ind))
                        If st2 < st1 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    Else
                        st3 = GeoArray(ind, i)
                        st4 = Temp2(ind)
                        If st4 < st3 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    End If
                Next j
            Next i
        Catch
            FuncSortStrArray = False
        End Try
        FuncSortStrArray = True
    End Function
    '===================================================================================================================================
    'функция сортирует массив по 2 элементам
    Public Shared Function FuncSortDblArray2(ByRef GeoArray(,) As Double, ByVal ind1 As Integer, ByVal ind2 As Integer) As Boolean
        FuncSortDblArray2 = False
        If IsArray(GeoArray) = False Then
            Return False
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            Return False
        End If
        If GeoArray.GetUpperBound(0) < 2 Then
            Return False
        End If
        Dim Temp1 As Double() = Nothing
        Dim Temp2 As Double() = Nothing
        Dim st1 As Double = 0
        Dim st2 As Double = 0
        Dim st3 As Double = 0
        Dim st4 As Double = 0
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                st1 = Val(GeoArray(ind1, i)) 'номер участка 1 строки
                st2 = Val(GeoArray(ind2, i)) 'номер контура 1 строки
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем вторую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st3 = Val(GeoArray(ind1, j)) 'номер участка 2 строки
                    st4 = Val(GeoArray(ind2, j)) 'номер контура 2 строки
                    'если номер участка 2 строки меньше чем номер участка 1 строки, меняем строки местами
                    If st3 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k
                        st1 = st3
                        st2 = st4
                    ElseIf st3 = st1 Then
                        If st4 < st2 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                            st1 = st3
                            st2 = st4
                        End If
                    End If
                Next j
            Next i
        Catch
            Return False
        End Try
        Return True
    End Function
    '===================================================================================================================================
    'функция сортирует массив по 2 элементам
    Public Shared Function FuncSortStrArray2(ByRef GeoArray(,) As String, ByVal ind1 As Integer, ByVal ind2 As Integer) As Boolean
        FuncSortStrArray2 = False
        If IsArray(GeoArray) = False Then
            Return False
        End If
        If GeoArray.GetUpperBound(1) < 1 Then
            Return False
        End If
        If GeoArray.GetUpperBound(0) < 2 Then
            Return False
        End If
        Dim Temp1 As String() = Nothing
        Dim Temp2 As String() = Nothing
        Dim st1 As Integer = 0
        Dim st2 As Integer = 0
        Dim st3 As Integer = 0
        Dim st4 As Integer = 0
        Try
            For i As Integer = 0 To GeoArray.GetUpperBound(1) - 1 'берем первый элемет
                st1 = Val(GeoArray(ind1, i)) 'номер  1 столбца
                st2 = Val(GeoArray(ind2, i)) 'номер  2 столбца
                For j As Integer = i + 1 To GeoArray.GetUpperBound(1) 'берем второй элемент
                    'копируем вторую строку во временный массив
                    Erase Temp2
                    For k As Integer = 0 To GeoArray.GetUpperBound(0)
                        ReDim Preserve Temp2(k)
                        Temp2(k) = GeoArray(k, j)
                    Next k
                    st3 = Val(GeoArray(ind1, j)) 'номер участка 2 строки
                    st4 = Val(GeoArray(ind2, j)) 'номер контура 2 строки
                    'если номер участка 2 строки меньше чем номер участка 1 строки, меняем строки местами
                    If st3 < st1 Then
                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, j) = GeoArray(k, i)
                        Next k

                        For k As Integer = 0 To Temp2.GetUpperBound(0)
                            GeoArray(k, i) = Temp2(k)
                        Next k

                    ElseIf st3 = st1 Then
                        If st4 < st2 Then
                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, j) = GeoArray(k, i)
                            Next k

                            For k As Integer = 0 To Temp2.GetUpperBound(0)
                                GeoArray(k, i) = Temp2(k)
                            Next k
                        End If
                    End If
                Next j
            Next i
        Catch
            Return False
        End Try
        Return True
    End Function
    '========================================================================================================================================
    'функция вычисляет точку пересечения двух лучей
    Public Shared Function FuncIntersectionTwoRay(ByVal r1 As Vector2D, ByVal r2 As Vector2D, ByVal p1 As Vector2D, ByVal p2 As Vector2D, ByRef pCross As Vector2D) As Boolean
        Dim v As Double = r2.X - r1.X
        Dim w As Double = r2.Y - r1.Y
        Dim v2 As Double = p2.X - p1.X
        Dim w2 As Double = p2.Y - p1.Y

        If v = 0 AndAlso w = 0 AndAlso v2 = 0 AndAlso w2 = 0 Then
            Return False
        ElseIf v = 0 AndAlso w = 0 Then
            Return False
        ElseIf v2 = 0 AndAlso w2 = 0 Then
            Return False
        End If

        Dim crossDirections As Double = v * w2 - w * v2
        Dim epsilon As Double = 0.000001 * Math.Sqrt((v * v + w * w) * (v2 * v2 + w2 * w2))
        If Math.Abs(crossDirections) <= epsilon Then
            'info.Id = 21
            'info.Message = "Лучи параллельны"
            Return False
        End If

        Dim deltaX As Double = p1.X - r1.X
        Dim deltaY As Double = p1.Y - r1.Y
        Dim t As Double = (deltaX * w2 - deltaY * v2) / crossDirections
        Dim t2 As Double = (deltaX * w - deltaY * v) / crossDirections

        If t < 0 OrElse t2 < 0 Then
            'info.Id = 20
            'info.Message = "Пересечения нет"
            Return False
        End If

        pCross.X = r1.X + v * t
        pCross.Y = r1.Y + w * t
        'info.Id = 0
        'info.Message = "Пересечение есть"
        Return True
    End Function
    Public Shared Function FuncFindLineIntersection(p1 As Vector2D, p2 As Vector2D, p3 As Vector2D, p4 As Vector2D, Optional ByRef boolResult As Boolean = False, Optional ByVal roundZn As Integer = 3) As Vector2D
        ' Вычисляем коэффициенты для уравнений прямых
        Dim A1 As Double = p2.Y - p1.Y
        Dim B1 As Double = p1.X - p2.X
        Dim C1 As Double = A1 * p1.X + B1 * p1.Y

        Dim A2 As Double = p4.Y - p3.Y
        Dim B2 As Double = p3.X - p4.X
        Dim C2 As Double = A2 * p3.X + B2 * p3.Y

        ' Определитель системы уравнений
        Dim determinant As Double = A1 * B2 - A2 * B1

        ' Если определитель равен 0 - прямые параллельны или совпадают
        If Math.Abs(determinant) < 0.000001 Then
            boolResult = False
            Return Nothing
        End If

        ' Находим точку пересечения
        Dim x As Double = Math.Round((B2 * C1 - B1 * C2) / determinant, roundZn)
        Dim y As Double = Math.Round((A1 * C2 - A2 * C1) / determinant, roundZn)
        boolResult = True
        Return New Vector2D(x, y)
    End Function
    '===============================================================================================
    'функция вычисляет точку пересечения двух отрезков (только явное пересечение)
    Public Shared Function FuncIntersectionTwoSegments(ByVal Pt1 As Vector2D, ByVal Pt2 As Vector2D, ByVal Pt3 As Vector2D, ByVal pt4 As Vector2D, ByRef InPoint As Vector2D) As Boolean
        FuncIntersectionTwoSegments = False
        Dim x1 As Double = Pt1.X
        Dim y1 As Double = Pt1.Y

        Dim x2 As Double = Pt2.X
        Dim y2 As Double = Pt2.Y

        Dim x3 As Double = Pt3.X
        Dim y3 As Double = Pt3.Y

        Dim x4 As Double = pt4.X
        Dim y4 As Double = pt4.Y
        Try
            Dim d As Double = (Pt1.X - Pt2.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (pt4.X - Pt3.X)
            Dim da As Double = (Pt1.X - Pt3.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt3.Y) * (pt4.X - Pt3.X)
            Dim db As Double = (Pt1.X - Pt2.X) * (Pt1.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (Pt1.X - Pt3.X)

            Dim ta As Double = da / d
            Dim tb As Double = db / d

            If (ta >= 0 And ta <= 1 And tb >= 0 And tb <= 1) Then
                Dim dx As Double = Pt1.X + ta * (Pt2.X - Pt1.X)
                Dim dy As Double = Pt1.Y + ta * (Pt2.Y - Pt1.Y)
                InPoint = New Vector2D(dx, dy)
                FuncIntersectionTwoSegments = True
            End If
        Catch ex As Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция находит точку пересечения перепендикуляра проведенного к линии
    Public Shared Function funcGetNearestPointOnLine(ByRef points As Double(,), ByVal pt As Topomatic.Cad.Foundation.Vector3D) As Topomatic.Cad.Foundation.Vector3D
        funcGetNearestPointOnLine = New Topomatic.Cad.Foundation.Vector3D(0, 0, 0)
        Try
            If points.GetUpperBound(1) > 0 Then
                For i As Integer = 0 To points.GetUpperBound(1) - 1
                    Dim x1 As Double = points(0, i)
                    Dim Y1 As Double = points(1, i)
                    Dim z1 As Double = points(2, i)

                    Dim x2 As Double = points(0, i + 1)
                    Dim Y2 As Double = points(1, i + 1)
                    Dim z2 As Double = points(2, i + 1)

                    If x1 = x2 And Y1 = Y2 Then
                        Continue For
                    End If
                    Dim k As Double = ((Y2 - Y1) * (pt.X - x1) - (x2 - x1) * (pt.Y - Y1)) / ((Y2 - Y1) ^ 2 + (x2 - x1) ^ 2)
                    Dim tempX As Double = pt.X - k * (Y2 - Y1)
                    Dim tempY As Double = pt.Y + k * (x2 - x1)
                    Dim tempZ As Double = (z1 + z2) / 2
                    Dim d1 As Double = Math.Sqrt((tempX - x1) ^ 2 + (tempY - Y1) ^ 2)
                    Dim d2 As Double = Math.Sqrt((tempX - x2) ^ 2 + (tempY - Y2) ^ 2)
                    Dim d0 As Double = Math.Sqrt((x1 - x2) ^ 2 + (Y1 - Y2) ^ 2)
                    If Math.Round(d1, 4) + Math.Round(d2, 4) = Math.Round(d0, 4) Then
                        Dim newPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(tempX, tempY, tempZ)
                        Return newPoint
                    End If
                Next i
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '=================================================================================================================================================================
    'функция находит точку пересечения перепендикуляра проведенного к линии
    Public Shared Function funcGetNearestPointOnLine(ByVal startLinePoint As Vector2D, ByVal endLinePoint As Vector2D, ByVal pointNearest As Topomatic.Cad.Foundation.Vector2D, ByRef Optional boolRez As Boolean = False) As Topomatic.Cad.Foundation.Vector2D
        Try
            Dim x1 As Double = startLinePoint.X
            Dim Y1 As Double = startLinePoint.Y

            Dim x2 As Double = endLinePoint.X
            Dim Y2 As Double = endLinePoint.Y

            If x1 = x2 And Y1 = Y2 Then
                boolRez = False
                Return Nothing
            End If
            Dim k As Double = ((Y2 - Y1) * (pointNearest.X - x1) - (x2 - x1) * (pointNearest.Y - Y1)) / ((Y2 - Y1) ^ 2 + (x2 - x1) ^ 2)
            Dim tempX As Double = pointNearest.X - k * (Y2 - Y1)
            Dim tempY As Double = pointNearest.Y + k * (x2 - x1)
            Dim d1 As Double = Math.Sqrt((tempX - x1) ^ 2 + (tempY - Y1) ^ 2)
            Dim d2 As Double = Math.Sqrt((tempX - x2) ^ 2 + (tempY - Y2) ^ 2)
            Dim d0 As Double = Math.Sqrt((x1 - x2) ^ 2 + (Y1 - Y2) ^ 2)

            Dim kontr1 As Double = Math.Round(d1 + d2, 3)
            Dim kontr2 As Double = Math.Round(d0, 3)
            If Math.Abs(kontr1 - kontr2) < 0.002 Then
                Dim newPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector2D(tempX, tempY)
                boolRez = True
                Return newPoint
            Else
                boolRez = False
            End If
        Catch ex As System.Exception
            boolRez = False
        End Try
    End Function
    Public Shared Function FuncFindPerpendicularPoint(point As Vector2D, linePoint1 As Vector2D, linePoint2 As Vector2D, ByRef distance As Double) As Vector2D
        ' Вектор прямой
        Dim lineVector As New Vector2D(linePoint2.X - linePoint1.X, linePoint2.Y - linePoint1.Y)

        ' Вектор от точки linePoint1 до заданной точки
        Dim pointVector As New Vector2D(point.X - linePoint1.X, point.Y - linePoint1.Y)

        ' Проекция вектора pointVector на вектор lineVector
        Dim dotProduct As Double = pointVector.X * lineVector.X + pointVector.Y * lineVector.Y
        Dim lineLengthSquared As Double = lineVector.X * lineVector.X + lineVector.Y * lineVector.Y

        ' Параметр t для положения основания перпендикуляра на прямой
        Dim t As Double = dotProduct / lineLengthSquared


        ' Вычисляем основание перпендикуляра
        Dim basePoint As New Vector2D(
            linePoint1.X + t * lineVector.X,
            linePoint1.Y + t * lineVector.Y
        )

        ' Вычисляем расстояние
        distance = Math.Sqrt((point.X - basePoint.X) ^ 2 + (point.Y - basePoint.Y) ^ 2)

        Return basePoint
    End Function
    '===============================================================================================
    'функция вычисляет точку пересечения двух отрезков
    Public Shared Function FuncIntersectPolyline(ByVal Pt1 As Vector2D, ByVal Pt2 As Vector2D, ByVal Pt3 As Vector2D, ByVal pt4 As Vector2D, ByRef InPoint As Vector2D) As Boolean
        FuncIntersectPolyline = False
        Dim x1 As Double = Pt1.X
        Dim y1 As Double = Pt1.Y

        Dim x2 As Double = Pt2.X
        Dim y2 As Double = Pt2.Y

        Dim x3 As Double = Pt3.X
        Dim y3 As Double = Pt3.Y

        Dim x4 As Double = pt4.X
        Dim y4 As Double = pt4.Y
        Try
            Dim d As Double = (Pt1.X - Pt2.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (pt4.X - Pt3.X)
            Dim da As Double = (Pt1.X - Pt3.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt3.Y) * (pt4.X - Pt3.X)
            Dim db As Double = (Pt1.X - Pt2.X) * (Pt1.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (Pt1.X - Pt3.X)

            Dim ta As Double = da / d
            Dim tb As Double = db / d

            If (ta >= 0 And ta <= 1 And tb >= 0 And tb <= 1) Then
                Dim dx As Double = Pt1.X + ta * (Pt2.X - Pt1.X)
                Dim dy As Double = Pt1.Y + ta * (Pt2.Y - Pt1.Y)
                InPoint = New Vector2D(dx, dy)
                FuncIntersectPolyline = True
            End If
        Catch ex As Exception
        End Try
    End Function
    '====================================================================================
    'функция проверяет лежит ли точка внутри полигона (Новый вариант)
    Public Shared Function FuncCrossPointPolyline(ByVal arrayPoint As Double(,), ByVal p As Vector2D) As Boolean
        Dim p1, p2 As Vector2D
        Dim inside As Boolean = False
        Dim poly As Double(,) = Nothing
        Dim countPoly As Integer = 0
        If IsNothing(arrayPoint) = True Then
            Return inside
        ElseIf arrayPoint.GetUpperBound(1) < 2 Then
            Return inside
        Else
            For i As Integer = 0 To arrayPoint.GetUpperBound(1)
                ReDim Preserve poly(1, i)
                poly(0, i) = arrayPoint(0, i)
                poly(1, i) = arrayPoint(1, i)
                countPoly += 1
            Next i
            If poly(0, 0) <> poly(0, poly.GetUpperBound(1)) And poly(1, 0) <> poly(1, poly.GetUpperBound(1)) Then
                ReDim Preserve poly(1, countPoly)
                poly(0, countPoly) = arrayPoint(0, 0)
                poly(1, countPoly) = arrayPoint(1, 0)
            End If
        End If
        '=========================================================================================
        Dim XOld As Double = poly(0, poly.GetUpperBound(1))
        Dim YOld As Double = poly(1, poly.GetUpperBound(1))
        Dim oldPoint As Vector2D = New Vector2D(XOld, YOld)
        For i As Integer = 0 To poly.GetUpperBound(1)
            Dim newPoint As Vector2D = New Vector2D(poly(0, i), poly(1, i))
            If newPoint.X > oldPoint.X Then
                p1 = oldPoint
                p2 = newPoint
            Else
                p1 = newPoint
                p2 = oldPoint
            End If
            If (newPoint.X < p.X) = (p.X <= oldPoint.X) AndAlso (p.Y - p1.Y) * (p2.X - p1.X) < (p2.Y - p1.Y) * (p.X - p1.X) Then
                inside = Not inside
            End If
            oldPoint = newPoint
        Next
        Return inside
    End Function
    '=================================================================================================
    'вычисление геометрического центра полигона
    Public Shared Function FuncFindCrossCenterPointPolyline(ByVal ArrayCoord As Double(,)) As Vector2D
        FuncFindCrossCenterPointPolyline = New Vector2D(-1, -1)
        Try
            If IsArray(ArrayCoord) = False Then
                Return New Vector2D(-1, -1)
            End If
            If ArrayCoord.GetUpperBound(1) < 2 Then
                Return New Vector2D(-1, -1)
            End If
            '2-ищем габаритный контейнер
            Dim minX As Double = 999999999999
            Dim minY As Double = 999999999999
            Dim maxX As Double = -999999999999
            Dim maxY As Double = -999999999999
            For i As Integer = 0 To ArrayCoord.GetUpperBound(1)
                Dim x As Double = ArrayCoord(0, i)
                Dim y As Double = ArrayCoord(1, i)
                If x > maxX Then
                    maxX = x
                End If
                If x < minX Then
                    minX = x
                End If
                If y > maxY Then
                    maxY = y
                End If
                If y < minY Then
                    minY = y
                End If
            Next
            '3-проводим горизонтальную линию по середине габаритного контейнера (ГОРИЗОНТАЛЬНУЮ ИЛИ ВЕРТИКАЛЬНУЮ)
            Dim pt1 As Vector2D
            Dim pt2 As Vector2D
            Dim dx As Double = Math.Abs(maxX - minX)
            Dim dy As Double = Math.Abs(maxY - minY)
            If dx >= dy Then
                pt1 = New Vector2D(minX, (minY + maxY) / 2)
                pt2 = New Vector2D(maxX, (minY + maxY) / 2)
            Else
                pt1 = New Vector2D((minX + maxX) / 2, minY)
                pt2 = New Vector2D((minX + maxX) / 2, maxY)
            End If
            '3-заносим в массив все пересечения фигуры с этой горизонтальной линией
            Dim CountArrayPt As Integer = 0
            Dim ArrayPt As Double(,) = Nothing
            For i As Integer = 0 To ArrayCoord.GetUpperBound(1) - 1
                Dim pt3 As Vector2D = New Vector2D(ArrayCoord(0, i), ArrayCoord(1, i))
                Dim pt4 As Vector2D = New Vector2D(ArrayCoord(0, i + 1), ArrayCoord(1, i + 1))
                If MathFunction.funcCalcDistanceByToPoints2d(pt3, pt4) > 0 Then
                    Dim IntersectPoint As Vector2D = New Vector2D(-1, -1)
                    Dim boolIntersectLine As Boolean = MathFunction.FuncIntersectPolyline(pt1, pt2, pt3, pt4, IntersectPoint)
                    If boolIntersectLine = True Then
                        ReDim Preserve ArrayPt(1, CountArrayPt)
                        ArrayPt(0, CountArrayPt) = IntersectPoint.X
                        ArrayPt(1, CountArrayPt) = IntersectPoint.Y
                        CountArrayPt += 1
                    End If
                End If
            Next
            '4-сортируем массив (в массиве должно быть по меньшей мере 2 точки)
            If ArrayPt.GetUpperBound(1) > 0 Then
                If ArrayPt.GetUpperBound(1) > 0 Then
                    If dx >= dy Then
                        Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(ArrayPt, 0)
                    Else
                        Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(ArrayPt, 1)
                    End If
                End If
            End If
            '5-точки должна быть максимально удалена от точек пересечения
            Dim DistInsert As Double = 0
            If ArrayPt.GetUpperBound(1) > 0 Then
                For j As Integer = 0 To ArrayPt.GetUpperBound(1) - 1
                    Dim xStart As Double = ArrayPt(0, j)
                    Dim yStart As Double = ArrayPt(1, j)
                    Dim xEnd As Double = ArrayPt(0, j + 1)
                    Dim yEnd As Double = ArrayPt(1, j + 1)
                    Dim ptStart As Vector2D = New Vector2D(xStart, yStart)
                    Dim ptEnd As Vector2D = New Vector2D(xEnd, yEnd)
                    Dim TempDist As Double = MathFunction.funcCalcDistanceByToPoints2d(ptStart, ptEnd)
                    'ищем среднюю точку
                    Dim FindCenterPoint As Vector2D = New Vector2D(-1, -1)
                    If dx >= dy Then
                        FindCenterPoint = New Vector2D((xStart + xEnd) / 2, yStart)
                    Else
                        FindCenterPoint = New Vector2D(xStart, (yStart + yEnd) / 2)
                    End If
                    Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(ArrayCoord, FindCenterPoint)
                    If PointCrosPoligon = True Then
                        If TempDist > DistInsert Then
                            FuncFindCrossCenterPointPolyline = New Vector2D(FindCenterPoint.X, FindCenterPoint.Y)
                            DistInsert = TempDist
                        End If
                    End If
                Next j
            End If
            '6-если неудача попробуем простой вариант
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                Dim FindCenterPoint As Vector2D = New Vector2D((maxX + minX) / 2, (maxY + minY) / 2)
                Dim PointCrosPoligon As Boolean = MathFunction.FuncCrossPointPolyline(ArrayCoord, FindCenterPoint)
                If PointCrosPoligon = True Then
                    FuncFindCrossCenterPointPolyline = New Vector2D(FindCenterPoint.X, FindCenterPoint.Y)
                End If
            End If
            If FuncFindCrossCenterPointPolyline.X = -1 And FuncFindCrossCenterPointPolyline.Y = -1 Then
                FuncFindCrossCenterPointPolyline = New Vector2D(ArrayCoord(0, 0), ArrayCoord(0, 1))
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона
    Public Shared Function FuncSQRTPolylineByArray(ByVal ArrayCoord As Double(,), Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolylineByArray = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If IsArray(ArrayCoord) = True Then
            If ArrayCoord.GetUpperBound(1) > 1 Then
                Dim Xish As Double = Math.Round(ArrayCoord(0, 0), RoundVertZn)
                Dim Yish As Double = Math.Round(ArrayCoord(1, 0), RoundVertZn)
                Dim count As Integer = ArrayCoord.GetUpperBound(1)
                For i As Integer = 0 To ArrayCoord.GetUpperBound(1) - 1
                    X1 = Math.Round(ArrayCoord(0, i), RoundVertZn)
                    Y1 = Math.Round(ArrayCoord(1, i), RoundVertZn)
                    X2 = Math.Round(ArrayCoord(0, i + 1), RoundVertZn)
                    Y2 = Math.Round(ArrayCoord(1, i + 1), RoundVertZn)
                    PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
                Next
                If Not (Xish = ArrayCoord(0, count) And Yish = ArrayCoord(1, count)) Then
                    PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
                End If
                PSqr = Math.Abs(PSqr)
                FuncSQRTPolylineByArray = Math.Round(PSqr / 2, RoundAreatZn)
            End If
        End If
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона
    Public Shared Function FuncSQRTPolylineByList(ByVal pt2dList As List(Of Vector2D), Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolylineByList = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If pt2dList.Count > 2 Then
            Dim Xish As Double = Math.Round(pt2dList.Item(0).X, RoundVertZn)
            Dim Yish As Double = Math.Round(pt2dList.Item(0).Y, RoundVertZn)
            For i As Integer = 0 To pt2dList.Count - 2
                X1 = Math.Round(pt2dList.Item(i).X, RoundVertZn)
                Y1 = Math.Round(pt2dList.Item(i).Y, RoundVertZn)
                X2 = Math.Round(pt2dList.Item(i + 1).X, RoundVertZn)
                Y2 = Math.Round(pt2dList.Item(i + 1).Y, RoundVertZn)
                PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
            Next
            If Not (Xish = pt2dList.Item(pt2dList.Count - 1).X And Yish = pt2dList.Item(pt2dList.Count - 1).Y) Then
                PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
            End If
            PSqr = Math.Abs(PSqr)
            FuncSQRTPolylineByList = Math.Round(PSqr / 2, RoundAreatZn)
        End If
    End Function
    '=============================================================================================
    'функция вычисляет площадь полигона
    Public Shared Function FuncSQRTPolylineByList2(ByVal pt2dList As List(Of BugleVector2D), Optional RoundVertZn As Integer = 2, Optional RoundAreatZn As Integer = 0) As Double
        FuncSQRTPolylineByList2 = 0
        Dim X1 As Double = 0
        Dim Y1 As Double = 0
        Dim X2 As Double = 0
        Dim Y2 As Double = 0
        Dim PSqr As Double = 0
        If pt2dList.Count > 2 Then
            Dim Xish As Double = Math.Round(pt2dList.Item(0).Vertex.X, RoundVertZn)
            Dim Yish As Double = Math.Round(pt2dList.Item(0).Vertex.Y, RoundVertZn)
            For i As Integer = 0 To pt2dList.Count - 2
                X1 = Math.Round(pt2dList.Item(0).Vertex.X, RoundVertZn)
                Y1 = Math.Round(pt2dList.Item(0).Vertex.Y, RoundVertZn)
                X2 = Math.Round(pt2dList.Item(i + 1).Vertex.X, RoundVertZn)
                Y2 = Math.Round(pt2dList.Item(i + 1).Vertex.Y, RoundVertZn)
                PSqr = PSqr + (Y1 + Y2) * (X1 - X2)
            Next
            If Not (Xish = pt2dList.Item(pt2dList.Count - 1).Vertex.X And Yish = pt2dList.Item(pt2dList.Count - 1).Vertex.Y) Then
                PSqr = PSqr + (Y2 + Yish) * (X2 - Xish)
            End If
            PSqr = Math.Abs(PSqr)
            FuncSQRTPolylineByList2 = Math.Round(PSqr / 2, RoundAreatZn)
        End If
    End Function
    'функция удлинняет отрезок
    Public Shared Function FuncExtendPos(ByRef startPoint As Topomatic.Cad.Foundation.Vector3D, ByRef endPoint As Topomatic.Cad.Foundation.Vector3D, ByVal startLenght As Double, ByVal endLenght As Double, Optional round As Integer = 3) As Boolean
        FuncExtendPos = False
        Try
            Dim L As Double = (startPoint - endPoint).Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = startPoint.X + k1 * (endPoint.X - startPoint.X)
            Dim Y1 As Double = startPoint.Y + k1 * (endPoint.Y - startPoint.Y)
            Dim z1 As Double = startPoint.Z + k1 * (endPoint.Z - startPoint.Z)
            Dim newStartPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = endPoint.X + k2 * (endPoint.X - startPoint.X)
            Dim Y2 As Double = endPoint.Y + k2 * (endPoint.Y - startPoint.Y)
            Dim z2 As Double = endPoint.Z + k2 * (endPoint.Z - startPoint.Z)
            Dim newEndPoint As Topomatic.Cad.Foundation.Vector3D = New Topomatic.Cad.Foundation.Vector3D(X2, Y2, z2)
            startPoint = newStartPoint
            endPoint = newEndPoint
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'функция возвращает точки начала и конца виртуального удлиннения балок
    Public Shared Function FuncVirtualExtendLine(ByRef line As DwgLine, ByVal startLenght As Double, ByVal endLenght As Double, ByRef startPt As Topomatic.Cad.Foundation.Vector3D, ByRef endPt As Topomatic.Cad.Foundation.Vector3D, Optional round As Integer = 3) As Boolean
        FuncVirtualExtendLine = False
        Try
            Dim L As Double = line.Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = line.StartPoint.X + k1 * (line.EndPoint.X - line.StartPoint.X)
            Dim Y1 As Double = line.StartPoint.Y + k1 * (line.EndPoint.Y - line.StartPoint.Y)
            Dim z1 As Double = line.StartPoint.Z + k1 * (line.EndPoint.Z - line.StartPoint.Z)
            startPt = New Topomatic.Cad.Foundation.Vector3D(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = line.EndPoint.X + k2 * (line.EndPoint.X - line.StartPoint.X)
            Dim Y2 As Double = line.EndPoint.Y + k2 * (line.EndPoint.Y - line.StartPoint.Y)
            Dim z2 As Double = line.EndPoint.Z + k2 * (line.EndPoint.Z - line.StartPoint.Z)
            endPt = New Topomatic.Cad.Foundation.Vector3D(X2, Y2, z2)
            Return True
        Catch ex As Exception
        End Try
    End Function
    'апроксимация дуги линейными се
    Public Shared Function FuncCreateCurveByLineWithPoint(ByVal startPointLine1 As Vector2D, ByVal endPointLine1 As Vector2D, ByVal startPointLine2 As Vector2D, ByVal endPointLine2 As Vector2D, ByVal curveRadius As Double, ByVal countSegments As Integer, ByRef arrayCoord As Double(,)) As Boolean
        Dim intersectPoint As Vector2D = New Vector2D(0, 0)
        Dim boolIntersect As Boolean = FuncIntersectPolyline(startPointLine1, endPointLine1, endPointLine2, startPointLine2, intersectPoint)
        'пересечение явного нет, делаем пересечение лучей
        If boolIntersect = False Then
            boolIntersect = FuncIntersectionTwoRay(startPointLine1, endPointLine1, endPointLine2, startPointLine2, intersectPoint)
        End If
        'пересечение найдено
        If boolIntersect = True Then
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'строим новые отрезки
            Dim tempLine1 As DwgLine = New DwgLine()
            tempLine1.StartPoint = startPointLine1
            tempLine1.EndPoint = endPointLine1
            Dim tempLine2 As DwgLine = New DwgLine()
            tempLine2.StartPoint = startPointLine2
            tempLine2.EndPoint = endPointLine2
            Dim direction1 As Double = tempLine1.Rotation
            Dim direction2 As Double = tempLine2.Rotation
            Dim phi As Double = direction2 - direction1
            Dim T As Double = curveRadius * Math.Tan(phi / 2)
            Dim pointLine1 As Vector2D = FuncCalcPoint2DInLine(intersectPoint, startPointLine1, T)
            Dim pointLine2 As Vector2D = FuncCalcPoint2DInLine(intersectPoint, endPointLine2, T)
            Dim middlePoint As Vector2D = funcCalcMiddleCoordByToPoints2d(pointLine1, pointLine2)
            Dim lineB As DwgLine = New DwgLine()
            lineB.StartPoint = intersectPoint
            lineB.EndPoint = middlePoint
            Dim sec As Double = 1 / Math.Cos(phi / 2)
            Dim B As Double = curveRadius * (sec - 1)
            Dim centerCircle As Vector2D = funcCalcCoordinatesByInsPointAndAngle(intersectPoint, lineB.Rotation, curveRadius + B)
            'делаем разбивку
            Dim LenghtLine As Double = (pointLine1 - pointLine2).Length
            Dim delta As Double = LenghtLine / countSegments
            For i As Integer = 0 To countSegments
                Dim tempPointLen As Vector2D = FuncCalcPoint2DInLine(pointLine1, pointLine2, i * delta)
                Dim angle As Double = (tempPointLen - centerCircle).Angle
                Dim rezPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(centerCircle, angle, curveRadius)
                ReDim Preserve arrayCoord(1, i)
                arrayCoord(0, i) = rezPoint.X
                arrayCoord(1, i) = rezPoint.Y
            Next i
        Else
            Return False
        End If
        If arrayCoord.GetUpperBound(1) = countSegments Then
            Return True
        Else
            Return False
        End If

    End Function
    '==============================================================================================
    'перемножение матриц
    Public Shared Function MatrixMultiplication(ByVal matrixA As Double(,), ByVal matrixB As Double(,)) As Double(,)
        If matrixA.GetLength(1) <> matrixB.GetLength(0) Then
            Throw New Exception("Умножение не возможно! Количество столбцов первой матрицы не равно количеству строк второй матрицы.")
        End If
        Dim matrixC = New Double(matrixA.GetUpperBound(0), matrixB.GetUpperBound(1)) {}
        For i = 0 To matrixA.GetUpperBound(0)
            For j = 0 To matrixB.GetUpperBound(1)
                matrixC(i, j) = 0
                For k = 0 To matrixA.GetUpperBound(1)
                    matrixC(i, j) += matrixA(i, k) * matrixB(k, j)
                Next
            Next
        Next
        Return matrixC
    End Function
    '================================================================================================
    'трансформирование координат
    Public Shared Function FuncTransformCoordinates(ByVal anglA As Double, ByVal anglW As Double, anglK As Double, ByVal dx As Double, ByVal dy As Double, ByVal dz As Double) As Double(,)
        Dim Ra As Double(,) = Nothing
        ReDim Ra(3, 3)
        Ra(0, 0) = 1 : Ra(1, 0) = 0 : Ra(2, 0) = 0 : Ra(3, 0) = 0
        Ra(0, 1) = 0 : Ra(1, 1) = Math.Cos(anglA) : Ra(2, 1) = -1 * Math.Sin(anglA) : Ra(3, 1) = 0
        Ra(0, 2) = 0 : Ra(1, 2) = Math.Sin(anglA) : Ra(2, 2) = Math.Cos(anglA) : Ra(3, 2) = 0
        Ra(0, 3) = 0 : Ra(1, 3) = 0 : Ra(2, 3) = 0 : Ra(3, 3) = 1

        Dim Rw As Double(,) = Nothing
        ReDim Rw(3, 3)
        Rw(0, 0) = Math.Cos(anglW) : Rw(1, 0) = 0 : Rw(2, 0) = Math.Sin(anglW) : Rw(3, 0) = 0
        Rw(0, 1) = 0 : Rw(1, 1) = 1 : Rw(2, 1) = 0 : Rw(3, 1) = 0
        Rw(0, 2) = -1 * Math.Sin(anglW) : Rw(1, 2) = 0 : Rw(2, 2) = Math.Cos(anglW) : Rw(3, 2) = 0
        Rw(0, 3) = 0 : Rw(1, 3) = 0 : Rw(2, 3) = 0 : Rw(3, 3) = 1

        Dim Rk As Double(,) = Nothing
        ReDim Rk(3, 3)
        Rk(0, 0) = Math.Cos(anglK) : Rk(1, 0) = -1 * Math.Sin(anglK) : Rk(2, 0) = 0 : Rk(3, 0) = 0
        Rk(0, 1) = Math.Sin(anglK) : Rk(1, 1) = Math.Cos(anglK) : Rk(2, 1) = 0 : Rk(3, 1) = 0
        Rk(0, 2) = 0 : Rk(1, 2) = 0 : Rk(2, 2) = 1 : Rk(3, 2) = 0
        Rk(0, 3) = 0 : Rk(1, 3) = 0 : Rk(2, 3) = 0 : Rk(3, 3) = 1

        Dim R1 As Double(,) = MatrixMultiplication(Ra, Rw)
        Dim R2 As Double(,) = MatrixMultiplication(R1, Rk)


        'Dim rez As Double(,) = MatrixMultiplication(R2, D)
        Return R2
    End Function
    'функция проверяет есть ли в одномерном массиве нужное значение
    Public Shared Function FuncFindValueToFArray(ByVal findStr As String, ByVal arrayData As String()) As Integer
        FuncFindValueToFArray = -1
        If IsArray(arrayData) = True Then
            For i As Integer = 0 To arrayData.GetUpperBound(0)
                Dim tempZn As String = arrayData(i)
                If IsNothing(tempZn) = False Then
                    If tempZn.Trim Like findStr.Trim Then
                        Return i
                    End If
                End If
            Next i
        End If
    End Function
    '======================================================================================
    'функция ищет значение в двумерном массиве
    Public Shared Function FuncFindValueToArray2d(ByVal ArrayData As String(,), ByVal nameField As String, Optional indexField As Integer = 0, Optional indexVal As Integer = 1) As String
        FuncFindValueToArray2d = ""
        If IsNothing(nameField) Then Return ""
        If IsArray(ArrayData) = False Then Return ""
        If ArrayData.GetUpperBound(0) < indexField Then Return ""
        If ArrayData.GetUpperBound(0) < indexVal Then Return ""
        Try
            If IsArray(ArrayData) = True Then
                For i As Integer = 0 To ArrayData.GetUpperBound(1)
                    Dim userField As String = ArrayData(indexField, i)
                    If IsNothing(userField) = False Then
                        If userField.Trim Like nameField.Trim Then
                            Dim valStr As String = ArrayData(indexVal, i)
                            If IsNothing(valStr) = True Then valStr = ""
                            Return valStr
                        End If
                    End If
                Next i
            End If
        Catch ex As Exception
        End Try
    End Function
    'функция проверяет есть ли в массиве нужное значение
    Public Shared Function FuncFindValueToArray(ByVal str As String, ByVal posArray As Integer, ByVal arrayData As String(,)) As Boolean
        FuncFindValueToArray = False
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(0) >= posArray Then
                For i As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim tempZn As String = arrayData(posArray, i)
                    If IsNothing(tempZn) = False Then
                        If tempZn.Trim Like str.Trim Then
                            Return True
                        End If
                    End If
                Next i
            End If
        End If
    End Function
    'функция проверяет есть ли в массиве нужное значение
    Public Shared Function FuncFindValueByFieldsToArray(ByVal nameField As String, ByVal value As String, ByVal arrayData As String(,)) As Boolean
        FuncFindValueByFieldsToArray = False
        If IsArray(arrayData) = True Then
            If arrayData.GetUpperBound(0) > 0 Then
                For i As Integer = 0 To arrayData.GetUpperBound(1)
                    Dim tempField As String = arrayData(0, i)
                    Dim tempVal As String = arrayData(1, i)
                    If IsNothing(tempField) = False Then
                        If tempField.Trim Like nameField.Trim Then
                            If IsNothing(tempVal) = False Then
                                If tempVal.Trim Like value.Trim Then
                                    Return True
                                End If
                            End If
                            Return False
                        End If
                    End If
                Next i
            End If
        End If
    End Function

    'функция вычисляет середину между смежными балками
    Public Shared Function calculateMiddlePointByLines(ByVal prevLine As DwgLine, ByVal line As DwgLine) As Vector3D
        calculateMiddlePointByLines = Nothing
        If IsNothing(line) = True Then Exit Function
        If IsNothing(prevLine) = True Then Exit Function
        Dim startPointPrevBeam As Vector3D = New Vector3D()
        Dim endPointPrevBeam As Vector3D = New Vector3D()
        Dim boolFindPoint As Boolean = MathFunction.FuncVirtualExtendLine(prevLine, 0, 0, startPointPrevBeam, endPointPrevBeam)

        Dim startPointBeam As Vector3D = New Vector3D()
        Dim endPointBeam As Vector3D = New Vector3D()
        boolFindPoint = MathFunction.FuncVirtualExtendLine(line, 0, 0, startPointBeam, endPointBeam)

        Dim middlePoint As Vector3D = MathFunction.funcCalcMiddleCoordByToPoints3d(endPointPrevBeam, startPointBeam)
        Return middlePoint
    End Function



    'функция добавляет элемент 
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '
    '=========================================================================================
    'функция удаляет совпадающие вершины (всех 2д и 3д)
    Public Shared Function FuncDeleteDubleVertexPline(ByRef ObjEnt As Topomatic.Dwg.DwgObject, Optional ByVal DopuskXY As Double = 0) As Boolean
        FuncDeleteDubleVertexPline = False
        If TypeOf ObjEnt Is Topomatic.Dwg.Entities.DwgPolyline Then
            Dim pLine As Topomatic.Dwg.Entities.DwgPolyline = ObjEnt
            Try
line1:
                'проверяем на совпадение последней и первой вершины
                Dim StartPt As Vector2D = New Vector2D(pLine.ElementAt(0).Vertex.X, pLine.ElementAt(0).Vertex.Y)
                Dim EndPt As Vector2D = New Vector2D(pLine.ElementAt(pLine.Count - 1).Vertex.X, pLine.ElementAt(pLine.Count - 1).Vertex.Y)
                Dim dist As Double = (StartPt - EndPt).Length
                If dist <= DopuskXY And pLine.Count > 1 Then
                    pLine.RemoveAt(pLine.Count - 1)
                    pLine.Closed = True
                    GoTo line1
                End If
                'делаем проверку для остальных вершин
                For i As Integer = 0 To pLine.Count - 2
                    Dim StartPt1 As Vector2D = New Vector2D(pLine.ElementAt(i).Vertex.X, pLine.ElementAt(i).Vertex.Y)
                    Dim EndPt1 As Vector2D = New Vector2D(pLine.ElementAt(i + 1).Vertex.X, pLine.ElementAt(i + 1).Vertex.Y)
                    Dim dist1 As Double = (StartPt1 - EndPt1).Length
                    If dist1 <= DopuskXY And pLine.Count > 1 Then
                        pLine.RemoveAt(i)
                        GoTo line1
                    End If
                Next i
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End If
    End Function

    '=========================================================================================
    'функция округляет вершины полилиний (всех 2д и 3д)
    Public Shared Function FuncRoundVertexPline(ByRef ObjEnt As Topomatic.Dwg.DwgObject, Optional ByVal RoundZnXY As Integer = 2) As Boolean
        FuncRoundVertexPline = False
        If TypeOf ObjEnt Is Topomatic.Dwg.Entities.DwgPolyline Then
            Dim pLine As Topomatic.Dwg.Entities.DwgPolyline = ObjEnt
            Try
                For i As Integer = 0 To pLine.Count - 1
                    Dim x As Double = pLine.ElementAt(i).Vertex.X
                    Dim y As Double = pLine.ElementAt(i).Vertex.Y
                    Dim newVertex As Vector2D = New Vector2D(Math.Round(x, RoundZnXY), Math.Round(y, RoundZnXY))
                    pLine.Item(i) = New BugleVector2D(newVertex, 0)
                Next i
                Return True
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

    '========================================================================================
    'функция делает реверс полилинии
    Public Shared Function FuncReversVertexPline(ByRef ObjEnt As Topomatic.Dwg.DwgObject) As Boolean
        FuncReversVertexPline = False
        If TypeOf ObjEnt Is Topomatic.Dwg.Entities.DwgPolyline Then
            Dim pLine As Topomatic.Dwg.Entities.DwgPolyline = ObjEnt
            Try
                Dim vertData As List(Of BugleVector2D) = pLine.ToList
                Dim count As Integer = 0
                For i As Integer = pLine.Count - 1 To 0 Step -1
                    Dim buildVert As BugleVector2D = vertData.Item(i)
                    If i = 0 Then
                        buildVert.Bugle = 0
                    Else
                        buildVert.Bugle = -1 * vertData.Item(i - 1).Bugle
                    End If
                    pLine.Item(count) = buildVert
                    count += 1
                Next
            Catch ex As Exception
            End Try
        End If
    End Function

    '========================================================================================
    'функция делает начальным северо-западный угол
    Public Shared Function FuncNorthWestAnglePline(ByRef ObjEnt As Topomatic.Dwg.DwgObject) As Boolean
        FuncNorthWestAnglePline = False
        If TypeOf ObjEnt Is Topomatic.Dwg.Entities.DwgPolyline Then
            Dim pLine As Topomatic.Dwg.Entities.DwgPolyline = ObjEnt
            Dim bound As BoundingBox2D = pLine.Bounds
            Dim NWAngleBound As Vector2D = New Vector2D(bound.Min.X, bound.Max.Y)
            Try
                Dim vertData As List(Of BugleVector2D) = pLine.ToList
                Dim startVertex As Integer = 0
                Dim startDist As Double = (bound.Max - bound.Min).Length
                For i As Integer = 0 To pLine.Count - 1
                    Dim tempVertex As Vector2D = New Vector2D(pLine.Item(i).Vertex.X, pLine.Item(i).Vertex.Y)
                    If (tempVertex - NWAngleBound).Length < startDist Then
                        startDist = (tempVertex - NWAngleBound).Length
                        startVertex = i
                    End If
                Next
                Dim newVertData As List(Of BugleVector2D) = New List(Of BugleVector2D)
                For i As Integer = startVertex To pLine.Count - 1
                    newVertData.Add(vertData.Item(i))
                Next
                For i As Integer = 0 To startVertex - 1
                    newVertData.Add(vertData.Item(i))
                Next
                For i As Integer = 0 To pLine.Count - 1
                    pLine.Item(i) = newVertData.Item(i)
                Next
            Catch ex As Exception
            End Try
        End If
    End Function

    '-===================================================================================
    'Функция возвращает координаты вставки атрибута и параметры точки выравнивания(массив-X,Y,ГорВыр 1-лево 2-право,ВертВыр 1-низ 2-верх)
    Public Shared Function FuncCalculatePositionAttribute(ByVal Vert1 As Vector2D, ByVal Vert2 As Vector2D, ByVal Vert3 As Vector2D, ByVal offsetX As Double, ByVal offsetY As Double) As Vector2D
        FuncCalculatePositionAttribute = New Vector2D(-1, -1)
        Try
            'определяем дирекционный угол на первую линию
            Dim Angle1 As Double = (Vert2 - Vert1).Angle
            'определяем дирекционый угол на вторую линию
            Dim Angle2 As Double = (Vert2 - Vert3).Angle
            'определяем внутренний угол
            Dim Angle As Double = Angle2 - Angle1
            If Angle < 0 Then Angle = Angle + Math.PI * 2 'если угол меньше 0 тогда добавим 360
            Dim Angle2D As Double = Angle / 2 'разделим угол пополам
            Dim Napravlenie As Double = Angle1 - Angle2D
            If Napravlenie > Math.PI * 2 Then Napravlenie = Napravlenie - Math.PI * 2 'если угол больше 360 градусов
            If Napravlenie < 0 Then Napravlenie = Napravlenie + Math.PI * 2
            'вычисляем координаты точки вставки атрибута
            Dim Dx As Double = offsetX * Math.Cos(Napravlenie) 'вычисляем приращения координат
            Dim Dy As Double = offsetY * Math.Sin(Napravlenie)
            Dim tempPoint2d As Vector2D = New Vector2D(Vert2.X + Dx, Vert2.Y + Dy)
            FuncCalculatePositionAttribute = tempPoint2d

        Catch ex As System.Exception

        End Try
    End Function

    Public Shared Function FuncAddRecordsToDictionary(ByRef userDectionary As Dictionary(Of String, String(,)), ByVal nameElement As String, ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal param4 As String, ByVal param5 As String) As Boolean
        FuncAddRecordsToDictionary = False
        If IsNothing(userDectionary) = True Then
            userDectionary = New Dictionary(Of String, String(,))
        End If
        Try
            If userDectionary.ContainsKey(nameElement) = True Then
                Dim arrayLine As String(,) = userDectionary.Item(nameElement)
                Dim boolFindObject As Boolean = False
                Dim ind As Integer = 0
                If IsArray(arrayLine) = True Then
                    For i As Integer = 0 To arrayLine.GetUpperBound(1)
                        Dim hgObject As String = arrayLine(4, i)
                        If hgObject Like param5 Then
                            arrayLine(1, i) = param2
                            boolFindObject = True
                            Exit For
                        End If
                    Next i
                End If
                If boolFindObject = False Then
                    ind = arrayLine.GetUpperBound(1) + 1
                    ReDim Preserve arrayLine(4, ind)
                    arrayLine(0, ind) = param1
                    arrayLine(1, ind) = param2
                    arrayLine(2, ind) = param3
                    arrayLine(3, ind) = param4
                    arrayLine(4, ind) = param5
                End If
                userDectionary.Item(nameElement) = arrayLine
            Else
                Dim arrayLine As String(,) = {}
                ReDim arrayLine(4, 0)
                arrayLine(0, 0) = param1
                arrayLine(1, 0) = param2
                arrayLine(2, 0) = param3
                arrayLine(3, 0) = param4
                arrayLine(4, 0) = param5
                userDectionary.Add(nameElement, arrayLine)
            End If
        Catch ex As System.Exception
        End Try
        Return True
    End Function
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'площадь участка
    '=============================================================================================================================
    Public Shared Function GetArea(pt1 As Vector2D, pt2 As Vector2D, pt3 As Vector2D) As Double
        Return (((pt2.X - pt1.X) * (pt3.Y - pt1.Y)) - ((pt3.X - pt1.X) * (pt2.Y - pt1.Y))) / 2
    End Function
    'Public Shared Function GetArea(arc As DwgArc) As Double
    '    Dim rad As Double = arc.Radius
    '    Dim ang As Double = If(arc.IsClockWise, arc.StartAngle - arc.EndAngle, arc.EndAngle - arc.StartAngle)
    '    Return rad * rad * (ang - Math.Sin(ang)) / 2
    'End Function
    Public Shared Function GetArea(pline As DwgPolyline, Optional RoundVertex As Integer = 4) As Double
        'Dim arc As New CircularArc2d()
        Dim area As Double = 0
        Dim last As Integer = pline.Count - 1
        Dim p0 As Vector2D = New Vector2D(Math.Round(pline.Item(0).Vertex.X, RoundVertex), Math.Round(pline.Item(0).Vertex.Y, RoundVertex))
        'If pline.GetBulgeAt(0) <> 0 Then
        '    area += GetArea(pline.GetArcSegment2dAt(0))
        'End If
        For i As Integer = 1 To last - 1
            Dim pt2Vert As Vector2D = New Vector2D(Math.Round(pline.Item(i).Vertex.X, RoundVertex), Math.Round(pline.Item(i).Vertex.Y, RoundVertex))
            Dim pt3Vert As Vector2D = New Vector2D(Math.Round(pline.Item(i + 1).Vertex.X, RoundVertex), Math.Round(pline.Item(i + 1).Vertex.Y, RoundVertex))
            area += GetArea(p0, pt2Vert, pt3Vert)
            'If pline.GetBulgeAt(i) <> 0 Then
            '    area += GetArea(pline.GetArcSegment2dAt(i))
            'End If
        Next
        'If (pline.GetBulgeAt(last) <> 0) AndAlso pline.Closed Then
        '    area += GetArea(pline.GetArcSegment2dAt(last))
        'End If
        Return area
    End Function
    Public Shared Function FuncAppPointArc(ByVal line1 As DwgLine, ByVal line2 As DwgLine, ByVal radius As Double, ByVal segments As Integer, ByRef arrayCoord As Double(,)) As Boolean
        Dim ptStartLine1 As Vector2D = line1.StartPoint.Pos
        Dim ptEndLine1 As Vector2D = line1.EndPoint.Pos

        Dim ptStartLine2 As Vector2D = line2.StartPoint.Pos
        Dim ptEndLine2 As Vector2D = line2.EndPoint.Pos

        Dim intersectPoint As Vector2D = New Vector2D(0, 0)
        Dim boolIntersect As Boolean = FuncIntersectPolyline(ptStartLine1, ptEndLine1, ptStartLine2, ptEndLine2, intersectPoint)
        'пересечение явного нет, делаем пересечение лучей
        If boolIntersect = False Then
            boolIntersect = FuncIntersectionTwoRay(ptStartLine1, ptEndLine1, ptStartLine2, ptEndLine2, intersectPoint)
        End If
        'пересечения нет, меняем направление одного отрезка
        If boolIntersect = False Then
            boolIntersect = FuncIntersectionTwoRay(ptEndLine1, ptStartLine1, ptStartLine2, ptEndLine2, intersectPoint)
        End If
        If boolIntersect = False Then
            boolIntersect = FuncIntersectionTwoRay(ptStartLine1, ptEndLine1, ptEndLine2, ptStartLine2, intersectPoint)
        End If
        'пересечение найдено
        If boolIntersect = True Then
            Dim lrn1 As Double = (intersectPoint - ptStartLine1).Length
            Dim lrn2 As Double = (intersectPoint - ptEndLine1).Length
            If lrn2 > lrn1 Then
                Dim tempPoint As Vector2D = ptStartLine1
                ptStartLine1 = ptEndLine1
                ptEndLine1 = tempPoint
            End If
            Dim lrn3 As Double = (intersectPoint - ptStartLine2).Length
            Dim lrn4 As Double = (intersectPoint - ptEndLine2).Length
            If lrn3 > lrn1 Then
                Dim tempPoint As Vector2D = ptStartLine1
                ptStartLine1 = ptEndLine1
                ptEndLine1 = tempPoint
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'строим новые отрезки
            Dim pointLine1 As Vector2D = FuncCalcPoint2DInLine(intersectPoint, ptStartLine1, radius)
            Dim pointLine2 As Vector2D = FuncCalcPoint2DInLine(intersectPoint, ptEndLine2, radius)

            Dim middlePoint As Vector2D = funcCalcMiddleCoordByToPoints2d(pointLine1, pointLine2)
            Dim andgle1 As Double = (middlePoint - intersectPoint).Angle
            Dim distance1 As Double = (intersectPoint - middlePoint).Length
            Dim centerCircle As Vector2D = funcCalcCoordinatesByInsPointAndAngle(middlePoint, andgle1, distance1)
            'делаем разбивку
            Dim LenghtLine As Double = (pointLine1 - pointLine2).Length
            Dim delta As Double = LenghtLine / segments
            For i As Integer = 0 To segments
                Dim tempPointLen As Vector2D = FuncCalcPoint2DInLine(pointLine1, pointLine2, i * delta)
                Dim angle As Double = (tempPointLen - centerCircle).Angle
                Dim rezPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(centerCircle, angle, radius)
                ReDim Preserve arrayCoord(1, i)
                arrayCoord(0, i) = rezPoint.X
                arrayCoord(1, i) = rezPoint.Y
            Next
        End If
        Return True
    End Function
    Public Shared Function FuncCalculatePerpendicularPointSimple(ByVal p1 As Topomatic.Cad.Foundation.Vector3D, ByVal p2 As Topomatic.Cad.Foundation.Vector3D, ByVal p3 As Topomatic.Cad.Foundation.Vector3D, ByVal length As Double) As Topomatic.Cad.Foundation.Vector3D
        ' 1. Вычисляем вектора от p2 к p1 и от p2 к p3
        Dim vectorP2P1 As Topomatic.Cad.Foundation.Vector3D = p1 - p2
        Dim vectorP2P3 As Topomatic.Cad.Foundation.Vector3D = p3 - p2
        ' 2. Вычисляем нормаль плоскости
        Dim normal As Topomatic.Cad.Foundation.Vector3D = Topomatic.Cad.Foundation.Vector3D.Cross(vectorP2P1, vectorP2P3)
        ' 3. Проверяем коллинеарность
        If normal.LengthSquared() < 0.00001 Then
            ' Коллинеарный случай
            Dim vectorP1P3 As Topomatic.Cad.Foundation.Vector3D = p3 - p1
            Dim perpendicularDir As Topomatic.Cad.Foundation.Vector3D
            ' Выбираем ось для векторного произведения
            If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                perpendicularDir = Topomatic.Cad.Foundation.Vector3D.Cross(vectorP1P3, Topomatic.Cad.Foundation.Vector3D.UnitY)
            Else
                perpendicularDir = Topomatic.Cad.Foundation.Vector3D.Cross(vectorP1P3, Topomatic.Cad.Foundation.Vector3D.UnitX)
            End If
            normal = Topomatic.Cad.Foundation.Vector3D.Normalize(perpendicularDir)
            ' Корректируем направление для гарантии Z < 0
            If normal.Z > 0 Then
                normal = -normal
            End If
            ' Дополнительная проверка по X
            If normal.X > 0 And normal.Z < 0 Then
                vectorP1P3 = p1 - p3
                If Math.Abs(vectorP1P3.X) > 0.00001 Or Math.Abs(vectorP1P3.Y) > 0.00001 Then
                    perpendicularDir = Topomatic.Cad.Foundation.Vector3D.Cross(vectorP1P3, Topomatic.Cad.Foundation.Vector3D.UnitY)
                Else
                    perpendicularDir = Topomatic.Cad.Foundation.Vector3D.Cross(vectorP1P3, Topomatic.Cad.Foundation.Vector3D.UnitX)
                End If
                normal = Topomatic.Cad.Foundation.Vector3D.Normalize(perpendicularDir)
                ' Снова проверяем и корректируем Z
                If normal.Z > 0 Then
                    normal = -normal
                End If
            End If
        Else
            ' Нормализуем вектор нормали
            normal = Topomatic.Cad.Foundation.Vector3D.Normalize(normal)
            ' Корректируем направление для гарантии Z < 0
            If normal.Z > 0 Then
                normal = -normal
            End If
        End If
        ' 4. Восстанавливаем перпендикуляр
        Dim resultPoint As Topomatic.Cad.Foundation.Vector3D = p2 + normal * CSng(length)
        ' 5. Финальная проверка и коррекция
        If resultPoint.Z >= p2.Z Then
            ' Если точка не ниже, принудительно опускаем
            Dim verticalDown As Topomatic.Cad.Foundation.Vector3D = -Topomatic.Cad.Foundation.Vector3D.UnitZ
            resultPoint = p2 + verticalDown * CSng(length)
        End If
        Return resultPoint
    End Function

    '===================================================================================================================================
    'возвращает расстояние от начала отрезка до бесконечной вертикальной плоскости
    Public Shared Function SignedDistanceFromSegmentStartToVerticalPlane(ByVal planePoint1 As Vector3D, ByVal planePoint2 As Vector3D, ByVal segmentStart As Vector3D, ByVal segmentEnd As Vector3D) As Double
        Dim planeDirectionX As Double = planePoint2.X - planePoint1.X
        Dim planeDirectionY As Double = planePoint2.Y - planePoint1.Y
        Dim planeDirectionLength As Double = Math.Sqrt(planeDirectionX * planeDirectionX + planeDirectionY * planeDirectionY)
        If planeDirectionLength <= Tolerance.GlobalEqualPoint Then
            Return Double.NaN
        End If

        Dim normalX As Double = -planeDirectionY / planeDirectionLength
        Dim normalY As Double = planeDirectionX / planeDirectionLength
        Dim startSignedDistance As Double = normalX * (segmentStart.X - planePoint1.X) + normalY * (segmentStart.Y - planePoint1.Y)
        If Math.Abs(startSignedDistance) <= Tolerance.GlobalEqualPoint Then
            Return 0.0
        End If

        Dim endSignedDistance As Double = normalX * (segmentEnd.X - planePoint1.X) + normalY * (segmentEnd.Y - planePoint1.Y)
        If Math.Abs(endSignedDistance) <= Tolerance.GlobalEqualPoint OrElse startSignedDistance * endSignedDistance < 0.0 Then
            Dim t As Double = startSignedDistance / (startSignedDistance - endSignedDistance)
            t = Math.Max(0.0, Math.Min(1.0, t))
            Return -((segmentEnd - segmentStart).Length * t)
        End If

        Return Math.Abs(startSignedDistance)
    End Function
    Private Shared Function IsPointInQuadrilateral(ByVal point As Vector3D, ByVal p1 As Vector3D, ByVal p2 As Vector3D, ByVal p3 As Vector3D, ByVal p4 As Vector3D) As Boolean
        ' Проверяем, что все точки лежат в одной плоскости
        ' (упрощенно - считаем что они в одной плоскости по условию задачи)
        ' Метод разбиения на два треугольника и проверка принадлежности
        Return IsPointInTriangle(point, p1, p2, p3) OrElse IsPointInTriangle(point, p1, p3, p4)
    End Function
    Private Shared Function IsPointInTriangle(ByVal point As Vector3D, ByVal t1 As Vector3D, ByVal t2 As Vector3D, ByVal t3 As Vector3D) As Boolean
        ' Вычисляем барицентрические координаты
        Dim v0 As Vector3D = t3 - t1
        Dim v1 As Vector3D = t2 - t1
        Dim v2 As Vector3D = point - t1
        ' Вычисляем скалярные произведения
        Dim dot00 As Double = DotProduct(v0, v0)
        Dim dot01 As Double = DotProduct(v0, v1)
        Dim dot02 As Double = DotProduct(v0, v2)
        Dim dot11 As Double = DotProduct(v1, v1)
        Dim dot12 As Double = DotProduct(v1, v2)
        ' Вычисляем барицентрические координаты
        Dim invDenom As Double = 1.0 / (dot00 * dot11 - dot01 * dot01)
        Dim u As Double = (dot11 * dot02 - dot01 * dot12) * invDenom
        Dim v As Double = (dot00 * dot12 - dot01 * dot02) * invDenom
        ' Проверяем, лежит ли точка внутри треугольника
        Return (u >= -Tolerance.GlobalEqualPoint) AndAlso
               (v >= -Tolerance.GlobalEqualPoint) AndAlso
               (u + v <= 1.0 + Tolerance.GlobalEqualPoint)
    End Function


    Public Shared Function CrossProduct(ByVal v1 As Vector3D, ByVal v2 As Vector3D) As Vector3D
        Dim result As New Vector3D()
        result.X = v1.Y * v2.Z - v1.Z * v2.Y
        result.Y = v1.Z * v2.X - v1.X * v2.Z
        result.Z = v1.X * v2.Y - v1.Y * v2.X
        Return result
    End Function

    Public Shared Function GetNormal(ByVal vector As Vector3D) As Vector3D
        Dim length As Double = Math.Sqrt(vector.X * vector.X +
                                      vector.Y * vector.Y +
                                      vector.Z * vector.Z)

        ' Проверка на нулевую длину
        If length < Tolerance.GlobalEqualPoint Then
            Return New Vector3D() ' Возвращаем нулевой вектор
        End If

        Return New Vector3D With {
        .X = vector.X / length,
        .Y = vector.Y / length,
        .Z = vector.Z / length
    }
    End Function

    Public Shared Function DotProduct(ByVal v1 As Vector3D, ByVal v2 As Vector3D) As Double
        Return v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z
    End Function
    Public Shared Function InscribeCircleBetweenSegments(startPtLine1 As Vector2D, endPtLine1 As Vector2D, startPtLine2 As Vector2D, endPtLine2 As Vector2D, radius As Double, countSegments As Integer) As List(Of Vector2D)
        Dim result As New List(Of Vector2D)()
        Dim l As Double = (endPtLine1 - startPtLine2).Length
        If l > 0 Then
            endPtLine1 = MathFunction.FuncFindLineIntersection(startPtLine1, endPtLine1, startPtLine2, endPtLine2)
        End If

        Dim line1 As DwgLine = New DwgLine
        line1.StartPoint = New Topomatic.Cad.Foundation.Vector3D(startPtLine1, 0)
        line1.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)

        Dim line2 As DwgLine = New DwgLine
        line2.StartPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)
        line2.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine2, 0)

        Dim dir1 As Double = line1.Rotation
        Dim dir2 As Double = line2.Rotation

        Dim deltaPhi As Double = Math.Abs(dir2 - dir1)
        Dim deltaPhiRad As Double = deltaPhi

        Dim T As Double = radius * Math.Tan(deltaPhiRad / 2)
        Dim B As Double = CalculateB(radius, deltaPhi)

        Dim pt1 As Vector2D = FuncCalcPoint2DInLine(endPtLine1, startPtLine1, T)
        Dim pt2 As Vector2D = FuncCalcPoint2DInLine(endPtLine1, endPtLine2, T)

        Dim middlePoint As Vector2D = funcCalcMiddleCoordByToPoints2d(pt1, pt2)
        Dim newLine As DwgLine = New DwgLine
        newLine.StartPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)
        newLine.EndPoint = New Topomatic.Cad.Foundation.Vector3D(middlePoint, 0)
        Dim centerPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(endPtLine1, newLine.Rotation, radius + B)
        Dim deltal As Double = (pt2 - pt1).Length / countSegments
        For i As Integer = 0 To countSegments
            Dim d As Double = deltal * i
            Dim ptTemp As Vector2D = FuncCalcPoint2DInLine(pt1, pt2, d)
            Dim tempLine As DwgLine = New DwgLine
            tempLine.StartPoint = New Topomatic.Cad.Foundation.Vector3D(centerPoint, 0)
            tempLine.EndPoint = New Topomatic.Cad.Foundation.Vector3D(ptTemp, 0)
            Dim calcPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(centerPoint, tempLine.Rotation, radius)
            result.Add(calcPoint)
        Next i
        Return result
    End Function
    Public Shared Function InscribeCircleBetweenSegmentsAutoRadius(startPtLine1 As Vector2D, endPtLine1 As Vector2D, startPtLine2 As Vector2D, endPtLine2 As Vector2D, countSegments As Integer) As List(Of Vector2D)

        Dim result As New List(Of Vector2D)()

        Dim l As Double = (endPtLine1 - startPtLine2).Length
        If l > 0 Then
            endPtLine1 = MathFunction.FuncFindLineIntersection(startPtLine1, endPtLine1, startPtLine2, endPtLine2)
        End If

        Dim line1 As DwgLine = New DwgLine
        line1.StartPoint = New Topomatic.Cad.Foundation.Vector3D(startPtLine1, 0)
        line1.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)

        Dim line2 As DwgLine = New DwgLine
        line2.StartPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)
        line2.EndPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine2, 0)

        Dim dir1 As Double = line1.Rotation
        Dim dir2 As Double = line2.Rotation

        Dim deltaPhi As Double = Math.Abs(dir2 - dir1)
        Dim deltaPhiRad As Double = deltaPhi

        If deltaPhiRad = 0 OrElse countSegments <= 0 Then
            Return result
        End If

        Dim lengthLine1 As Double = (endPtLine1 - startPtLine1).Length
        Dim lengthLine2 As Double = (endPtLine2 - endPtLine1).Length

        Dim T As Double = Math.Min(lengthLine1, lengthLine2)

        Dim radius As Double = T / Math.Tan(deltaPhiRad / 2)

        If radius <= 0 OrElse Double.IsNaN(radius) OrElse Double.IsInfinity(radius) Then
            Return result
        End If

        Dim B As Double = CalculateB(radius, deltaPhi)

        Dim pt1 As Vector2D = FuncCalcPoint2DInLine(endPtLine1, startPtLine1, T)
        Dim pt2 As Vector2D = FuncCalcPoint2DInLine(endPtLine1, endPtLine2, T)

        Dim middlePoint As Vector2D = funcCalcMiddleCoordByToPoints2d(pt1, pt2)

        Dim newLine As DwgLine = New DwgLine
        newLine.StartPoint = New Topomatic.Cad.Foundation.Vector3D(endPtLine1, 0)
        newLine.EndPoint = New Topomatic.Cad.Foundation.Vector3D(middlePoint, 0)

        Dim centerPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(
        endPtLine1,
        newLine.Rotation,
        radius + B
    )

        Dim deltal As Double = (pt2 - pt1).Length / countSegments

        For i As Integer = 0 To countSegments
            Dim d As Double = deltal * i

            Dim ptTemp As Vector2D = FuncCalcPoint2DInLine(pt1, pt2, d)

            Dim tempLine As DwgLine = New DwgLine
            tempLine.StartPoint = New Topomatic.Cad.Foundation.Vector3D(centerPoint, 0)
            tempLine.EndPoint = New Topomatic.Cad.Foundation.Vector3D(ptTemp, 0)

            Dim calcPoint As Vector2D = funcCalcCoordinatesByInsPointAndAngle(
            centerPoint,
            tempLine.Rotation,
            radius
        )

            result.Add(calcPoint)
        Next i

        Return result
    End Function



    Public Shared Function CalculateB(radius As Double, angleDegrees As Double) As Double
        ' Переводим угол в радианы
        Dim halfAngleRad As Double = angleDegrees / 2

        ' Вычисляем секанс (sec = 1/cos)
        Dim sec As Double = 1 / Math.Cos(halfAngleRad)

        ' Вычисляем результат
        Dim result As Double = radius * (sec - 1)

        Return result
    End Function

End Class


Public Class TriangleCenter
    ''' <summary>
    ''' Находит центр масс (центроид) треугольника
    ''' </summary>
    ''' <param name="a">Первая вершина треугольника</param>
    ''' <param name="b">Вторая вершина треугольника</param>
    ''' <param name="c">Третья вершина треугольника</param>
    ''' <returns>Центр масс треугольника</returns>
    Public Shared Function FindCentroid(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        Return New Vector2D((a.X + b.X + c.X) / 3.0, (a.Y + b.Y + c.Y) / 3.0)
    End Function

    ''' <summary>
    ''' Находит центр описанной окружности (circumcenter)
    ''' </summary>
    Public Shared Function FindCircumcenter(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        ' Проверка на коллинеарность точек
        If ArePointsCollinear(a, b, c) Then
            Throw New ArgumentException("Точки коллинеарны, невозможно найти центр описанной окружности")
        End If

        Dim d As Double = 2 * (a.X * (b.Y - c.Y) + b.X * (c.Y - a.Y) + c.X * (a.Y - b.Y))

        Dim ux As Double = ((a.X * a.X + a.Y * a.Y) * (b.Y - c.Y) +
                           (b.X * b.X + b.Y * b.Y) * (c.Y - a.Y) +
                           (c.X * c.X + c.Y * c.Y) * (a.Y - b.Y)) / d

        Dim uy As Double = ((a.X * a.X + a.Y * a.Y) * (c.X - b.X) +
                           (b.X * b.X + b.Y * b.Y) * (a.X - c.X) +
                           (c.X * c.X + c.Y * c.Y) * (b.X - a.X)) / d

        Return New Vector2D(ux, uy)
    End Function

    ''' <summary>
    ''' Находит центр вписанной окружности (incenter)
    ''' </summary>
    Public Shared Function FindIncenter(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        Dim ab As Double = Distance(b, c)  ' Противолежащая сторона для вершины A
        Dim bc As Double = Distance(a, c)  ' Противолежащая сторона для вершины B  
        Dim ca As Double = Distance(a, b)  ' Противолежащая сторона для вершины C

        Dim perimeter As Double = ab + bc + ca

        If perimeter < 0.000001 Then
            Throw New ArgumentException("Треугольник вырожденный")
        End If

        Dim x As Double = (ab * a.X + bc * b.X + ca * c.X) / perimeter
        Dim y As Double = (ab * a.Y + bc * b.Y + ca * c.Y) / perimeter

        Return New Vector2D(x, y)
    End Function

    ''' <summary>
    ''' Находит ортоцентр треугольника
    ''' </summary>
    Public Shared Function FindOrthocenter(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        If ArePointsCollinear(a, b, c) Then
            Throw New ArgumentException("Точки коллинеарны, невозможно найти ортоцентр")
        End If

        ' Находим координаты через барицентрические координаты
        Dim d As Double = (a.X - c.X) * (b.Y - c.Y) - (b.X - c.X) * (a.Y - c.Y)

        If Math.Abs(d) < 0.000001 Then
            Throw New ArgumentException("Точки коллинеарны")
        End If

        Dim x As Double = a.X + b.X + c.X - 2 * FindCircumcenter(a, b, c).X
        Dim y As Double = a.Y + b.Y + c.Y - 2 * FindCircumcenter(a, b, c).Y

        Return New Vector2D(x, y)
    End Function

    ''' <summary>
    ''' Находит центр тяжести (то же что и центроид)
    ''' </summary>
    Public Shared Function FindCenterOfMass(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        Return FindCentroid(a, b, c)
    End Function

    ''' <summary>
    ''' Находит центр по умолчанию (центроид)
    ''' </summary>
    Public Shared Function FindCenter(a As Vector2D, b As Vector2D, c As Vector2D) As Vector2D
        Return FindCentroid(a, b, c)
    End Function

    ''' <summary>
    ''' Проверяет, являются ли точки коллинеарными
    ''' </summary>
    Private Shared Function ArePointsCollinear(a As Vector2D, b As Vector2D, c As Vector2D) As Boolean
        ' Площадь треугольника через векторное произведение
        Dim area As Double = Math.Abs((b.X - a.X) * (c.Y - a.Y) - (c.X - a.X) * (b.Y - a.Y))
        Return area < 0.000001
    End Function

    ''' <summary>
    ''' Вычисляет расстояние между двумя точками
    ''' </summary>
    Private Shared Function Distance(p1 As Vector2D, p2 As Vector2D) As Double
        Return (p1 - p2).Length
    End Function

    ''' <summary>
    ''' Вычисляет площадь треугольника
    ''' </summary>
    Public Shared Function CalculateArea(a As Vector2D, b As Vector2D, c As Vector2D) As Double
        Return Math.Abs((b.X - a.X) * (c.Y - a.Y) - (c.X - a.X) * (b.Y - a.Y)) / 2.0
    End Function
End Class



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'триангуляция
'Public Class Triangle
'    ' Треугольник как три точки
'    Public Structure Triangle
'        Public A As Vector2D
'        Public B As Vector2D
'        Public C As Vector2D
'        Public Sub New(a As Vector2D, b As Vector2D, c As Vector2D)
'            Me.A = a : Me.B = b : Me.C = c
'        End Sub
'    End Structure
'    Private Const EPS As Double = 0.000000000001
'    ' Вспом. векторные операции
'    Private Function Cross(ax As Double, ay As Double, bx As Double, by As Double) As Double
'        Return ax * by - ay * bx
'    End Function
'    ' Векторное произведение (b - a) x (c - a)
'    Private Function Cross(a As Vector2D, b As Vector2D, c As Vector2D) As Double
'        Return Cross(b.X - a.X, b.Y - a.Y, c.X - a.X, c.Y - a.Y)
'    End Function
'    ' Площадь многоугольника (shoelace). >0 для CCW, <0 для CW
'    Private Function PolygonArea(pts As List(Of Vector2D)) As Double
'        Dim s As Double = 0
'        For i = 0 To pts.Count - 1
'            Dim j = (i + 1) Mod pts.Count
'            s += pts(i).X * pts(j).Y - pts(j).X * pts(i).Y
'        Next
'        Return 0.5 * s
'    End Function
'    ' Выпуклость вершины b между a и c с учётом ориентации
'    ' Для CW-порядка выпуклая вершина даёт Cross(a,b,c) < 0
'    Private Function IsConvex(a As Vector2D, b As Vector2D, c As Vector2D, cw As Boolean) As Boolean
'        Dim z = Cross(a, b, c)
'        If cw Then
'            Return z < -EPS
'        Else
'            Return z > EPS
'        End If
'    End Function
'    ' Проверка попадания точки в треугольник (включая границу)
'    Private Function PointInTriangle(p As Vector2D, a As Vector2D, b As Vector2D, c As Vector2D) As Boolean
'        Dim ab = Cross(a, b, p)
'        Dim bc = Cross(b, c, p)
'        Dim ca = Cross(c, a, p)
'        Dim hasNeg As Boolean = (ab < -EPS) OrElse (bc < -EPS) OrElse (ca < -EPS)
'        Dim hasPos As Boolean = (ab > EPS) OrElse (bc > EPS) OrElse (ca > EPS)
'        Return Not (hasNeg AndAlso hasPos)
'    End Function
'    ' Является ли вершина i "ушком"
'    Private Function IsEar(i As Integer, idx As List(Of Integer), pts As List(Of Vector2D), cw As Boolean) As Boolean
'        Dim iPrev = If(i = 0, idx.Count - 1, i - 1)
'        Dim iNext = If(i = idx.Count - 1, 0, i + 1)
'        Dim a = pts(idx(iPrev))
'        Dim b = pts(idx(i))
'        Dim c = pts(idx(iNext))
'        If Not IsConvex(a, b, c, cw) Then Return False
'        ' Отбрасываем вырожденные "тонкие" уши
'        If Math.Abs(Cross(a, b, c)) < EPS Then Return False
'        ' Внутри треугольника не должно быть других вершин
'        For j = 0 To idx.Count - 1
'            If j = iPrev OrElse j = i OrElse j = iNext Then Continue For
'            Dim p = pts(idx(j))
'            If PointInTriangle(p, a, b, c) Then Return False
'        Next
'        Return True
'    End Function
'    ' Главная функция: триангуляция CW-многоугольника
'    ' ВХОД: pts — список вершин замкнутой полилинии (повторять первую точку в конце НЕ нужно).
'    ' ВЫХОД: словарь {номер_треугольника -> Triangle(координаты вершин)}.
'    Public Function TriangulateClockwiseSimplePolygon(ByVal inputPts As List(Of Vector2D), ByRef dictionaryTriangle As Dictionary(Of Integer, Triangle)) As String
'        Dim result As New Dictionary(Of Integer, Triangle)()
'        If inputPts Is Nothing OrElse inputPts.Count < 3 Then
'            Dim str As String = "Нужно минимум 3 точки."
'            Return str
'        End If
'        ' Копия + удалим возможное дублирование последней точки с первой
'        Dim pts As New List(Of Vector2D)(inputPts)
'        If pts.Count >= 2 Then
'            Dim dx = pts(0).X - pts(pts.Count - 1).X
'            Dim dy = pts(0).Y - pts(pts.Count - 1).Y
'            If dx * dx + dy * dy < 1.0E-24 Then
'                pts.RemoveAt(pts.Count - 1)
'            End If
'        End If
'        If pts.Count < 3 Then
'            Dim str As String = "После нормализации осталось < 3 точек."
'            Return str
'        End If
'        ' Индексы текущего многоугольника
'        Dim idx As New List(Of Integer)(pts.Count)
'        For i = 0 To pts.Count - 1 : idx.Add(i) : Next
'        ' Гарантируем CW-порядок (если вдруг пришёл CCW — развернём индексы)
'        Dim isCW = PolygonArea(pts) < 0
'        If Not isCW Then
'            idx.Reverse()
'            isCW = True
'        End If
'        Dim t As Integer = 0
'        ' Вырезаем уши, пока не останется один треугольник
'        Dim guard As Integer = 0
'        While idx.Count > 3
'            Dim clipped As Boolean = False
'            For i = 0 To idx.Count - 1
'                If IsEar(i, idx, pts, isCW) Then
'                    Dim iPrev = If(i = 0, idx.Count - 1, i - 1)
'                    Dim iNext = If(i = idx.Count - 1, 0, i + 1)
'                    Dim a = pts(idx(iPrev))
'                    Dim b = pts(idx(i))
'                    Dim c = pts(idx(iNext))
'                    result.Add(t, New Triangle(a, b, c))
'                    t += 1
'                    idx.RemoveAt(i) ' отрезаем ухо
'                    clipped = True
'                    Exit For
'                End If
'            Next
'            guard += 1
'            If Not clipped Then
'                Dim str As String = "Триангуляция не удалась. Возможно, многоугольник вырожден или самопересекается."
'                Return str
'            End If
'            If guard > 100000 Then
'                Dim str As String = "Guard exceeded (защитный лимит)."
'                Return str
'            End If
'        End While
'        ' Финальный треугольник
'        Dim aF = pts(idx(0))
'        Dim bF = pts(idx(1))
'        Dim cF = pts(idx(2))
'        result.Add(t, New Triangle(aF, bF, cF))
'        dictionaryTriangle = result
'        Return "OK"
'    End Function
'End Class
