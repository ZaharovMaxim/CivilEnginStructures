Imports System.Windows.Media.Media3D
Public Class AffineTransform3D
    ' Матрица 4x4 в формате [row, column]
    Private m_Matrix As Double(,)

    Public ReadOnly Property Matrix As Double(,)
        Get
            Return m_Matrix
        End Get
    End Property

    Public Sub New()
        ' Инициализация единичной матрицей
        m_Matrix = New Double(3, 3) {}
        For i As Integer = 0 To 3
            m_Matrix(i, i) = 1.0
        Next
    End Sub
    Public Class Point3D
        Public Property X As Double
        Public Property Y As Double
        Public Property Z As Double

        Public Sub New(x As Double, y As Double, z As Double)
            Me.X = x
            Me.Y = y
            Me.Z = z
        End Sub

        Public Overrides Function ToString() As String
            Return String.Format("({0:F3}, {1:F3}, {2:F3})", X, Y, Z)
        End Function
    End Class
    ''' <summary>
    ''' Вычисляет матрицу аффинного преобразования по набору точек (метод наименьших квадратов)
    ''' </summary>
    ''' <param name="sourcePoints">Исходные точки (минимум 4)</param>
    ''' <param name="targetPoints">Целевые точки (минимум 4)</param>
    ''' <returns>True, если вычисление успешно</returns>
    Public Function ComputeFromPoints(sourcePoints As List(Of Point3D), targetPoints As List(Of Point3D)) As Boolean
        If sourcePoints.Count <> targetPoints.Count OrElse sourcePoints.Count < 4 Then
            Throw New ArgumentException("Необходимо минимум 4 точки, количество исходных и целевых точек должно совпадать")
        End If

        Dim n As Integer = sourcePoints.Count
        Dim meanSourceX As Double = 0
        Dim meanSourceY As Double = 0
        Dim meanSourceZ As Double = 0
        Dim meanTargetX As Double = 0
        Dim meanTargetY As Double = 0
        Dim meanTargetZ As Double = 0

        For i As Integer = 0 To n - 1
            meanSourceX += sourcePoints(i).X
            meanSourceY += sourcePoints(i).Y
            meanSourceZ += sourcePoints(i).Z
            meanTargetX += targetPoints(i).X
            meanTargetY += targetPoints(i).Y
            meanTargetZ += targetPoints(i).Z
        Next

        meanSourceX /= n
        meanSourceY /= n
        meanSourceZ /= n
        meanTargetX /= n
        meanTargetY /= n
        meanTargetZ /= n

        Dim sourceScaleX As Double = 0
        Dim sourceScaleY As Double = 0
        Dim sourceScaleZ As Double = 0
        Dim targetScaleX As Double = 0
        Dim targetScaleY As Double = 0
        Dim targetScaleZ As Double = 0

        For i As Integer = 0 To n - 1
            sourceScaleX = Math.Max(sourceScaleX, Math.Abs(sourcePoints(i).X - meanSourceX))
            sourceScaleY = Math.Max(sourceScaleY, Math.Abs(sourcePoints(i).Y - meanSourceY))
            sourceScaleZ = Math.Max(sourceScaleZ, Math.Abs(sourcePoints(i).Z - meanSourceZ))
            targetScaleX = Math.Max(targetScaleX, Math.Abs(targetPoints(i).X - meanTargetX))
            targetScaleY = Math.Max(targetScaleY, Math.Abs(targetPoints(i).Y - meanTargetY))
            targetScaleZ = Math.Max(targetScaleZ, Math.Abs(targetPoints(i).Z - meanTargetZ))
        Next

        If sourceScaleX = 0 OrElse sourceScaleY = 0 OrElse sourceScaleZ = 0 Then
            Return False
        End If

        ' Построение матрицы системы A * x = B
        ' Центрирование отделяет перенос, а масштабирование сохраняет ранг
        ' системы при больших абсолютных и малых локальных координатах.
        Dim A As Double(,) = New Double(n - 1, 2) {}
        For i As Integer = 0 To n - 1
            A(i, 0) = (sourcePoints(i).X - meanSourceX) / sourceScaleX
            A(i, 1) = (sourcePoints(i).Y - meanSourceY) / sourceScaleY
            A(i, 2) = (sourcePoints(i).Z - meanSourceZ) / sourceScaleZ
        Next

        ' Вектора целевых значений для X, Y, Z
        Dim Bx As Double() = New Double(n - 1) {}
        Dim By As Double() = New Double(n - 1) {}
        Dim Bz As Double() = New Double(n - 1) {}

        For i As Integer = 0 To n - 1
            Bx(i) = If(targetScaleX = 0, 0, (targetPoints(i).X - meanTargetX) / targetScaleX)
            By(i) = If(targetScaleY = 0, 0, (targetPoints(i).Y - meanTargetY) / targetScaleY)
            Bz(i) = If(targetScaleZ = 0, 0, (targetPoints(i).Z - meanTargetZ) / targetScaleZ)
        Next

        ' Решение систем методом наименьших квадратов
        ' x = (A^T * A)^-1 * A^T * B
        Dim AtA As Double(,) = MultiplyTranspose(A, A)
        Dim detAtA As Double = Determinant3x3(AtA)
        If Math.Abs(detAtA) < 0.000000000001 Then
            Return False
        End If

        Dim AtBx As Double(,) = MultiplyTransposeByVector(A, Bx)
        Dim AtBy As Double(,) = MultiplyTransposeByVector(A, By)
        Dim AtBz As Double(,) = MultiplyTransposeByVector(A, Bz)

        Dim coeffX As Double() = Solve3x3(AtA, New Double() {AtBx(0, 0), AtBx(1, 0), AtBx(2, 0)}, detAtA)
        Dim coeffY As Double() = Solve3x3(AtA, New Double() {AtBy(0, 0), AtBy(1, 0), AtBy(2, 0)}, detAtA)
        Dim coeffZ As Double() = Solve3x3(AtA, New Double() {AtBz(0, 0), AtBz(1, 0), AtBz(2, 0)}, detAtA)

        coeffX(0) = coeffX(0) * targetScaleX / sourceScaleX
        coeffX(1) = coeffX(1) * targetScaleX / sourceScaleY
        coeffX(2) = coeffX(2) * targetScaleX / sourceScaleZ
        coeffY(0) = coeffY(0) * targetScaleY / sourceScaleX
        coeffY(1) = coeffY(1) * targetScaleY / sourceScaleY
        coeffY(2) = coeffY(2) * targetScaleY / sourceScaleZ
        coeffZ(0) = coeffZ(0) * targetScaleZ / sourceScaleX
        coeffZ(1) = coeffZ(1) * targetScaleZ / sourceScaleY
        coeffZ(2) = coeffZ(2) * targetScaleZ / sourceScaleZ

        Dim offsetX As Double = meanTargetX - coeffX(0) * meanSourceX - coeffX(1) * meanSourceY - coeffX(2) * meanSourceZ
        Dim offsetY As Double = meanTargetY - coeffY(0) * meanSourceX - coeffY(1) * meanSourceY - coeffY(2) * meanSourceZ
        Dim offsetZ As Double = meanTargetZ - coeffZ(0) * meanSourceX - coeffZ(1) * meanSourceY - coeffZ(2) * meanSourceZ

        ' Заполнение матрицы 4x4
        Dim newMatrix As Double(,) = New Double(3, 3) {}
        newMatrix(0, 0) = coeffX(0) : newMatrix(0, 1) = coeffX(1) : newMatrix(0, 2) = coeffX(2) : newMatrix(0, 3) = offsetX
        newMatrix(1, 0) = coeffY(0) : newMatrix(1, 1) = coeffY(1) : newMatrix(1, 2) = coeffY(2) : newMatrix(1, 3) = offsetY
        newMatrix(2, 0) = coeffZ(0) : newMatrix(2, 1) = coeffZ(1) : newMatrix(2, 2) = coeffZ(2) : newMatrix(2, 3) = offsetZ
        newMatrix(3, 0) = 0 : newMatrix(3, 1) = 0 : newMatrix(3, 2) = 0 : newMatrix(3, 3) = 1
        m_Matrix = newMatrix

        Return True
    End Function

    ''' <summary>
    ''' Вычисляет матрицу преобразования по 4 точкам (точное решение)
    ''' </summary>
    Public Function ComputeFrom4Points(src As Point3D(), dst As Point3D()) As Boolean
        If src.Length < 4 OrElse dst.Length < 4 Then
            Return False
        End If

        ' Строим систему по векторам относительно первой точки, чтобы
        ' определитель не зависел от абсолютных координат.
        Dim A As Double(,) = New Double(2, 2) {}
        Dim scale As Double = 0
        For i As Integer = 1 To 3
            A(i - 1, 0) = src(i).X - src(0).X
            A(i - 1, 1) = src(i).Y - src(0).Y
            A(i - 1, 2) = src(i).Z - src(0).Z
            scale = Math.Max(scale, Math.Abs(A(i - 1, 0)))
            scale = Math.Max(scale, Math.Abs(A(i - 1, 1)))
            scale = Math.Max(scale, Math.Abs(A(i - 1, 2)))
        Next

        If scale = 0 Then
            Return False
        End If

        For i As Integer = 0 To 2
            For j As Integer = 0 To 2
                A(i, j) /= scale
            Next
        Next

        Dim detA As Double = Determinant3x3(A)
        If Math.Abs(detA) < 0.000000000001 Then
            Return False
        End If

        Dim coeffX As Double() = Solve3x3(A, New Double() {dst(1).X - dst(0).X, dst(2).X - dst(0).X, dst(3).X - dst(0).X}, detA)
        Dim coeffY As Double() = Solve3x3(A, New Double() {dst(1).Y - dst(0).Y, dst(2).Y - dst(0).Y, dst(3).Y - dst(0).Y}, detA)
        Dim coeffZ As Double() = Solve3x3(A, New Double() {dst(1).Z - dst(0).Z, dst(2).Z - dst(0).Z, dst(3).Z - dst(0).Z}, detA)

        For i As Integer = 0 To 2
            coeffX(i) /= scale
            coeffY(i) /= scale
            coeffZ(i) /= scale
        Next

        Dim a14 As Double = dst(0).X - coeffX(0) * src(0).X - coeffX(1) * src(0).Y - coeffX(2) * src(0).Z
        Dim a24 As Double = dst(0).Y - coeffY(0) * src(0).X - coeffY(1) * src(0).Y - coeffY(2) * src(0).Z
        Dim a34 As Double = dst(0).Z - coeffZ(0) * src(0).X - coeffZ(1) * src(0).Y - coeffZ(2) * src(0).Z

        ' Заполнение матрицы
        Dim newMatrix As Double(,) = New Double(3, 3) {}
        newMatrix(0, 0) = coeffX(0) : newMatrix(0, 1) = coeffX(1) : newMatrix(0, 2) = coeffX(2) : newMatrix(0, 3) = a14
        newMatrix(1, 0) = coeffY(0) : newMatrix(1, 1) = coeffY(1) : newMatrix(1, 2) = coeffY(2) : newMatrix(1, 3) = a24
        newMatrix(2, 0) = coeffZ(0) : newMatrix(2, 1) = coeffZ(1) : newMatrix(2, 2) = coeffZ(2) : newMatrix(2, 3) = a34
        newMatrix(3, 0) = 0 : newMatrix(3, 1) = 0 : newMatrix(3, 2) = 0 : newMatrix(3, 3) = 1
        m_Matrix = newMatrix

        Return True
    End Function

    ''' <summary>
    ''' Применяет преобразование к точке
    ''' </summary>
    Public Function TransformPoint(point As Point3D) As Point3D
        Dim x As Double = m_Matrix(0, 0) * point.X + m_Matrix(0, 1) * point.Y + m_Matrix(0, 2) * point.Z + m_Matrix(0, 3)
        Dim y As Double = m_Matrix(1, 0) * point.X + m_Matrix(1, 1) * point.Y + m_Matrix(1, 2) * point.Z + m_Matrix(1, 3)
        Dim z As Double = m_Matrix(2, 0) * point.X + m_Matrix(2, 1) * point.Y + m_Matrix(2, 2) * point.Z + m_Matrix(2, 3)
        Return New Point3D(x, y, z)
    End Function

    ''' <summary>
    ''' Возвращает матрицу в виде строки
    ''' </summary>
    Public Overrides Function ToString() As String
        Return String.Format("[{0:F6}, {1:F6}, {2:F6}, {3:F6}]" & vbCrLf &
                                 "[{4:F6}, {5:F6}, {6:F6}, {7:F6}]" & vbCrLf &
                                 "[{8:F6}, {9:F6}, {10:F6}, {11:F6}]" & vbCrLf &
                                 "[{12:F6}, {13:F6}, {14:F6}, {15:F6}]",
                                 m_Matrix(0, 0), m_Matrix(0, 1), m_Matrix(0, 2), m_Matrix(0, 3),
                                 m_Matrix(1, 0), m_Matrix(1, 1), m_Matrix(1, 2), m_Matrix(1, 3),
                                 m_Matrix(2, 0), m_Matrix(2, 1), m_Matrix(2, 2), m_Matrix(2, 3),
                                 m_Matrix(3, 0), m_Matrix(3, 1), m_Matrix(3, 2), m_Matrix(3, 3))
    End Function

    ''' <summary>
    ''' Возвращает матрицу в формате для CAD/ГИС систем (однострочная)
    ''' </summary>
    Public Function ToCadMatrixString() As String
        Return String.Format("matrix({0:F6}, {1:F6}, {2:F6}, {3:F6}, {4:F6}, {5:F6}, {6:F6}, {7:F6}, {8:F6}, {9:F6}, {10:F6}, {11:F6})",
                                 m_Matrix(0, 0), m_Matrix(0, 1), m_Matrix(0, 2), m_Matrix(0, 3),
                                 m_Matrix(1, 0), m_Matrix(1, 1), m_Matrix(1, 2), m_Matrix(1, 3),
                                 m_Matrix(2, 0), m_Matrix(2, 1), m_Matrix(2, 2), m_Matrix(2, 3))
    End Function

    ' ==================== Вспомогательные методы линейной алгебры ====================

    Private Shared Function MultiplyTranspose(A As Double(,), B As Double(,)) As Double(,)
        Dim n As Integer = A.GetLength(1)
        Dim m As Integer = B.GetLength(1)
        Dim result As Double(,) = New Double(n - 1, m - 1) {}

        For i As Integer = 0 To n - 1
            For j As Integer = 0 To m - 1
                Dim sum As Double = 0
                For k As Integer = 0 To A.GetLength(0) - 1
                    sum += A(k, i) * B(k, j)
                Next
                result(i, j) = sum
            Next
        Next
        Return result
    End Function

    Private Shared Function Determinant3x3(m As Double(,)) As Double
        Return m(0, 0) * (m(1, 1) * m(2, 2) - m(1, 2) * m(2, 1)) -
               m(0, 1) * (m(1, 0) * m(2, 2) - m(1, 2) * m(2, 0)) +
               m(0, 2) * (m(1, 0) * m(2, 1) - m(1, 1) * m(2, 0))
    End Function

    Private Shared Function Solve3x3(m As Double(,), values As Double(), determinant As Double) As Double()
        Dim result As Double() = New Double(2) {}
        For column As Integer = 0 To 2
            Dim replaced As Double(,) = DirectCast(m.Clone(), Double(,))
            For row As Integer = 0 To 2
                replaced(row, column) = values(row)
            Next
            result(column) = Determinant3x3(replaced) / determinant
        Next
        Return result
    End Function

    Private Shared Function MultiplyTransposeByVector(A As Double(,), b As Double()) As Double(,)
        Dim n As Integer = A.GetLength(1)
        Dim result As Double(,) = New Double(n - 1, 0) {}

        For i As Integer = 0 To n - 1
            Dim sum As Double = 0
            For j As Integer = 0 To A.GetLength(0) - 1
                sum += A(j, i) * b(j)
            Next
            result(i, 0) = sum
        Next
        Return result
    End Function

    Private Shared Function MultiplyMatrixByVector(M As Double(,), v As Double(,)) As Double()
        Dim n As Integer = M.GetLength(0)
        Dim result As Double() = New Double(n - 1) {}

        For i As Integer = 0 To n - 1
            Dim sum As Double = 0
            For j As Integer = 0 To n - 1
                sum += M(i, j) * v(j, 0)
            Next
            result(i) = sum
        Next
        Return result
    End Function

    Private Shared Function Determinant4x4(m As Double(,)) As Double
        Return m(0, 0) * (m(1, 1) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                              m(1, 2) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) +
                              m(1, 3) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1))) -
                   m(0, 1) * (m(1, 0) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                              m(1, 2) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                              m(1, 3) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0))) +
                   m(0, 2) * (m(1, 0) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) -
                              m(1, 1) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                              m(1, 3) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0))) -
                   m(0, 3) * (m(1, 0) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1)) -
                              m(1, 1) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0)) +
                              m(1, 2) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0)))
    End Function

    Private Shared Function InvertMatrix4x4(m As Double(,), ByRef inv As Double(,)) As Boolean
        inv = New Double(3, 3) {}
        Dim det As Double = Determinant4x4(m)

        If Math.Abs(det) < 0.0000000001 Then
            Return False
        End If

        inv(0, 0) = (m(1, 1) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                         m(1, 2) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) +
                         m(1, 3) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1))) / det

        inv(0, 1) = -(m(0, 1) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                          m(0, 2) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) +
                          m(0, 3) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1))) / det

        inv(0, 2) = (m(0, 1) * (m(1, 2) * m(3, 3) - m(1, 3) * m(3, 2)) -
                         m(0, 2) * (m(1, 1) * m(3, 3) - m(1, 3) * m(3, 1)) +
                         m(0, 3) * (m(1, 1) * m(3, 2) - m(1, 2) * m(3, 1))) / det

        inv(0, 3) = -(m(0, 1) * (m(1, 2) * m(2, 3) - m(1, 3) * m(2, 2)) -
                          m(0, 2) * (m(1, 1) * m(2, 3) - m(1, 3) * m(2, 1)) +
                          m(0, 3) * (m(1, 1) * m(2, 2) - m(1, 2) * m(2, 1))) / det

        inv(1, 0) = -(m(1, 0) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                          m(1, 2) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                          m(1, 3) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0))) / det

        inv(1, 1) = (m(0, 0) * (m(2, 2) * m(3, 3) - m(2, 3) * m(3, 2)) -
                         m(0, 2) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                         m(0, 3) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0))) / det

        inv(1, 2) = -(m(0, 0) * (m(1, 2) * m(3, 3) - m(1, 3) * m(3, 2)) -
                          m(0, 2) * (m(1, 0) * m(3, 3) - m(1, 3) * m(3, 0)) +
                          m(0, 3) * (m(1, 0) * m(3, 2) - m(1, 2) * m(3, 0))) / det

        inv(1, 3) = (m(0, 0) * (m(1, 2) * m(2, 3) - m(1, 3) * m(2, 2)) -
                         m(0, 2) * (m(1, 0) * m(2, 3) - m(1, 3) * m(2, 0)) +
                         m(0, 3) * (m(1, 0) * m(2, 2) - m(1, 2) * m(2, 0))) / det

        inv(2, 0) = (m(1, 0) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) -
                         m(1, 1) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                         m(1, 3) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0))) / det

        inv(2, 1) = -(m(0, 0) * (m(2, 1) * m(3, 3) - m(2, 3) * m(3, 1)) -
                          m(0, 1) * (m(2, 0) * m(3, 3) - m(2, 3) * m(3, 0)) +
                          m(0, 3) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0))) / det

        inv(2, 2) = (m(0, 0) * (m(1, 1) * m(3, 3) - m(1, 3) * m(3, 1)) -
                         m(0, 1) * (m(1, 0) * m(3, 3) - m(1, 3) * m(3, 0)) +
                         m(0, 3) * (m(1, 0) * m(3, 1) - m(1, 1) * m(3, 0))) / det

        inv(2, 3) = -(m(0, 0) * (m(1, 1) * m(2, 3) - m(1, 3) * m(2, 1)) -
                          m(0, 1) * (m(1, 0) * m(2, 3) - m(1, 3) * m(2, 0)) +
                          m(0, 3) * (m(1, 0) * m(2, 1) - m(1, 1) * m(2, 0))) / det

        inv(3, 0) = -(m(1, 0) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1)) -
                          m(1, 1) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0)) +
                          m(1, 2) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0))) / det

        inv(3, 1) = (m(0, 0) * (m(2, 1) * m(3, 2) - m(2, 2) * m(3, 1)) -
                         m(0, 1) * (m(2, 0) * m(3, 2) - m(2, 2) * m(3, 0)) +
                         m(0, 2) * (m(2, 0) * m(3, 1) - m(2, 1) * m(3, 0))) / det

        inv(3, 2) = -(m(0, 0) * (m(1, 1) * m(3, 2) - m(1, 2) * m(3, 1)) -
                          m(0, 1) * (m(1, 0) * m(3, 2) - m(1, 2) * m(3, 0)) +
                          m(0, 2) * (m(1, 0) * m(3, 1) - m(1, 1) * m(3, 0))) / det

        inv(3, 3) = (m(0, 0) * (m(1, 1) * m(2, 2) - m(1, 2) * m(2, 1)) -
                         m(0, 1) * (m(1, 0) * m(2, 2) - m(1, 2) * m(2, 0)) +
                         m(0, 2) * (m(1, 0) * m(2, 1) - m(1, 1) * m(2, 0))) / det

        Return True
    End Function

    Private Shared Function ReplaceColumn(m As Double(,), col As Integer, values As Double()) As Double(,)
        Dim result As Double(,) = DirectCast(m.Clone(), Double(,))
        For i As Integer = 0 To 3
            result(i, col) = values(i)
        Next
        Return result
    End Function
End Class

Public Class CustomTransformer
    Private m_ScaleZ As Double
    Private m_OffsetX As Double
    Private m_OffsetY As Double
    Private m_OffsetZ As Double
    Private m_IsInitialized As Boolean

    Public ReadOnly Property ScaleZ As Double
        Get
            Return m_ScaleZ
        End Get
    End Property

    Public ReadOnly Property IsInitialized As Boolean
        Get
            Return m_IsInitialized
        End Get
    End Property

    Public Sub New()
        m_ScaleZ = 0
        m_OffsetX = 0
        m_OffsetY = 0
        m_OffsetZ = 0
        m_IsInitialized = False
    End Sub

    ''' <summary>
    ''' Вычисляет коэффициенты преобразования по набору точек
    ''' </summary>
    ''' <param name="sourcePoints">Исходные точки</param>
    ''' <param name="targetPoints">Целевые точки</param>
    ''' <returns>True, если вычисление успешно</returns>
    Public Function ComputeTransform(sourcePoints As List(Of AffineTransform3D.Point3D),
                                     targetPoints As List(Of AffineTransform3D.Point3D)) As Boolean
        If sourcePoints.Count <> targetPoints.Count OrElse sourcePoints.Count < 2 Then
            Return False
        End If

        ' Вычисляем смещения по X и Y, а также средние X и Z
        Dim sumOffsetX As Double = 0
        Dim sumOffsetY As Double = 0
        Dim meanX As Double = 0
        Dim meanZ As Double = 0

        For i As Integer = 0 To sourcePoints.Count - 1
            Dim src As AffineTransform3D.Point3D = sourcePoints(i)
            Dim tgt As AffineTransform3D.Point3D = targetPoints(i)

            sumOffsetX += tgt.X - src.X
            sumOffsetY += tgt.Y - src.Y
            meanX += src.X
            meanZ += tgt.Z
        Next

        meanX /= sourcePoints.Count
        meanZ /= sourcePoints.Count

        Dim sumXX As Double = 0
        Dim sumXZ As Double = 0
        For i As Integer = 0 To sourcePoints.Count - 1
            Dim deltaX As Double = sourcePoints(i).X - meanX
            sumXX += deltaX * deltaX
            sumXZ += deltaX * (targetPoints(i).Z - meanZ)
        Next

        If sumXX > 0 Then
            m_ScaleZ = sumXZ / sumXX
        Else
            m_ScaleZ = 0
        End If

        m_OffsetX = sumOffsetX / sourcePoints.Count
        m_OffsetY = sumOffsetY / sourcePoints.Count
        m_OffsetZ = meanZ - m_ScaleZ * meanX

        m_IsInitialized = True
        Return True
    End Function

    ''' <summary>
    ''' Преобразует точку используя вычисленные коэффициенты
    ''' </summary>
    Public Function TransformPoint(point As AffineTransform3D.Point3D) As AffineTransform3D.Point3D
        If Not m_IsInitialized Then
            Throw New InvalidOperationException("Преобразователь не инициализирован. Сначала вызовите ComputeTransform.")
        End If

        Dim newX As Double = point.X + m_OffsetX
        Dim newY As Double = point.Y + m_OffsetY
        Dim newZ As Double = point.X * m_ScaleZ + m_OffsetZ

        Return New AffineTransform3D.Point3D(newX, newY, newZ)
    End Function

    ''' <summary>
    ''' Преобразует список точек
    ''' </summary>
    Public Function TransformPoints(points As List(Of AffineTransform3D.Point3D)) As List(Of AffineTransform3D.Point3D)
        Dim result As New List(Of AffineTransform3D.Point3D)()
        For Each point In points
            result.Add(TransformPoint(point))
        Next
        Return result
    End Function

    ''' <summary>
    ''' Выводит параметры преобразования
    ''' </summary>
    Public Overrides Function ToString() As String
        If Not m_IsInitialized Then
            Return "Преобразователь не инициализирован"
        End If

        Return String.Format("Z = X * {0:F6} + {1:F6}" & vbCrLf &
                             "X = X + {2:F6}" & vbCrLf &
                             "Y = Y + {3:F6}",
                             m_ScaleZ, m_OffsetZ, m_OffsetX, m_OffsetY)
    End Function

    ''' <summary>
    ''' Получить ошибку преобразования для набора точек
    ''' </summary>
    Public Function GetError(sourcePoints As List(Of AffineTransform3D.Point3D),
                            targetPoints As List(Of AffineTransform3D.Point3D)) As Double
        If sourcePoints.Count <> targetPoints.Count OrElse Not m_IsInitialized Then
            Return -1
        End If

        Dim totalError As Double = 0
        For i As Integer = 0 To sourcePoints.Count - 1
            Dim transformed As AffineTransform3D.Point3D = TransformPoint(sourcePoints(i))
            Dim dx As Double = transformed.X - targetPoints(i).X
            Dim dy As Double = transformed.Y - targetPoints(i).Y
            Dim dz As Double = transformed.Z - targetPoints(i).Z
            totalError += Math.Sqrt(dx * dx + dy * dy + dz * dz)
        Next

        Return totalError / sourcePoints.Count
    End Function
End Class
