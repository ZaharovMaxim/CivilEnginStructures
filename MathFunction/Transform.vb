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

        ' Построение матрицы системы A * x = B
        ' A = [x y z 1] для каждой точки, размером [n x 4]
        Dim A As Double(,) = New Double(n - 1, 3) {}
        For i As Integer = 0 To n - 1
            A(i, 0) = sourcePoints(i).X
            A(i, 1) = sourcePoints(i).Y
            A(i, 2) = sourcePoints(i).Z
            A(i, 3) = 1.0
        Next

        ' Вектора целевых значений для X, Y, Z
        Dim Bx As Double() = New Double(n - 1) {}
        Dim By As Double() = New Double(n - 1) {}
        Dim Bz As Double() = New Double(n - 1) {}

        For i As Integer = 0 To n - 1
            Bx(i) = targetPoints(i).X
            By(i) = targetPoints(i).Y
            Bz(i) = targetPoints(i).Z
        Next

        ' Решение систем методом наименьших квадратов
        ' x = (A^T * A)^-1 * A^T * B
        Dim AtA As Double(,) = MultiplyTranspose(A, A)
        Dim AtA_inv As Double(,)

        If Not InvertMatrix4x4(AtA, AtA_inv) Then
            Return False
        End If

        Dim AtBx As Double(,) = MultiplyTransposeByVector(A, Bx)
        Dim AtBy As Double(,) = MultiplyTransposeByVector(A, By)
        Dim AtBz As Double(,) = MultiplyTransposeByVector(A, Bz)

        Dim coeffX As Double() = MultiplyMatrixByVector(AtA_inv, AtBx)
        Dim coeffY As Double() = MultiplyMatrixByVector(AtA_inv, AtBy)
        Dim coeffZ As Double() = MultiplyMatrixByVector(AtA_inv, AtBz)

        ' Заполнение матрицы 4x4
        m_Matrix = New Double(3, 3) {}
        m_Matrix(0, 0) = coeffX(0) : m_Matrix(0, 1) = coeffX(1) : m_Matrix(0, 2) = coeffX(2) : m_Matrix(0, 3) = coeffX(3)
        m_Matrix(1, 0) = coeffY(0) : m_Matrix(1, 1) = coeffY(1) : m_Matrix(1, 2) = coeffY(2) : m_Matrix(1, 3) = coeffY(3)
        m_Matrix(2, 0) = coeffZ(0) : m_Matrix(2, 1) = coeffZ(1) : m_Matrix(2, 2) = coeffZ(2) : m_Matrix(2, 3) = coeffZ(3)
        m_Matrix(3, 0) = 0 : m_Matrix(3, 1) = 0 : m_Matrix(3, 2) = 0 : m_Matrix(3, 3) = 1

        Return True
    End Function

    ''' <summary>
    ''' Вычисляет матрицу преобразования по 4 точкам (точное решение)
    ''' </summary>
    Public Function ComputeFrom4Points(src As Point3D(), dst As Point3D()) As Boolean
        If src.Length < 4 OrElse dst.Length < 4 Then
            Return False
        End If

        ' Построение матрицы A размером 4x4
        Dim A As Double(,) = New Double(3, 3) {}
        For i As Integer = 0 To 3
            A(i, 0) = src(i).X
            A(i, 1) = src(i).Y
            A(i, 2) = src(i).Z
            A(i, 3) = 1.0
        Next

        Dim detA As Double = Determinant4x4(A)
        If Math.Abs(detA) < 0.0000000001 Then
            Return False
        End If

        ' Решение для X
        Dim Ax0 As Double(,) = ReplaceColumn(A, 0, New Double() {dst(0).X, dst(1).X, dst(2).X, dst(3).X})
        Dim Ax1 As Double(,) = ReplaceColumn(A, 1, New Double() {dst(0).X, dst(1).X, dst(2).X, dst(3).X})
        Dim Ax2 As Double(,) = ReplaceColumn(A, 2, New Double() {dst(0).X, dst(1).X, dst(2).X, dst(3).X})
        Dim Ax3 As Double(,) = ReplaceColumn(A, 3, New Double() {dst(0).X, dst(1).X, dst(2).X, dst(3).X})

        Dim a11 As Double = Determinant4x4(Ax0) / detA
        Dim a12 As Double = Determinant4x4(Ax1) / detA
        Dim a13 As Double = Determinant4x4(Ax2) / detA
        Dim a14 As Double = Determinant4x4(Ax3) / detA

        ' Решение для Y
        Dim Ay0 As Double(,) = ReplaceColumn(A, 0, New Double() {dst(0).Y, dst(1).Y, dst(2).Y, dst(3).Y})
        Dim Ay1 As Double(,) = ReplaceColumn(A, 1, New Double() {dst(0).Y, dst(1).Y, dst(2).Y, dst(3).Y})
        Dim Ay2 As Double(,) = ReplaceColumn(A, 2, New Double() {dst(0).Y, dst(1).Y, dst(2).Y, dst(3).Y})
        Dim Ay3 As Double(,) = ReplaceColumn(A, 3, New Double() {dst(0).Y, dst(1).Y, dst(2).Y, dst(3).Y})

        Dim a21 As Double = Determinant4x4(Ay0) / detA
        Dim a22 As Double = Determinant4x4(Ay1) / detA
        Dim a23 As Double = Determinant4x4(Ay2) / detA
        Dim a24 As Double = Determinant4x4(Ay3) / detA

        ' Решение для Z
        Dim Az0 As Double(,) = ReplaceColumn(A, 0, New Double() {dst(0).Z, dst(1).Z, dst(2).Z, dst(3).Z})
        Dim Az1 As Double(,) = ReplaceColumn(A, 1, New Double() {dst(0).Z, dst(1).Z, dst(2).Z, dst(3).Z})
        Dim Az2 As Double(,) = ReplaceColumn(A, 2, New Double() {dst(0).Z, dst(1).Z, dst(2).Z, dst(3).Z})
        Dim Az3 As Double(,) = ReplaceColumn(A, 3, New Double() {dst(0).Z, dst(1).Z, dst(2).Z, dst(3).Z})

        Dim a31 As Double = Determinant4x4(Az0) / detA
        Dim a32 As Double = Determinant4x4(Az1) / detA
        Dim a33 As Double = Determinant4x4(Az2) / detA
        Dim a34 As Double = Determinant4x4(Az3) / detA

        ' Заполнение матрицы
        m_Matrix = New Double(3, 3) {}
        m_Matrix(0, 0) = a11 : m_Matrix(0, 1) = a12 : m_Matrix(0, 2) = a13 : m_Matrix(0, 3) = a14
        m_Matrix(1, 0) = a21 : m_Matrix(1, 1) = a22 : m_Matrix(1, 2) = a23 : m_Matrix(1, 3) = a24
        m_Matrix(2, 0) = a31 : m_Matrix(2, 1) = a32 : m_Matrix(2, 2) = a33 : m_Matrix(2, 3) = a34
        m_Matrix(3, 0) = 0 : m_Matrix(3, 1) = 0 : m_Matrix(3, 2) = 0 : m_Matrix(3, 3) = 1

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
        Dim n As Integer = A.GetLength(0)
        Dim m As Integer = B.GetLength(1)
        Dim result As Double(,) = New Double(n - 1, m - 1) {}

        For i As Integer = 0 To n - 1
            For j As Integer = 0 To m - 1
                Dim sum As Double = 0
                For k As Integer = 0 To A.GetLength(1) - 1
                    sum += A(i, k) * B(k, j)
                Next
                result(i, j) = sum
            Next
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

        ' Вычисляем масштабирование Z (Z = k * X)
        Dim sumK As Double = 0
        Dim countK As Integer = 0

        ' Вычисляем смещения по X, Y, Z
        Dim sumOffsetX As Double = 0
        Dim sumOffsetY As Double = 0
        Dim sumOffsetZ As Double = 0
        Dim countOffset As Integer = 0

        For i As Integer = 0 To sourcePoints.Count - 1
            Dim src As AffineTransform3D.Point3D = sourcePoints(i)
            Dim tgt As AffineTransform3D.Point3D = targetPoints(i)

            ' Для Z: если X не ноль, вычисляем коэффициент
            If Math.Abs(src.X) > 0.000001 Then
                sumK += tgt.Z / src.X
                countK += 1
            End If

            ' Для смещений: если X близко к 0, то это точки на левом краю
            If Math.Abs(src.X) < 0.000001 Then
                sumOffsetX += tgt.X - src.X
                sumOffsetY += tgt.Y - src.Y
                sumOffsetZ += tgt.Z - src.Z
                countOffset += 1
            End If
        Next

        ' Вычисляем средний коэффициент масштабирования Z
        If countK > 0 Then
            m_ScaleZ = sumK / countK
        Else
            m_ScaleZ = 0
        End If

        ' Вычисляем средние смещения
        If countOffset > 0 Then
            m_OffsetX = sumOffsetX / countOffset
            m_OffsetY = sumOffsetY / countOffset
            m_OffsetZ = sumOffsetZ / countOffset
        Else
            m_OffsetX = 0
            m_OffsetY = 0
            m_OffsetZ = 0
        End If

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