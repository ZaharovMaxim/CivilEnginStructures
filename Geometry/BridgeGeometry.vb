Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Crs.Rail
Imports Topomatic.Dwg.Entities
Public Class BridgeGeometry

    Public Class pointProjectionBridge
        Public Property code As String
        Public Property leftRight As String
        Public Property topBottom As String
        Public originalCoordinate As Vector3D
        Public projectionСoordinates As Vector3D
    End Class
    'функция удлинняет/укорачивает линию
    Public Shared Function extendLine(ByRef acLine As DwgLine, ByVal startLenght As Double, ByVal endLenght As Double, Optional round As Integer = 3) As Boolean
        extendLine = False
        If IsNothing(acLine) = True Then Return False
        Try
            Dim L As Double = acLine.Length
            Dim k1 As Double = -1 * (startLenght / L)
            Dim X1 As Double = acLine.StartPoint.X + k1 * (acLine.EndPoint.X - acLine.StartPoint.X)
            Dim Y1 As Double = acLine.StartPoint.Y + k1 * (acLine.EndPoint.Y - acLine.StartPoint.Y)
            Dim z1 As Double = acLine.StartPoint.Z + k1 * (acLine.EndPoint.Z - acLine.StartPoint.Z)
            Dim startPoint As Vector3D = New Vector3D(X1, Y1, z1)
            Dim k2 As Double = endLenght / L
            Dim X2 As Double = acLine.EndPoint.X + k2 * (acLine.EndPoint.X - acLine.StartPoint.X)
            Dim Y2 As Double = acLine.EndPoint.Y + k2 * (acLine.EndPoint.Y - acLine.StartPoint.Y)
            Dim z2 As Double = acLine.EndPoint.Z + k2 * (acLine.EndPoint.Z - acLine.StartPoint.Z)
            Dim endPoint As Vector3D = New Vector3D(X2, Y2, z2)
            acLine.StartPoint = startPoint
            acLine.EndPoint = endPoint
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    '==========================================================================================================
    'функция делает перенос линии
    Public Shared Function moveLine(ByRef acLine As DwgLine, ByVal deltaLenght As Double, Optional round As Integer = 3) As Boolean
        If IsNothing(acLine) = True Then Return False
        If deltaLenght = 0 Then Return True
        Dim L As Double = acLine.Length
        If L > 0 Then
            Try
                Dim k As Double = deltaLenght / L
                Dim X1 As Double = acLine.StartPoint.X + k * (acLine.EndPoint.X - acLine.StartPoint.X)
                Dim Y1 As Double = acLine.StartPoint.Y + k * (acLine.EndPoint.Y - acLine.StartPoint.Y)
                Dim z1 As Double = acLine.StartPoint.Z + k * (acLine.EndPoint.Z - acLine.StartPoint.Z)
                Dim startPoint As Vector3D = New Vector3D(X1, Y1, z1)

                Dim X2 As Double = acLine.EndPoint.X + k * (acLine.EndPoint.X - acLine.StartPoint.X)
                Dim Y2 As Double = acLine.EndPoint.Y + k * (acLine.EndPoint.Y - acLine.StartPoint.Y)
                Dim z2 As Double = acLine.EndPoint.Z + k * (acLine.EndPoint.Z - acLine.StartPoint.Z)
                Dim endPoint As Vector3D = New Vector3D(X2, Y2, z2)

                acLine.StartPoint = startPoint
                acLine.EndPoint = endPoint
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
        Return True
    End Function
    '==========================================================================================================
    'функция переносит линию к заданной точке с сохранением ее направления и длины
    Public Shared Function moveLineToPoint(ByRef userline As DwgLine, ByVal position As Vector2D, Optional round As Integer = 3) As Boolean
        If IsNothing(userline) = True Then Return False
        If round < 0 OrElse round > 15 Then Return False
        If IsNothing(position) = True Then Return False
        If Double.IsNaN(position.X) OrElse Double.IsInfinity(position.X) OrElse
           Double.IsNaN(position.Y) OrElse Double.IsInfinity(position.Y) Then Return False

        Dim startPoint As Vector3D = Nothing
        Dim endPoint As Vector3D = Nothing
        Dim originalPointsRead As Boolean = False
        Dim mutationStarted As Boolean = False
        Try
            startPoint = userline.StartPoint
            endPoint = userline.EndPoint
            originalPointsRead = True
            If Double.IsNaN(startPoint.X) OrElse Double.IsInfinity(startPoint.X) OrElse
               Double.IsNaN(startPoint.Y) OrElse Double.IsInfinity(startPoint.Y) OrElse
               Double.IsNaN(startPoint.Z) OrElse Double.IsInfinity(startPoint.Z) OrElse
               Double.IsNaN(endPoint.X) OrElse Double.IsInfinity(endPoint.X) OrElse
               Double.IsNaN(endPoint.Y) OrElse Double.IsInfinity(endPoint.Y) OrElse
               Double.IsNaN(endPoint.Z) OrElse Double.IsInfinity(endPoint.Z) Then Return False

            Dim roundedX As Double = Math.Round(position.X, round)
            Dim roundedY As Double = Math.Round(position.Y, round)
            Dim deltaX As Double = roundedX - startPoint.X
            Dim deltaY As Double = roundedY - startPoint.Y
            Dim newStartPoint As Vector3D = New Vector3D(startPoint.X + deltaX, startPoint.Y + deltaY, startPoint.Z)
            Dim newEndPoint As Vector3D = New Vector3D(endPoint.X + deltaX, endPoint.Y + deltaY, endPoint.Z)
            If Double.IsNaN(newStartPoint.X) OrElse Double.IsInfinity(newStartPoint.X) OrElse
               Double.IsNaN(newStartPoint.Y) OrElse Double.IsInfinity(newStartPoint.Y) OrElse
               Double.IsNaN(newStartPoint.Z) OrElse Double.IsInfinity(newStartPoint.Z) OrElse
               Double.IsNaN(newEndPoint.X) OrElse Double.IsInfinity(newEndPoint.X) OrElse
               Double.IsNaN(newEndPoint.Y) OrElse Double.IsInfinity(newEndPoint.Y) OrElse
               Double.IsNaN(newEndPoint.Z) OrElse Double.IsInfinity(newEndPoint.Z) Then Return False

            Dim vectorErrorX As Double = (newEndPoint.X - newStartPoint.X) - (endPoint.X - startPoint.X)
            Dim vectorErrorY As Double = (newEndPoint.Y - newStartPoint.Y) - (endPoint.Y - startPoint.Y)
            If Double.IsNaN(vectorErrorX) OrElse Double.IsInfinity(vectorErrorX) OrElse
               Double.IsNaN(vectorErrorY) OrElse Double.IsInfinity(vectorErrorY) OrElse
               Math.Abs(vectorErrorX) > 0.0005 OrElse Math.Abs(vectorErrorY) > 0.0005 Then Return False

            mutationStarted = True
            userline.StartPoint = newStartPoint
            userline.EndPoint = newEndPoint
            Return True
        Catch ex As Exception
            If originalPointsRead = True AndAlso mutationStarted = True Then
                Try
                    userline.StartPoint = startPoint
                Catch
                End Try
                Try
                    userline.EndPoint = endPoint
                Catch
                End Try
            End If
            Return False
        End Try
    End Function
    '==========================================================================================================
    'функция создает новую линию со смещением, перпендикулярным исходной линии в 3D
    Public Shared Function createPerpendicularOffsetLine(ByVal sourceLine As DwgLine, ByVal offsetHeight As Double) As DwgLine
        If IsNothing(sourceLine) = True Then Return Nothing
        If Double.IsNaN(offsetHeight) OrElse Double.IsInfinity(offsetHeight) Then Return Nothing

        Dim resultLine As DwgLine = Nothing
        Try
            Dim sourceStartPoint As Vector3D = sourceLine.StartPoint
            Dim sourceEndPoint As Vector3D = sourceLine.EndPoint
            If Double.IsNaN(sourceStartPoint.X) OrElse Double.IsInfinity(sourceStartPoint.X) OrElse
               Double.IsNaN(sourceStartPoint.Y) OrElse Double.IsInfinity(sourceStartPoint.Y) OrElse
               Double.IsNaN(sourceStartPoint.Z) OrElse Double.IsInfinity(sourceStartPoint.Z) OrElse
               Double.IsNaN(sourceEndPoint.X) OrElse Double.IsInfinity(sourceEndPoint.X) OrElse
               Double.IsNaN(sourceEndPoint.Y) OrElse Double.IsInfinity(sourceEndPoint.Y) OrElse
               Double.IsNaN(sourceEndPoint.Z) OrElse Double.IsInfinity(sourceEndPoint.Z) Then Return Nothing

            Dim directionX As Double = sourceEndPoint.X - sourceStartPoint.X
            Dim directionY As Double = sourceEndPoint.Y - sourceStartPoint.Y
            Dim directionZ As Double = sourceEndPoint.Z - sourceStartPoint.Z
            If Double.IsNaN(directionX) OrElse Double.IsInfinity(directionX) OrElse
               Double.IsNaN(directionY) OrElse Double.IsInfinity(directionY) OrElse
               Double.IsNaN(directionZ) OrElse Double.IsInfinity(directionZ) Then Return Nothing

            Dim maxDirectionComponent As Double = Math.Max(Math.Max(Math.Abs(directionX), Math.Abs(directionY)), Math.Abs(directionZ))
            If maxDirectionComponent = 0 Then Return Nothing

            Dim scaledDirectionX As Double = directionX / maxDirectionComponent
            Dim scaledDirectionY As Double = directionY / maxDirectionComponent
            Dim scaledDirectionZ As Double = directionZ / maxDirectionComponent
            Dim scaledDirectionLength As Double = Math.Sqrt(scaledDirectionX * scaledDirectionX +
                                                            scaledDirectionY * scaledDirectionY +
                                                            scaledDirectionZ * scaledDirectionZ)
            Dim directionUnitZ As Double = scaledDirectionZ / scaledDirectionLength
            Dim planMaxComponent As Double = Math.Max(Math.Abs(directionX), Math.Abs(directionY))
            Dim deltaX As Double
            Dim deltaY As Double
            Dim deltaZ As Double

            If planMaxComponent > 0 Then
                Dim scaledPlanX As Double = directionX / planMaxComponent
                Dim scaledPlanY As Double = directionY / planMaxComponent
                Dim scaledPlanLength As Double = Math.Sqrt(scaledPlanX * scaledPlanX + scaledPlanY * scaledPlanY)
                Dim deltaPlan As Double = -offsetHeight * directionUnitZ
                deltaX = deltaPlan * scaledPlanX / scaledPlanLength
                deltaY = deltaPlan * scaledPlanY / scaledPlanLength
                deltaZ = offsetHeight * (planMaxComponent / maxDirectionComponent) * scaledPlanLength / scaledDirectionLength
            Else
                deltaX = -offsetHeight * directionUnitZ
                deltaY = 0
                deltaZ = 0
            End If

            Dim resultStartPoint As Vector3D = New Vector3D(sourceStartPoint.X + deltaX,
                                                            sourceStartPoint.Y + deltaY,
                                                            sourceStartPoint.Z + deltaZ)
            Dim resultEndPoint As Vector3D = New Vector3D(sourceEndPoint.X + deltaX,
                                                          sourceEndPoint.Y + deltaY,
                                                          sourceEndPoint.Z + deltaZ)
            If Double.IsNaN(resultStartPoint.X) OrElse Double.IsInfinity(resultStartPoint.X) OrElse
               Double.IsNaN(resultStartPoint.Y) OrElse Double.IsInfinity(resultStartPoint.Y) OrElse
               Double.IsNaN(resultStartPoint.Z) OrElse Double.IsInfinity(resultStartPoint.Z) OrElse
               Double.IsNaN(resultEndPoint.X) OrElse Double.IsInfinity(resultEndPoint.X) OrElse
               Double.IsNaN(resultEndPoint.Y) OrElse Double.IsInfinity(resultEndPoint.Y) OrElse
               Double.IsNaN(resultEndPoint.Z) OrElse Double.IsInfinity(resultEndPoint.Z) Then Return Nothing

            resultLine = New DwgLine()
            resultLine.StartPoint = resultStartPoint
            resultLine.EndPoint = resultEndPoint
            Return resultLine
        Catch ex As Exception
            If IsNothing(resultLine) = False Then
                Try
                    resultLine.Dispose()
                Catch
                End Try
            End If
            Return Nothing
        End Try
    End Function
    '==========================================================================================================
    'функция перемещает ось опоры вдоль трассы на заданный пикет
    Public Shared Function moveLineToAlignmentPK(ByVal align As Alignment, ByRef axisPillar As DwgLine, ByVal station As Double) As Boolean
        moveLineToAlignmentPK = False
        Dim newTempStartPoint As Vector2D = New Vector2D()
        Dim boolPk As Boolean = align.Plan.CompoundLine.StaOffsetToPos(station, 0, newTempStartPoint)
        If boolPk = True Then
            Dim angle As Double = axisPillar.Rotation
            Dim reverseAngle As Double = angle + Math.PI
            If reverseAngle >= Math.PI * 2 Then
                reverseAngle -= Math.PI * 2
            End If
            'создаем произвольный вектор 
            Dim distStart As Double = (newTempStartPoint - axisPillar.StartPoint.Pos).Length
            Dim pt1 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, angle, distStart)
            Dim distend As Double = (newTempStartPoint - axisPillar.EndPoint.Pos).Length
            Dim pt2 As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(newTempStartPoint, reverseAngle, distend)
            axisPillar.StartPoint = New Vector3D(pt1, axisPillar.StartPoint.Z)
            axisPillar.EndPoint = New Vector3D(pt2, axisPillar.EndPoint.Z)
            Return True
        End If
    End Function
    'функция проверяет лежит ли точка на прямой
    Public Shared Function crossPointInLine(ByVal startPoint As Vector3D, ByVal endPoint As Vector3D, ByVal middlePoint As Vector3D) As Boolean
        ' Точность в 1 мм (0.001 метра)
        Const tolerance As Double = 0.001
        ' Вектор от startPoint к endPoint
        Dim lineVector As Vector3D = endPoint - startPoint
        ' Вектор от startPoint к middlePoint
        Dim pointVector As Vector3D = middlePoint - startPoint
        ' Если начальная и конечная точки совпадают
        If lineVector.Length < tolerance Then
            Return pointVector.Length < tolerance
        End If
        ' Вычисляем проекцию pointVector на lineVector
        Dim dotProduct As Double = pointVector.X * lineVector.X +
                               pointVector.Y * lineVector.Y +
                               pointVector.Z * lineVector.Z
        Dim lineLengthSquared As Double = lineVector.X * lineVector.X +
                                      lineVector.Y * lineVector.Y +
                                      lineVector.Z * lineVector.Z
        ' Параметрический коэффициент t
        Dim t As Double = dotProduct / lineLengthSquared
        Dim parameterTolerance As Double = tolerance / lineVector.Length
        ' Проверяем, что точка находится между startPoint и endPoint (необязательно)
        ' Если нужна точка только на отрезке, раскомментируйте следующие строки
        If t < -parameterTolerance OrElse t > 1 + parameterTolerance Then
            Return False
        End If
        ' Находим ближайшую точку на прямой
        Dim closestPoint As Vector3D
        closestPoint.X = startPoint.X + t * lineVector.X
        closestPoint.Y = startPoint.Y + t * lineVector.Y
        closestPoint.Z = startPoint.Z + t * lineVector.Z
        ' Вычисляем расстояние от middlePoint до ближайшей точки на прямой
        Dim distance As Double = Math.Sqrt(Math.Pow(middlePoint.X - closestPoint.X, 2) +
                                       Math.Pow(middlePoint.Y - closestPoint.Y, 2) +
                                       Math.Pow(middlePoint.Z - closestPoint.Z, 2))
        ' Проверяем, находится ли точка на прямой с заданной точностью
        Return distance <= tolerance
    End Function
    'функцтя восстанавливает прямоугольник по двум взаимно пересекаемым линиям
    Public Shared Function calculatePointBox(pt1 As Vector2D, pt2 As Vector2D, pt3 As Vector2D, pt4 As Vector2D) As List(Of Vector2D)
        Dim result As New List(Of Vector2D)()
        ' Находим точку пересечения отрезков
        Dim boolIntersect As Boolean = False
        Dim center As Vector2D = MathFunction.FuncFindLineIntersection(pt1, pt2, pt3, pt4, boolIntersect)
        If boolIntersect = False Then
            Return result
        End If

        ' Определяем направления от центра к концам первого отрезка
        Dim dir1_1 As New Vector2D(pt1.X - center.X, pt1.Y - center.Y)
        Dim dir1_2 As New Vector2D(pt2.X - center.X, pt2.Y - center.Y)

        ' Определяем направления от центра к концам второго отрезка
        Dim dir2_1 As New Vector2D(pt3.X - center.X, pt3.Y - center.Y)
        Dim dir2_2 As New Vector2D(pt4.X - center.X, pt4.Y - center.Y)

        ' Выбираем положительные направления (от центра к концам)
        Dim axis1 As Vector2D = dir1_1
        Dim axis2 As Vector2D = dir2_1

        ' Проверяем, нужно ли развернуть (чтобы получить полные размеры)
        ' Берем максимальное расстояние в каждую сторону
        Dim dist1_1 As Double = Math.Sqrt(dir1_1.X * dir1_1.X + dir1_1.Y * dir1_1.Y)
        Dim dist1_2 As Double = Math.Sqrt(dir1_2.X * dir1_2.X + dir1_2.Y * dir1_2.Y)
        Dim dist2_1 As Double = Math.Sqrt(dir2_1.X * dir2_1.X + dir2_1.Y * dir2_1.Y)
        Dim dist2_2 As Double = Math.Sqrt(dir2_2.X * dir2_2.X + dir2_2.Y * dir2_2.Y)

        ' Половины сторон прямоугольника
        Dim halfWidth As Double = Math.Max(dist1_1, dist1_2)
        Dim halfHeight As Double = Math.Max(dist2_1, dist2_2)

        ' Направление первой оси (горизонтальной)
        Dim horDir As Vector2D
        If dist1_1 >= dist1_2 Then
            horDir = New Vector2D(dir1_1.X / dist1_1, dir1_1.Y / dist1_1)
        Else
            horDir = New Vector2D(dir1_2.X / dist1_2, dir1_2.Y / dist1_2)
        End If

        ' Направление второй оси (вертикальной) - перпендикулярно
        Dim verDir As Vector2D
        If dist2_1 >= dist2_2 Then
            verDir = New Vector2D(dir2_1.X / dist2_1, dir2_1.Y / dist2_1)
        Else
            verDir = New Vector2D(dir2_2.X / dist2_2, dir2_2.Y / dist2_2)
        End If

        ' Дополнительная проверка: убеждаемся, что оси перпендикулярны
        ' Если нет - принудительно поворачиваем вертикальную ось
        Dim dot As Double = horDir.X * verDir.X + horDir.Y * verDir.Y
        If Math.Abs(dot) > 0.0001 Then
            ' Поворачиваем горизонтальную ось на 90 градусов для получения вертикальной
            verDir = New Vector2D(-horDir.Y, horDir.X)
        End If

        ' Вычисляем 4 вершины прямоугольника
        Dim v1 As New Vector2D(center.X + horDir.X * halfWidth + verDir.X * halfHeight,
                           center.Y + horDir.Y * halfWidth + verDir.Y * halfHeight)

        Dim v2 As New Vector2D(center.X + horDir.X * halfWidth - verDir.X * halfHeight,
                           center.Y + horDir.Y * halfWidth - verDir.Y * halfHeight)

        Dim v3 As New Vector2D(center.X - horDir.X * halfWidth - verDir.X * halfHeight,
                           center.Y - horDir.Y * halfWidth - verDir.Y * halfHeight)

        Dim v4 As New Vector2D(center.X - horDir.X * halfWidth + verDir.X * halfHeight,
                           center.Y - horDir.Y * halfWidth + verDir.Y * halfHeight)

        ' Упорядочиваем вершины в порядке обхода (по часовой стрелке)
        result.Add(v1)
        result.Add(v2)
        result.Add(v3)
        result.Add(v4)
        Return result
    End Function
    '============================================================================================================
    'функция вычисояет координаты точки перпердикулярно вверх
    Public Shared Function calculatePointP2(ByVal startPoint As Vector3D, ByVal endPoint As Vector3D, ByVal lenght As Double, Optional angle As Double = Math.PI / 2) As Vector3D
        Dim calculatePoint3d As Vector3D = New Vector3D(0, 0, 0)
        '1.делаем смещение балки вверх
        Dim b As Double = endPoint.Z - startPoint.Z
        Dim c As Double = (endPoint - startPoint).Length
        Dim tempLine As DwgLine = New DwgLine()
        tempLine.StartPoint = startPoint
        tempLine.EndPoint = endPoint
        Dim i As Double = Math.Asin(b / c)
        Dim insCoord As Vector2D = New Vector2D(0, 0)
        Dim deltaXZ As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(insCoord, i + angle, lenght)
        'получаем новые координаты верха балки
        Dim calculatePoint2d As Vector2D = MathFunction.funcCalcCoordinatesByInsPointAndAngle(startPoint.Pos, tempLine.Rotation, deltaXZ.X)
        'создаем верх балки
        calculatePoint3d = New Vector3D(calculatePoint2d.X, calculatePoint2d.Y, startPoint.Z + deltaXZ.Y)
        Return calculatePoint3d
    End Function
    Public Shared Function calculatePointP(ByVal startPoint As Vector3D, ByVal endPoint As Vector3D, ByVal length As Double, Optional ByVal zOffset As Double = 0, Optional ByVal angle As Double = Math.PI / 2) As Vector3D
        ' 1. Получаем направление линии в горизонтальной плоскости (X-Y)
        Dim dx As Double = endPoint.X - startPoint.X
        Dim dy As Double = endPoint.Y - startPoint.Y
        Dim horizontalDir As New Vector3D(dx, dy, 0)

        ' 2. Если линия вертикальная (dx=0 и dy=0), берем любое направление
        If horizontalDir.Length < 0.0001 Then
            horizontalDir = New Vector3D(1, 0, 0)
        Else
            horizontalDir.Normalize()
        End If

        ' 3. Находим перпендикуляр в горизонтальной плоскости (поворот на 90 градусов)
        ' Перпендикулярный вектор: (-dy, dx, 0)
        Dim perpendicularDir As New Vector3D(-horizontalDir.Y, horizontalDir.X, 0)

        ' 4. Если нужен угол, отличный от 90 градусов, поворачиваем
        Dim rotationAngle As Double = angle - Math.PI / 2
        If Math.Abs(rotationAngle) > 0.0001 Then
            Dim cosA As Double = Math.Cos(rotationAngle)
            Dim sinA As Double = Math.Sin(rotationAngle)

            ' Исходное направление
            Dim originalX As Double = perpendicularDir.X
            Dim originalY As Double = perpendicularDir.Y

            ' Поворачиваем вектор
            perpendicularDir.X = originalX * cosA - originalY * sinA
            perpendicularDir.Y = originalX * sinA + originalY * cosA
        End If

        ' 5. Нормализуем перпендикулярный вектор
        perpendicularDir.Normalize()

        ' 6. Вычисляем точку:
        ' - По горизонтали отступаем на length
        ' - По вертикали добавляем zOffset
        Dim result As New Vector3D()
        result.X = startPoint.X + perpendicularDir.X * length
        result.Y = startPoint.Y + perpendicularDir.Y * length
        result.Z = startPoint.Z + zOffset  ' Изменяем Z координату
        Return result
    End Function
    'функция вставляет прямоугольник
    Public Shared Function createRotatedRectangle(ByVal centerPoint As Topomatic.Cad.Foundation.Vector2D, ByVal Length As Double, width As Double, rotation As Double) As List(Of Topomatic.Cad.Foundation.Vector2D)
        Dim points As New List(Of Topomatic.Cad.Foundation.Vector2D)
        ' Половины размеров
        Dim halfLength As Double = Length / 2.0F
        Dim halfWidth As Double = width / 2.0F
        ' Исходные углы прямоугольника (до поворота, относительно центра)
        Dim corners As Topomatic.Cad.Foundation.Vector2D() = {
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, -halfWidth),  ' Верхний левый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, -halfWidth),   ' Верхний правый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, halfWidth),    ' Нижний правый
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, halfWidth)    ' Нижний левый
        }
        ' Поворачиваем и смещаем каждую точку
        For Each corner As Topomatic.Cad.Foundation.Vector2D In corners
            ' Поворот точки
            Dim rotatedX As Double = corner.X * Math.Cos(rotation) - corner.Y * Math.Sin(rotation)
            Dim rotatedY As Double = corner.X * Math.Sin(rotation) + corner.Y * Math.Cos(rotation)
            ' Смещение к центру
            Dim finalPoint As New Topomatic.Cad.Foundation.Vector2D(centerPoint.X + rotatedX, centerPoint.Y + rotatedY)
            points.Add(finalPoint)
        Next
        Return points
    End Function
    'функция вставляет прямоугольник
    Public Shared Function calculateRotatedRectangle3d(ByVal centerPoint As Topomatic.Cad.Foundation.Vector3D, ByVal Length As Double, width As Double, rotation As Double) As List(Of Topomatic.Cad.Foundation.Vector3D)
        Dim points As New List(Of Topomatic.Cad.Foundation.Vector3D)
        ' Половины размеров
        Dim halfLength As Double = Length / 2.0F
        Dim halfWidth As Double = width / 2.0F
        ' Исходные углы прямоугольника (до поворота, относительно центра)
        Dim corners As Topomatic.Cad.Foundation.Vector2D() = {
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, -halfWidth),  ' Верхний левый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, -halfWidth),   ' Верхний правый
            New Topomatic.Cad.Foundation.Vector2D(halfLength, halfWidth),    ' Нижний правый
            New Topomatic.Cad.Foundation.Vector2D(-halfLength, halfWidth)    ' Нижний левый
        }
        ' Поворачиваем и смещаем каждую точку
        For Each corner As Topomatic.Cad.Foundation.Vector2D In corners
            ' Поворот точки
            Dim rotatedX As Double = corner.X * Math.Cos(rotation) - corner.Y * Math.Sin(rotation)
            Dim rotatedY As Double = corner.X * Math.Sin(rotation) + corner.Y * Math.Cos(rotation)
            ' Смещение к центру
            Dim finalPoint As New Topomatic.Cad.Foundation.Vector3D(centerPoint.X + rotatedX, centerPoint.Y + rotatedY, centerPoint.Z)
            points.Add(finalPoint)
        Next
        Return points
    End Function

    ' Вспомогательная функция для поворота вектора вокруг линии
    Private Shared Function RotateAroundLine(ByVal vector As Vector3D, ByVal axis As Vector3D, ByVal angle As Double) As Vector3D
        ' Нормализуем ось вращения
        Dim normalizedAxis As Vector3D = axis
        normalizedAxis.Normalize()

        ' Разлагаем вектор на компоненты параллельные и перпендикулярные оси
        Dim parallelComponent As Vector3D = normalizedAxis * DotProduct(vector, normalizedAxis)
        Dim perpendicularComponent As Vector3D = vector - parallelComponent

        ' Если перпендикулярная компонента почти нулевая, возвращаем исходный вектор
        If perpendicularComponent.Length < 0.0001 Then
            Return vector
        End If

        ' Вычисляем третий базисный вектор для поворота
        Dim thirdAxis As Vector3D = CrossProduct(normalizedAxis, perpendicularComponent)
        thirdAxis.Normalize()

        ' Поворачиваем перпендикулярную компоненту
        Dim cosAngle As Double = Math.Cos(angle)
        Dim sinAngle As Double = Math.Sin(angle)

        Dim rotatedPerpendicular As Vector3D = perpendicularComponent * cosAngle + thirdAxis * sinAngle * perpendicularComponent.Length

        ' Собираем результат
        Return parallelComponent + rotatedPerpendicular
    End Function
    Public Shared Function CrossProduct(ByVal vector1 As Vector3D, ByVal vector2 As Vector3D) As Vector3D
        ' Вычисляем векторное произведение по формуле:
        ' result = vector1 × vector2 = (y1*z2 - z1*y2, z1*x2 - x1*z2, x1*y2 - y1*x2)

        Dim result As New Vector3D()

        result.X = vector1.Y * vector2.Z - vector1.Z * vector2.Y
        result.Y = vector1.Z * vector2.X - vector1.X * vector2.Z
        result.Z = vector1.X * vector2.Y - vector1.Y * vector2.X

        Return result
    End Function
    Public Shared Function DotProduct(ByVal vector1 As Vector3D, ByVal vector2 As Vector3D) As Double
        ' Вычисляем скалярное произведение по формуле:
        ' result = x1*x2 + y1*y2 + z1*z2

        Return vector1.X * vector2.X + vector1.Y * vector2.Y + vector1.Z * vector2.Z
    End Function
    Public Shared Function calculateVertexPolygon(ByVal centerPoint As Topomatic.Cad.Foundation.Vector3D, ByVal radius As Double, ByVal countVertex As Integer, rotation As Double) As List(Of Topomatic.Cad.Foundation.Vector3D)
        Dim vertices As New List(Of Topomatic.Cad.Foundation.Vector3D)

        If radius <= 0 OrElse countVertex < 3 Then
            Return vertices
        End If

        Dim angleStep As Double = (2.0R * Math.PI) / countVertex

        For i As Integer = 0 To countVertex - 1
            Dim angle As Double = rotation + angleStep * i
            Dim x As Double = centerPoint.X + radius * Math.Cos(angle)
            Dim y As Double = centerPoint.Y + radius * Math.Sin(angle)

            vertices.Add(New Topomatic.Cad.Foundation.Vector3D(x, y, centerPoint.Z))
        Next

        Return vertices
    End Function
End Class
