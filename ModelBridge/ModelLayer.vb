Imports System
Imports System.Drawing
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Public Class ModelLayer
    Inherits CadViewLayer

    ' Наша модель
    Private m_Model As Model

    ' Класс, отвечающий за выделение объектов
    Private m_SelectionSet As SelectionSet

    ' Guid нашего слоя
    Public Shared ReadOnly ID As Guid = New Guid("{36C745EB-2111-4D44-B4A1-9BE0B7DBD730}")

    Public Sub New()
        ' В качестве класса, отвечающего за выделение объектов мы используем заглушку которая реализует его по умолчанию, в этом случае он не выделяет ничего
        m_SelectionSet = New DefaultSelectionSet(Me)
    End Sub

    ' Для удобства определяем статический метод, позволяющий получить наш слой с видового экрана
    Public Shared Function GetModelLayer(cadView As CadView) As ModelLayer
        ' Сначала проверяем, есть ли наш слой сразу в самом видовом экране
        ' Это возможно, если видовой экран создан отдельно и слой расположен прямо на видовом экране
        Dim layer As ModelLayer = TryCast(cadView(ModelLayer.ID), ModelLayer)

        If layer Is Nothing Then
            ' Теперь проверяем не находится ли слой в составе нескольких слоёв модели
            ' Это наиболее распространённая ситуация
            ' Для этого мы получаем слой, который содержит внутри все слои всех моделей
            Dim multi As MultiLayer = TryCast(cadView(Consts.ModelsLayer), MultiLayer)

            If multi IsNot Nothing Then
                ' После этого получаем текущий слой активной модели
                Dim active As CadViewLayer = multi.ResolveActive()
                ' Проверяем его на соответствие нашему слою
                layer = TryCast(active, ModelLayer)

                If layer Is Nothing Then
                    ' Кроме того возможен вариант, что у нашей модели несколько слоев
                    ' В этом случае они объединяются внутри общего слоя модели, который и будет являться активным
                    Dim compound As CompoundLayer = TryCast(active, CompoundLayer)

                    If compound IsNot Nothing Then
                        layer = TryCast(compound(ModelLayer.ID), ModelLayer)
                    End If
                End If
            End If
        End If

        ' Если наш слой найден, но он заблокирован на редактирование, то мы не можем его вернуть
        If (layer IsNot Nothing) AndAlso (Not layer.ResolveEnable()) Then
            Return Nothing
        End If

        Return layer
    End Function

    ' Возвращаем в качестве LayerId ID объявленный выше
    Public Overrides ReadOnly Property LayerGuid As Guid
        Get
            Return ID
        End Get
    End Property

    ' Возвращаем нашу заглушку
    Public Overrides ReadOnly Property SelectionSet As SelectionSet
        Get
            Return m_SelectionSet
        End Get
    End Property

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Слой тестовой модели"
        End Get
    End Property

    Public Property Model As Model
        Get
            Return m_Model
        End Get
        Set(value As Model)
            m_Model = value
        End Set
    End Property

    ' Рассчитываем границы слоя
    Protected Overrides Function OnGetLimits(ByRef limits As BoundingBox2D) As Boolean
        ' Если есть точки в модели
        If m_Model.Points.Count > 0 Then
            ' Создаем рамку вокруг первой точки
            limits = New BoundingBox2D(m_Model.Points(0), m_Model.Points(0))

            For i As Integer = 1 To m_Model.Points.Count - 1
                ' И добавляем в нее все остальные точки
                limits.AddPoint(m_Model.Points(i))
            Next

            Return True
        End If

        ' Если точек в модели нет, возвращаем пустую рамку
        limits = BoundingBox2D.Empty
        Return False
    End Function

    Protected Overrides Sub OnGetSnapObjects(e As ObjectSnapEventArgs)
        ' Поскольку мы не реализуем привязки, то здесь мы не делаем ничего
    End Sub

    Protected Overrides Sub OnPaint(pen As CadPen)
        ' Рисуем нашу линию жёлтым цветом
        pen.Color = Color.Yellow

        ' Начинаем рисовать
        pen.BeginDraw()

        Try
            ' Для отрисовки используем возможность нарисовать массив нескольких точек
            ' Для этого вызываем начало отрисовки массива
            pen.BeginArray()

            For i As Integer = 0 To m_Model.Points.Count - 1
                ' Добавляем точки
                pen.Vertex(m_Model.Points(i).X, m_Model.Points(i).Y, 0)
            Next

            ' Заканчиваем отрисовку массива, в виде линии
            pen.EndArray(ArrayMode.Polyline)
        Finally
            ' Заканчиваем рисовать
            pen.EndDraw()
        End Try

        ' В каждой точке пишем номер оранжевым цветом
        pen.Color = Color.Orange

        ' Начинаем рисовать
        pen.BeginDraw()

        Try
            Dim font = FontManager.Current.DefaultFont
            For i As Integer = 0 To m_Model.Points.Count - 1
                ' Пишем номер точки, высотой в 2 единицы чертежа
                font.DrawString(i.ToString(), pen, m_Model.Points(i), 0.0, 2.0)
            Next
        Finally
            ' Заканчиваем рисовать
            pen.EndDraw()
        End Try
    End Sub
End Class
