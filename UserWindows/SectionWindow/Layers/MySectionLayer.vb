Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Runtime.Remoting.Messaging
Imports Topomatic.ApplicationPlatform
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Design
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Layer
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
Imports Topomatic.Controls.NativeMethods
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports Topomatic.Dwg.Entities
Imports Topomatic.Sfc
Friend Class MySectionLayer
    Inherits CadViewLayer
    Public Shared ReadOnly GUID As Guid = New Guid("{018D4CF1-1C12-42F3-ADAE-31FF270375A5}")
    Private m_SelectionSet As SelectionSet
    Private m_LastEnitityId As List(Of UInteger)
    Private m_CachedView As Drawing = New Drawing()
    Private elevation As Double = 0

    Public Sub New()
        m_SelectionSet = New MySelectionSet(Me)
    End Sub
    Public Overrides ReadOnly Property LayerGuid As Guid
        Get
            Return GUID
        End Get
    End Property
    Public Overrides ReadOnly Property SelectionSet As SelectionSet
        Get
            Return m_SelectionSet
        End Get
    End Property
    Public Overrides ReadOnly Property Name As String
        Get
            Return "My layer"
        End Get
    End Property

    Protected Overrides Sub OnGetSnapObjects(e As ObjectSnapEventArgs)
        'Throw New NotImplementedException()
    End Sub

    Private Class MySelectionSet
        Inherits SelectionSet
        Private Class MyGrip
            Inherits Grip
            Private m_Layer As MySectionLayer
            Public Sub New(layer As MySectionLayer)
                MyBase.New(layer.CadView)
                m_Layer = layer
                Dim elevation As Double
                m_Layer.GetCachedView(elevation)
                Me.Location = New Vector3D(0.0, elevation, 0.0)

                Dim click_grip As New ClickGrip()
                AddHandler click_grip.Click, Sub(s As Object, e As EventArgs)
                                                 'Действие по выбору из меню грипа
                                             End Sub
                Me.AddGrip("Меню1", "menu1", click_grip)
            End Sub

            'Срабатывает когда перемещение выполнено
            Public Overrides Sub OnMove(vertex As Vector3D)
                MyBase.OnMove(vertex)
                Dim entListModel As List(Of DwgModel3DElement) = m_Layer.FindSelectedEntity()
                If entListModel.Count > 0 Then
                    Dim entity As DwgModel3DElement = entListModel.First
                    entity.Position = New Vector3D(entity.Position.Pos, vertex.Y)
                End If
            End Sub

            'Срабатывает когда объект перемещается под курсором мыши
            Public Overrides Sub OnDynamicRender(dc As DeviceContext, position As Vector3D)
                MyBase.OnDynamicRender(dc, position)
                Dim elevation As Double
                Dim view = m_Layer.GetCachedView(elevation)
                If view IsNot Nothing Then
                    AuxiliaryDrawer.DrawDimmentionLinear3d(dc, New Vector3D(0.0, elevation, 0.0), New Vector3D(0.0, position.Y, 0.0))
                End If
            End Sub
        End Class

        Private m_Layer As MySectionLayer
        Private m_Selected As Boolean

        Public Sub New(layer As MySectionLayer)
            MyBase.New(layer)
            m_Layer = layer
        End Sub

        'Возвращает количество выделенных объектов
        Public Overrides ReadOnly Property Count As Integer
            Get
                Return If(m_Selected, 1, 0)
            End Get
        End Property
        'Очищает список выделенных объектов
        Public Overrides Sub Clear()
            m_Selected = False
        End Sub
        'Удаляет выделенные объекты при нажатии клавиши Del
        Public Overrides Sub [Erase]()
            Throw New NotImplementedException()
        End Sub
        'Возвращает все выделенные объекты
        Public Overrides Iterator Function GetEnumerator() As IEnumerator
            If m_Selected Then
                Dim elevation As Double
                Yield m_Layer.GetCachedView(elevation)
            End If
        End Function
        'Public Overrides Function GetEnumerator() As IEnumerator
        '    If m_Selected Then
        '        Dim elevation As Double
        '        Return m_Layer.GetCachedView(elevation)
        '    End If
        'End Function

        'Определяет объекты под курсором мыши
        'waitTimeOut - максимальное время выделенное на опреацию, если 0 - время не ограничено
        Public Overrides Iterator Function GetObjectsAtPoint(ByVal point As Vector3D, ByVal match As Predicate(Of Object), ByVal waitTimeOut As Integer) As IEnumerable(Of KeyValuePair(Of Double, Object))
            Dim elevation As Double
            Dim view = m_Layer.GetCachedView(elevation)
            If view Is Nothing Then Return

            If (match Is Nothing) OrElse match(view) Then
                Dim pos = New Vector2D(point.X, point.Y - elevation)

                If view.Blocks.ModelSpace.Bounds.Contains(pos) <> ContainmentType.Disjoint Then
                    Yield New KeyValuePair(Of Double, Object)(0.0F, view)
                End If
            End If
        End Function
        'Возвращает список грипов для объекта obj
        Public Overrides Iterator Function GetObjectGrips(obj As Object) As IEnumerable(Of IGrip)
            If TypeOf obj Is Drawing Then
                Yield New MyGrip(m_Layer)
            End If
        End Function
        'Определяет объекты выделенные рамкой
        Public Overrides Sub GetObjectsByFrame(mode As FrameSelectType, rect As RectangleD, match As Predicate(Of Object), action As Action(Of Object))
            Dim elevation As Double
            Dim view As Drawing = m_Layer.GetCachedView(elevation)
            If view Is Nothing Then
                Return
            End If
            If match Is Nothing OrElse match(view) Then
                Dim bounds = BoundingBox2D.Transform(view.Blocks.ModelSpace.Bounds, Matrix.CreateTranslation(0.0, elevation, 0.0))
                Dim rbounds = rect.ToBoundingBox()

                If mode = FrameSelectType.Contains Then
                    If rbounds.Contains(bounds) = ContainmentType.Contains Then
                        action(view)
                    End If
                Else
                    If rbounds.Contains(bounds) <> ContainmentType.Disjoint Then
                        action(view)
                    End If
                End If
            End If
        End Sub
        'Пока не используется
        Public Overrides Sub GetObjectsByPolygon(mode As FrameSelectType, pointsList As List(Of Vector2D), match As Predicate(Of Object), action As Action(Of Object))
            'Throw New NotImplementedException()
        End Sub
        'Возвращает все объекты, которые в принципе возможно выделить
        Public Overrides Iterator Function GetSelectable() As IEnumerable
            Dim elevation As Double
            Dim view = m_Layer.GetCachedView(elevation)
            If view IsNot Nothing Then Yield view
        End Function
        'Определяет можно ли редактировать объект
        Public Overrides Function IsEnable(obj As Object) As Boolean
            'пока объект один нам всё равно
            Return m_Layer.Enable
        End Function
        'Определяет принадлежит ли объект текущему выделению
        'Для всех объектов из GetSelectable обязан возвращать true
        Public Overrides Function IsOwned(obj As Object) As Boolean
            'пока объект один нам всё равно
            Return True
        End Function
        'Определяет выделен ли объект
        Public Overrides Function IsSelected(obj As Object) As Boolean
            'пока объект один нам всё равно
            Return m_Selected
        End Function
        'Выделяет объект или снимает с него выдедение в зависимости от bFlag
        Public Overrides Sub [Select](item As Object, bFlag As Boolean)
            m_Selected = bFlag
        End Sub
    End Class
    Private Sub RefreshCache(entModel As List(Of DwgModel3DElement))
        If entModel.Count = 0 Then
            m_CachedView = Nothing
            m_LastEnitityId.Clear()
        Else
            If m_CachedView Is Nothing Then
                m_CachedView = New Drawing()
            End If
            Dim boolFindId As Boolean = False
            For Each entity As DwgModel3DElement In entModel
                'm_CachedView = New Drawing()
                Dim idModel As UInteger = entity.ObjectID
                If IsNothing(m_LastEnitityId) = False Then
                    For i As Integer = 0 To m_LastEnitityId.Count - 1
                        If m_LastEnitityId(i) = idModel Then
                            boolFindId = True
                            Exit For
                        End If
                    Next
                End If
                If boolFindId = False Then
                    If IsNothing(m_LastEnitityId) = True Then
                        m_LastEnitityId = New List(Of UInteger)
                        m_LastEnitityId.Add(entity.ObjectID)
                    Else
                        m_LastEnitityId.Add(entity.ObjectID)
                    End If
                    Dim e = TryCast(entity.Element, ImViewElement)
                    If e IsNot Nothing Then
                        Dim acLineSection As DwgLine = Bridges.getAxisElementsByModel3d(entity)
                        If IsNothing(acLineSection) = False Then
                            Dim v1 As Vector3D = acLineSection.Delta
                            Dim v2 As Vector3D = New Vector3D(0, v1.Z, -1 * v1.Y)
                            Dim v3 As Vector3D = New Vector3D(v1.Y, -1 * v1.X, 0)
                            If entity.Element.Name Like "Балка 33м" Then
                                Dim newView As Model3DView = New Model3DView(v1, v2)
                                e.GetView(newView, m_CachedView.Blocks.ModelSpace, Matrix.Identity, 1.0)
                            Else
                                Dim newM As Vector2D = New Vector3D(10, 0, 0)
                                e.GetView(Model3DView.Front, m_CachedView.Blocks.ModelSpace, Matrix.CreateTranslation(New Vector3D(20, 10, 0)), 1.0)
                            End If
                        Else
                            e.GetView(Model3DView.Front, m_CachedView.Blocks.ModelSpace, Matrix.Identity, 1.0)
                        End If
                    End If

                    ApplicationHost.Current.InvokeDelayed(5, Sub()
                                                                 CadView.SolveLimits()
                                                                 CadView.Unlock()
                                                                 CadView.Invalidate()
                                                             End Sub, False, True)
                End If
            Next

        End If
    End Sub
    'начало всего
    Protected Function GetCachedView(<Out> ByRef elevation As Double) As Drawing
        Dim entModel As List(Of DwgModel3DElement) = FindSelectedEntity()
        If IsNothing(entModel) = True Then
            elevation = 0
            Return m_CachedView
        End If
        If entModel.Count > 0 Then
            Me.RefreshCache(entModel)
            Dim tempElevation As Double = -9999999
            For i As Integer = 0 To entModel.Count - 1
                Dim acModel As DwgModel3DElement = entModel.Item(i)
                Dim elementPropertiesObject As ImElement = acModel.Element
                Dim generalPropertiesObject1 As ImProperties = elementPropertiesObject.GetProperties()
                Dim boolFindElev As Boolean = False
                If generalPropertiesObject1.Count > 0 Then
                    For k As Integer = 0 To generalPropertiesObject1.Count - 1
                        Dim generalPropLevel1 As ImProperty = generalPropertiesObject1.Item(k)
                        If generalPropLevel1.Name Like "Отметка тчк. опирания А" Then
                            Dim temPElev As Double = generalPropLevel1.Value
                            If temPElev > tempElevation Then
                                tempElevation = temPElev
                                boolFindElev = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If boolFindElev = False Then
                    Dim temPElev As Double = acModel.Position.Elevation
                    If temPElev > tempElevation Then
                        tempElevation = temPElev
                    End If
                End If
            Next i
            elevation = tempElevation
            Return m_CachedView
        Else
            elevation = 0
        End If
    End Function

    Public Sub InvalidateSelection(Optional ent As List(Of DwgModel3DElement) = Nothing)
        If m_SelectionSet.Count > 0 Then
            Dim entModel As List(Of DwgModel3DElement) = Nothing
            If IsNothing(ent) = False Then
                If ent.Count > 0 Then
                    entModel = FindSelectedEntity(ent)

                Else
                    entModel = FindSelectedEntity()
                End If
            Else
                entModel = FindSelectedEntity()
            End If
            If IsNothing(entModel) = True Then
                m_SelectionSet.Clear()
            ElseIf entModel.Count = 0 Then
                m_SelectionSet.Clear()
            End If
            'If (entity Is Nothing) OrElse (entity.ObjectID <> m_LastEnitityId) Then
        Else
            m_CachedView = Nothing
            If IsNothing(m_LastEnitityId) = False Then
                m_LastEnitityId.Clear()
            End If

        End If
        CadView.Unlock()
        CadView.Invalidate()
    End Sub

    Private Function FindSelectedEntity(Optional ent As List(Of DwgModel3DElement) = Nothing) As List(Of DwgModel3DElement)
        Dim cad_view = CadViewDesignUtils.OnCadViewSelect(CadViewDesignUtils.PlanCadViewAlias)
        Dim listSelectionObject As List(Of DwgModel3DElement) = New List(Of DwgModel3DElement)
        If cad_view IsNot Nothing Then
            Dim dwg_layer = DrawingLayer.GetDrawingLayer(cad_view)
            If dwg_layer IsNot Nothing Then
                For Each selected In dwg_layer.SelectionSet
                    If TypeOf selected Is DwgModel3DElement Then
                        listSelectionObject.Add(selected)

                    End If
                Next
                If listSelectionObject.Count > 0 Then
                    Return listSelectionObject
                End If
                If IsNothing(ent) = False Then
                    Return ent
                End If
            End If
        End If
        Return Nothing
    End Function

    Protected Overrides Function OnGetLimits(<Out> ByRef limits As BoundingBox2D) As Boolean
        Dim elevation As Double
        Dim view = GetCachedView(elevation)

        If view Is Nothing Then
            limits = BoundingBox2D.Empty
            Return False
        End If

        limits = BoundingBox2D.Transform(view.Blocks.ModelSpace.Bounds, Matrix.CreateTranslation(0.0, elevation, 0.0))
        Return True
    End Function

    Protected Overrides Sub OnPaint(ByVal pen As CadPen)
        Dim elevation As Double
        Dim view = GetCachedView(elevation)
        If view Is Nothing Then Return
        PaintEntityEventArgs.PaintEntities(Nothing, view.Blocks.ModelSpace, New Vector3D(0.0, elevation, 0.0), New Vector3D(1.0, 1.0, 1.0), 0.0, New PaintEntityEventArgs(Me.Enable, pen))
    End Sub

    Protected Overrides Sub OnHilightObject(ByVal pen As CadPen, ByVal obj As Object)
        MyBase.OnHilightObject(pen, obj)

        If TypeOf obj Is Drawing Then
            pen.DrawingMode = DrawingMode.Highlight
            pen.HighlightMode = HighlightMode.DoubleBlack
            Dim elevation As Double
            Dim view = GetCachedView(elevation)
            If view Is Nothing Then Return
            PaintEntityEventArgs.PaintEntities(Nothing, view.Blocks.ModelSpace, New Vector3D(0.0, elevation, 0.0), New Vector3D(1.0, 1.0, 1.0), 0.0, New PaintEntityEventArgs(Me.Enable, pen))
            pen.Reset()
        End If
    End Sub

    Protected Overrides Sub OnDynamicDraw(ByVal pen As CadPen, ByVal location As Vector3D)
        MyBase.OnDynamicDraw(pen, location)

        If m_SelectionSet.Count > 0 AndAlso Visible Then
            pen.HighlightMode = HighlightMode.Black
            Dim elevation As Double
            Dim view = GetCachedView(elevation)
            If view Is Nothing Then Return
            PaintEntityEventArgs.PaintEntities(Nothing, view.Blocks.ModelSpace, New Vector3D(0.0, elevation, 0.0), New Vector3D(1.0, 1.0, 1.0), 0.0, New PaintEntityEventArgs(Me.Enable, pen))
            pen.Reset()
        End If
    End Sub
End Class
