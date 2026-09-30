Imports System.Collections
Imports System.Drawing.Printing
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Controls
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Dwg.Layer
Imports Topomatic.Stg
Imports Topomatic.Visualization.Runtime

Friend NotInheritable Class BridgeDrawingLayer
    Inherits DrawingLayer

    Private Const ExitBridgeMenuTag As String = "INFRA_EXIT_BRIDGE_EDIT"

    Private ReadOnly _containerState As New BridgeContainerState()
    Private ReadOnly _selectionSet As BridgeSelectionSet
    Private _bridgeIndex As Dictionary(Of String, BridgeTaggedEntityRecord)
    Private _entityBridgeIds As Dictionary(Of DwgEntity, String)
    Private _bridgeIndexDirty As Boolean = True

    Public Sub New(name As String)
        MyBase.New(name)
        _selectionSet = New BridgeSelectionSet(Me, MyBase.SelectionSet)
    End Sub

    Public Overrides ReadOnly Property SelectionSet As SelectionSet
        Get
            Return _selectionSet
        End Get
    End Property

    Public Overrides Property Drawing As Drawing
        Get
            Return MyBase.Drawing
        End Get
        Set(value As Drawing)
            If Object.ReferenceEquals(MyBase.Drawing, value) Then Return

            DetachDrawing(MyBase.Drawing)
            _selectionSet.Clear()
            _containerState.Reset()
            InvalidateBridgeIndex()
            MyBase.Drawing = value
            AttachDrawing(value)
        End Set
    End Property

    Friend Function ResolveSelectionObject(item As Object) As Object
        Dim record As BridgeTaggedEntityRecord = Nothing
        If Not TryGetBridgeRecord(item, record) Then Return item

        Dim result As Object = _containerState.ResolveSelectionObject(item, record.IdStructure)
        Dim presentationObject As BridgePresentationObject = TryCast(result, BridgePresentationObject)
        If presentationObject IsNot Nothing Then presentationObject.Update(record.Bridge)
        Return result
    End Function

    Friend Function IsBridgeItem(item As Object) As Boolean
        Dim record As BridgeTaggedEntityRecord = Nothing
        Return TryGetBridgeRecord(item, record)
    End Function

    Friend Function IsPresentationObjectOwned(item As BridgePresentationObject) As Boolean
        If Not _containerState.Owns(item) Then Return False
        Dim record As BridgeTaggedEntityRecord = Nothing
        If Not TryGetBridgeRecord(item.IdStructure, record) Then Return False
        item.Update(record.Bridge)
        Return True
    End Function

    Friend Function IsPlanSelectable(item As Object) As Boolean
        Dim entity As DwgEntity = GetEntity(item)
        If entity Is Nothing Then Return True

        Dim record As BridgeTaggedEntityRecord = Nothing
        Dim isBridgeEntity As Boolean = TryGetBridgeRecord(entity, record)
        Return BridgePlanVisibilityPolicy.ShouldDrawInPlan(
            isBridgeEntity,
            TypeOf entity Is DwgModel3DElement)
    End Function

    Friend Function IsWholeBridgeInFrame(presentationObject As BridgePresentationObject,
                                         hits As HashSet(Of Object)) As Boolean
        Dim record As BridgeTaggedEntityRecord = Nothing
        If Not TryGetBridgeRecord(presentationObject.IdStructure, record) Then Return False

        Dim hasSelectableEntity As Boolean = False
        For Each entity As DwgEntity In record.Entities
            If Not IsEntityInActiveSpace(entity) OrElse
               Not entity.IsVisible OrElse
               Not IsPlanSelectable(entity) Then
                Continue For
            End If

            hasSelectableEntity = True
            If Not hits.Contains(entity) Then Return False
        Next
        Return hasSelectableEntity
    End Function

    Protected Overrides Sub OnPaint(pen As CadPen)
        If Drawing Is Nothing OrElse Drawing.ActiveLayout Is Nothing Then Return
        If Not Drawing.ActiveLayout.IsModel Then
            MyBase.OnPaint(pen)
            Return
        End If

        PaintPlan(pen, includeSelectionHighlight:=True)
    End Sub

    Protected Overrides Sub OnPrint(pen As CadPen, e As PrintPageEventArgs)
        If Drawing Is Nothing OrElse Drawing.ActiveLayout Is Nothing Then Return
        If Not Drawing.ActiveLayout.IsModel Then
            MyBase.OnPrint(pen, e)
            Return
        End If

        PrintPlan(pen)
    End Sub

    Protected Overrides Sub OnHilightObject(pen As CadPen, obj As Object)
        Dim presentationObject As BridgePresentationObject = TryCast(obj, BridgePresentationObject)
        If presentationObject IsNot Nothing Then
            Dim record As BridgeTaggedEntityRecord = Nothing
            If TryGetBridgeRecord(presentationObject.IdStructure, record) Then
                presentationObject.Update(record.Bridge)
                PaintHighlightedPlanEntities(pen, PlanEntities(record.Entities))
            End If
            Return
        End If

        If Not IsPlanSelectable(obj) Then Return
        MyBase.OnHilightObject(pen, obj)
    End Sub

    Protected Overrides Sub OnHilightObject3d(dc As DeviceContext, obj As Object)
        If TypeOf obj Is BridgePresentationObject Then Return
        MyBase.OnHilightObject3d(dc, obj)
    End Sub

    Protected Overrides Sub OnMenuAction(e As CreateMenuEventArgs)
        MyBase.OnMenuAction(e)
        If e Is Nothing OrElse e.Root Is Nothing Then Return

        If Not String.IsNullOrEmpty(_containerState.OpenBridgeId) Then
            e.Root.Add("Завершить редактирование моста", AddressOf HandleBridgeMenuAction, ExitBridgeMenuTag)
            Return
        End If

        Dim presentationObject As BridgePresentationObject = SelectedPresentationObject()
        If presentationObject IsNot Nothing Then
            e.Root.Add("Редактировать мост", AddressOf HandleBridgeMenuAction, presentationObject)
        End If
    End Sub

    Private Sub HandleBridgeMenuAction(sender As Object, e As MenuActionEventArgs)
        If e Is Nothing Then Return

        If String.Equals(TryCast(e.Tag, String), ExitBridgeMenuTag, StringComparison.Ordinal) Then
            _selectionSet.Clear()
            _containerState.ExitBridge()
        Else
            Dim presentationObject As BridgePresentationObject = TryCast(e.Tag, BridgePresentationObject)
            If presentationObject Is Nothing OrElse Not IsPresentationObjectOwned(presentationObject) Then Return
            _selectionSet.Clear()
            _containerState.EnterBridge(presentationObject.IdStructure)
        End If

        If CadView IsNot Nothing Then
            CadView.Unlock()
            CadView.Invalidate()
        End If
    End Sub

    Private Function SelectedPresentationObject() As BridgePresentationObject
        For Each item As Object In _selectionSet
            Dim presentationObject As BridgePresentationObject = TryCast(item, BridgePresentationObject)
            If presentationObject IsNot Nothing AndAlso IsPresentationObjectOwned(presentationObject) Then
                Return presentationObject
            End If
        Next
        Return Nothing
    End Function

    Private Sub PaintPlan(pen As CadPen, includeSelectionHighlight As Boolean)
        Dim activeSpace As DwgBlock = Drawing.ActiveSpace
        If activeSpace Is Nothing Then Return

        Dim viewBounds As BoundingBox2D = pen.ViewBounds
        Dim orthoScale As Single = pen.DeviceContext.OrthoScale
        Dim isolinesScale As Single = pen.DeviceContext.IsolinesScale
        Dim drawingMode As DrawingMode = pen.DrawingMode
        Dim highlightMode As HighlightMode = pen.HighlightMode
        Dim linetypeScale As Double = pen.LinetypeScale
        Dim pushed As Boolean = False

        Try
            pen.DeviceContext.PushMatrix()
            pushed = True

            Dim matrix As Matrix = Transform
            Dim inverse As Matrix = Matrix.Invert(matrix)
            pen.DeviceContext.MultMatrix(matrix)
            Dim transformedBounds As BoundingBox2D
            viewBounds.Transform(inverse, transformedBounds)
            pen.ViewBounds = transformedBounds
            pen.DeviceContext.OrthoScale *= CSng(1.0 / Scale)
            pen.DeviceContext.IsolinesScale *= CSng(1.0 / Scale)
            pen.LinetypeScale = linetypeScale * Drawing.GlobalLinetypeScale

            Dim normalEntities As New List(Of DwgEntity)()
            For Each entity As DwgEntity In OrderPlanEntities(PlanEntities(activeSpace))
                If Not includeSelectionHighlight OrElse
                   Not _selectionSet.IsSelected(entity) OrElse
                   (entity.IsBackgroud AndAlso Not TypeOf entity Is BridgeBeamPlanEntity) Then
                    normalEntities.Add(entity)
                End If
            Next
            PaintEntities(pen, normalEntities)

            If includeSelectionHighlight Then
                PaintHighlightedEntities(pen, SelectedPlanEntities())
            End If
        Finally
            pen.LinetypeScale = linetypeScale
            pen.DrawingMode = drawingMode
            pen.HighlightMode = highlightMode
            If pushed Then pen.DeviceContext.PopMatrix()
            pen.ViewBounds = viewBounds
            pen.DeviceContext.OrthoScale = orthoScale
            pen.DeviceContext.IsolinesScale = isolinesScale
        End Try
    End Sub

    Private Sub PrintPlan(pen As CadPen)
        Dim activeSpace As DwgBlock = Drawing.ActiveSpace
        If activeSpace Is Nothing Then Return

        Dim annotationScale As Double = 1.0
        Dim annotationRotation As Double = 0.0
        If CadView IsNot Nothing Then
            annotationScale = CadView.AnnotationScale
            annotationRotation = CadView.ScreenRotation
        End If

        Dim ownerColor As System.Drawing.Color = System.Drawing.Color.White
        If pen.IsWhiteBackColor Then ownerColor = System.Drawing.Color.Black
        Dim disabledOwnerColor As System.Drawing.Color = CadColor.VisualColor(CadColor.White, pen, False)
        Dim args As New PaintEntityEventArgs(Enable, pen) With {
            .AnnotationScale = annotationScale,
            .AnnotationRotation = annotationRotation
        }

        For Each entity As DwgEntity In OrderPlanEntities(activeSpace.Cast(Of DwgEntity)())
            If entity.Layer.Visible AndAlso
               entity.IntersectWith(pen.ViewBounds, annotationScale) AndAlso
               IsPlanSelectable(entity) Then
                args.OwnerEnable = entity.IsEnable
                args.OwnerColor = If(entity.IsEnable, ownerColor, disabledOwnerColor)
                PaintEntityEventArgs.PaintEntity(entity, args)
            End If
        Next
    End Sub

    Private Sub PaintHighlightedEntities(pen As CadPen, entities As IEnumerable(Of DwgEntity))
        Dim drawingMode As DrawingMode = pen.DrawingMode
        Dim highlightMode As HighlightMode = pen.HighlightMode
        Try
            pen.DrawingMode = DrawingMode.Highlight
            pen.HighlightMode = HighlightMode.DoubleBlack
            PaintEntities(pen, OrderPlanEntities(entities))
        Finally
            pen.DrawingMode = drawingMode
            pen.HighlightMode = highlightMode
        End Try
    End Sub

    Private Sub PaintHighlightedPlanEntities(pen As CadPen, entities As IEnumerable(Of DwgEntity))
        Dim viewBounds As BoundingBox2D = pen.ViewBounds
        Dim orthoScale As Single = pen.DeviceContext.OrthoScale
        Dim isolinesScale As Single = pen.DeviceContext.IsolinesScale
        Dim linetypeScale As Double = pen.LinetypeScale
        Dim pushed As Boolean = False

        Try
            pen.DeviceContext.PushMatrix()
            pushed = True
            Dim matrix As Matrix = Transform
            Dim inverse As Matrix = Matrix.Invert(matrix)
            pen.DeviceContext.MultMatrix(matrix)
            Dim transformedBounds As BoundingBox2D
            viewBounds.Transform(inverse, transformedBounds)
            pen.ViewBounds = transformedBounds
            pen.DeviceContext.OrthoScale *= CSng(1.0 / Scale)
            pen.DeviceContext.IsolinesScale *= CSng(1.0 / Scale)
            pen.LinetypeScale = linetypeScale * Drawing.GlobalLinetypeScale
            PaintHighlightedEntities(pen, entities)
        Finally
            pen.LinetypeScale = linetypeScale
            If pushed Then pen.DeviceContext.PopMatrix()
            pen.ViewBounds = viewBounds
            pen.DeviceContext.OrthoScale = orthoScale
            pen.DeviceContext.IsolinesScale = isolinesScale
        End Try
    End Sub

    Private Sub PaintEntities(pen As CadPen, entities As IEnumerable(Of DwgEntity))
        Dim args As New PaintEntityEventArgs(Enable, pen)
        If CadView IsNot Nothing Then
            args.AnnotationScale = CadView.AnnotationScale
            args.AnnotationRotation = CadView.ScreenRotation
        End If
        PaintEntityEventArgs.PaintEntities(
            Nothing,
            entities,
            Vector3D.Empty,
            Vector3D.One,
            0.0,
            args)
    End Sub

    Private Iterator Function SelectedPlanEntities() As IEnumerable(Of DwgEntity)
        Dim yielded As New HashSet(Of DwgEntity)()
        For Each item As Object In _selectionSet
            Dim presentationObject As BridgePresentationObject = TryCast(item, BridgePresentationObject)
            If presentationObject IsNot Nothing Then
                Dim record As BridgeTaggedEntityRecord = Nothing
                If TryGetBridgeRecord(presentationObject.IdStructure, record) Then
                    presentationObject.Update(record.Bridge)
                    For Each entity As DwgEntity In OrderPlanEntities(PlanEntities(record.Entities))
                        If yielded.Add(entity) Then Yield entity
                    Next
                End If
                Continue For
            End If

            Dim selectedEntity As DwgEntity = GetEntity(item)
            If selectedEntity IsNot Nothing AndAlso
               IsEntityInActiveSpace(selectedEntity) AndAlso
               selectedEntity.IsVisible AndAlso
               IsPlanSelectable(selectedEntity) AndAlso
               yielded.Add(selectedEntity) Then
                Yield selectedEntity
            End If
        Next
    End Function

    Private Iterator Function PlanEntities(entities As IEnumerable(Of DwgEntity)) As IEnumerable(Of DwgEntity)
        For Each entity As DwgEntity In entities
            If entity IsNot Nothing AndAlso
               IsEntityInActiveSpace(entity) AndAlso
               entity.IsVisible AndAlso
               IsPlanSelectable(entity) Then
                Yield entity
            End If
        Next
    End Function

    Private Shared Function OrderPlanEntities(entities As IEnumerable(Of DwgEntity)) As IEnumerable(Of DwgEntity)
        Return entities.OrderBy(AddressOf PlanPaintPriority)
    End Function

    Private Shared Function PlanPaintPriority(entity As DwgEntity) As Integer
        If TypeOf entity Is BridgeBeamPlanEntity Then Return 0
        Dim data As StructureElement = Nothing
        If BridgeTaggedEntityIndex.TryReadStructureData(entity, data) Then
            Dim typeValue As Integer = CInt(data.Name)
            If typeValue >= CInt(StructureElement.typeObject.axisBridge) AndAlso
               typeValue <= CInt(StructureElement.typeObject.axisSiteMonolitPillar) Then
                Return 2
            End If
        End If
        Return 1
    End Function

    Private Function IsEntityInActiveSpace(entity As DwgEntity) As Boolean
        Return entity IsNot Nothing AndAlso
               Drawing IsNot Nothing AndAlso
               Object.ReferenceEquals(entity.FindBlock(), Drawing.ActiveSpace)
    End Function

    Private Function TryGetBridgeRecord(item As Object, ByRef record As BridgeTaggedEntityRecord) As Boolean
        Dim presentationObject As BridgePresentationObject = TryCast(item, BridgePresentationObject)
        If presentationObject IsNot Nothing Then
            Return TryGetBridgeRecord(presentationObject.IdStructure, record)
        End If

        Dim entity As DwgEntity = GetEntity(item)
        If entity Is Nothing OrElse Not Object.ReferenceEquals(entity.Drawing, Drawing) Then Return False

        EnsureBridgeIndex()
        Dim canonicalId As String = Nothing
        Return _entityBridgeIds.TryGetValue(entity, canonicalId) AndAlso
               _bridgeIndex.TryGetValue(canonicalId, record)
    End Function

    Private Function TryGetBridgeRecord(idStructure As String, ByRef record As BridgeTaggedEntityRecord) As Boolean
        Dim canonicalId As String = BridgeDrawingGroupManager.GetGroupName(idStructure)
        If canonicalId.Length = 0 Then Return False
        EnsureBridgeIndex()
        Return _bridgeIndex.TryGetValue(canonicalId, record)
    End Function

    Private Sub EnsureBridgeIndex()
        If Not _bridgeIndexDirty AndAlso _bridgeIndex IsNot Nothing Then Return

        _bridgeIndex = New Dictionary(Of String, BridgeTaggedEntityRecord)(StringComparer.Ordinal)
        _entityBridgeIds = New Dictionary(Of DwgEntity, String)()
        _bridgeIndexDirty = False
        If Drawing Is Nothing Then Return

        _bridgeIndex = BridgeTaggedEntityIndex.Build(
            Drawing.Blocks.SelectMany(Function(block As DwgBlock) block.Entities))
        For Each item As KeyValuePair(Of String, BridgeTaggedEntityRecord) In _bridgeIndex
            For Each entity As DwgEntity In item.Value.Entities
                _entityBridgeIds(entity) = item.Key
            Next
        Next
    End Sub

    Private Shared Function GetEntity(item As Object) As DwgEntity
        Dim entity As DwgEntity = TryCast(item, DwgEntity)
        If entity IsNot Nothing Then Return entity

        Dim subEntity As DwgSubEntity = TryCast(item, DwgSubEntity)
        Return If(subEntity Is Nothing, Nothing, subEntity.Owner)
    End Function

    Private Sub AttachDrawing(drawing As Drawing)
        If drawing Is Nothing Then Return
        AddHandler drawing.AfterAddEntity, AddressOf OnDrawingChanged
        AddHandler drawing.ModifyEntity, AddressOf OnDrawingChanged
        AddHandler drawing.BeforeRemoveEntity, AddressOf OnDrawingChanged
        AddHandler drawing.AfterUpdate, AddressOf OnDrawingChanged
    End Sub

    Private Sub DetachDrawing(drawing As Drawing)
        If drawing Is Nothing Then Return
        RemoveHandler drawing.AfterAddEntity, AddressOf OnDrawingChanged
        RemoveHandler drawing.ModifyEntity, AddressOf OnDrawingChanged
        RemoveHandler drawing.BeforeRemoveEntity, AddressOf OnDrawingChanged
        RemoveHandler drawing.AfterUpdate, AddressOf OnDrawingChanged
    End Sub

    Private Sub OnDrawingChanged(sender As Object, e As EventArgs)
        InvalidateBridgeIndex()
        If Object.ReferenceEquals(sender, Drawing) Then ReconcileContainerState()
    End Sub

    Private Sub ReconcileContainerState()
        EnsureBridgeIndex()
        Dim openBridgeExists As Boolean = String.IsNullOrEmpty(_containerState.OpenBridgeId)
        If Not openBridgeExists Then
            For Each record As BridgeTaggedEntityRecord In _bridgeIndex.Values
                If _containerState.IsBridgeOpen(record.IdStructure) Then
                    openBridgeExists = True
                    Exit For
                End If
            Next
        End If
        _selectionSet.Reconcile()
        If Not openBridgeExists Then _containerState.ExitBridge()
    End Sub

    Private Sub InvalidateBridgeIndex()
        _bridgeIndexDirty = True
        _bridgeIndex = Nothing
        _entityBridgeIds = Nothing
    End Sub

End Class

Friend NotInheritable Class BridgeSelectionSet
    Inherits SelectionSet

    Private ReadOnly _layer As BridgeDrawingLayer
    Private ReadOnly _inner As SelectionSet
    Private ReadOnly _selected As New HashSet(Of Object)(ReferenceObjectComparer.Instance)

    Public Sub New(layer As BridgeDrawingLayer, inner As SelectionSet)
        MyBase.New(layer)
        _layer = layer
        _inner = inner
    End Sub

    Public Overrides ReadOnly Property Count As Integer
        Get
            Dim result As Integer
            For Each item As Object In Me
                result += 1
            Next
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property SupportClipboard As Boolean
        Get
            Return _selected.Count = 0 AndAlso _inner.SupportClipboard
        End Get
    End Property

    Public Overrides ReadOnly Property SupportCopyTransform As Boolean
        Get
            Return SupportTransform AndAlso _inner.SupportCopyTransform
        End Get
    End Property

    Public Overrides ReadOnly Property SupportDragAndDrop As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property SupportTransform As Boolean
        Get
            If Not _inner.SupportTransform Then Return False
            For Each item As Object In Me
                If TypeOf item Is BridgePresentationObject Then Return False
            Next
            Return True
        End Get
    End Property

    Friend Sub Reconcile()
        Dim invalidItems As New List(Of Object)()
        For Each item As Object In _selected
            Dim presentationObject As BridgePresentationObject = TryCast(item, BridgePresentationObject)
            If presentationObject IsNot Nothing Then
                If Not _layer.IsPresentationObjectOwned(presentationObject) Then invalidItems.Add(item)
            ElseIf Not _inner.IsOwned(item) OrElse
                   Not _layer.IsBridgeItem(item) OrElse
                   Not Object.ReferenceEquals(_layer.ResolveSelectionObject(item), item) Then
                invalidItems.Add(item)
            End If
        Next

        For Each item As Object In invalidItems
            _selected.Remove(item)
            If _inner.IsOwned(item) Then SetEntitySelected(item, False)
        Next
    End Sub

    Public Overrides Sub Clear()
        For Each item As Object In _selected
            SetEntitySelected(item, False)
        Next
        _selected.Clear()
        _inner.Clear()
    End Sub

    Public Overrides Sub [Erase]()
        _inner.Erase()
        _selected.Clear()
    End Sub

    Public Overrides Iterator Function GetEnumerator() As IEnumerator
        Dim yielded As New HashSet(Of Object)(ReferenceObjectComparer.Instance)
        For Each item As Object In _selected
            If yielded.Add(item) Then Yield item
        Next
        For Each item As Object In _inner
            Dim resolved As Object = _layer.ResolveSelectionObject(item)
            If resolved IsNot Nothing AndAlso yielded.Add(resolved) Then Yield resolved
        Next
    End Function

    Public Overrides Function GetObjectsAtPoint(point As Vector3D,
                                                match As Predicate(Of Object),
                                                waitTimeOut As Integer) As IEnumerable(Of KeyValuePair(Of Double, Object))
        Return MapPointResults(
            _inner.GetObjectsAtPoint(point, AddressOf _layer.IsPlanSelectable, waitTimeOut),
            match)
    End Function

    Public Overrides Function GetObjectsAtRay(ray As Ray3D,
                                              match As Predicate(Of Object),
                                              waitTimeOut As Integer) As IEnumerable(Of KeyValuePair(Of Double, Object))
        Return MapPointResults(_inner.GetObjectsAtRay(ray, Nothing, waitTimeOut), match)
    End Function

    Public Overrides Iterator Function GetObjectsAtFrustum(frustum As BoundingFrustum,
                                                           match As Predicate(Of Object),
                                                           waitTimeOut As Integer) As IEnumerable(Of KeyValuePair(Of Vector3D, Object))
        Dim yielded As New HashSet(Of Object)(ReferenceObjectComparer.Instance)
        For Each hit As KeyValuePair(Of Vector3D, Object) In _inner.GetObjectsAtFrustum(frustum, Nothing, waitTimeOut)
            Dim resolved As Object = _layer.ResolveSelectionObject(hit.Value)
            If resolved IsNot Nothing AndAlso
               (match Is Nothing OrElse match(resolved)) AndAlso
               yielded.Add(resolved) Then
                Yield New KeyValuePair(Of Vector3D, Object)(hit.Key, resolved)
            End If
        Next
    End Function

    Public Overrides Sub GetObjectsByFrame(mode As FrameSelectType,
                                           rect As RectangleD,
                                           match As Predicate(Of Object),
                                           action As Action(Of Object))
        Dim hits As New List(Of Object)()
        _inner.GetObjectsByFrame(
            mode,
            rect,
            AddressOf _layer.IsPlanSelectable,
            Sub(item As Object) hits.Add(item))
        EmitFrameResults(mode, hits, match, action)
    End Sub

    Public Overrides Sub GetObjectsByPolygon(mode As FrameSelectType,
                                             pointsList As List(Of Vector2D),
                                             match As Predicate(Of Object),
                                             action As Action(Of Object))
        Dim hits As New List(Of Object)()
        _inner.GetObjectsByPolygon(
            mode,
            pointsList,
            AddressOf _layer.IsPlanSelectable,
            Sub(item As Object) hits.Add(item))
        EmitFrameResults(mode, hits, match, action)
    End Sub

    Public Overrides Iterator Function GetSelectable() As IEnumerable
        Dim yielded As New HashSet(Of Object)(ReferenceObjectComparer.Instance)
        For Each item As Object In _inner.GetSelectable()
            If Not _layer.IsPlanSelectable(item) Then Continue For
            Dim resolved As Object = _layer.ResolveSelectionObject(item)
            If resolved IsNot Nothing AndAlso yielded.Add(resolved) Then Yield resolved
        Next
    End Function

    Public Overrides Function IsOwned(obj As Object) As Boolean
        Dim presentationObject As BridgePresentationObject = TryCast(obj, BridgePresentationObject)
        If presentationObject IsNot Nothing Then Return _layer.IsPresentationObjectOwned(presentationObject)
        If _layer.IsBridgeItem(obj) Then Return _inner.IsOwned(obj)
        Return _inner.IsOwned(obj)
    End Function

    Public Overrides Function IsSelected(obj As Object) As Boolean
        Dim resolved As Object = _layer.ResolveSelectionObject(obj)
        If resolved IsNot Nothing AndAlso _selected.Contains(resolved) Then Return True
        Return _inner.IsSelected(obj)
    End Function

    Public Overrides Function IsEnable(obj As Object) As Boolean
        Dim presentationObject As BridgePresentationObject = TryCast(obj, BridgePresentationObject)
        If presentationObject IsNot Nothing Then Return _layer.Enable AndAlso _layer.IsPresentationObjectOwned(presentationObject)
        Return _inner.IsEnable(obj)
    End Function

    Public Overrides Sub [Select](item As Object, bFlag As Boolean)
        Dim resolved As Object = _layer.ResolveSelectionObject(item)
        If TypeOf resolved Is BridgePresentationObject OrElse _layer.IsBridgeItem(item) Then
            If bFlag Then
                If _selected.Add(resolved) Then SetEntitySelected(resolved, True)
            ElseIf _selected.Remove(resolved) Then
                SetEntitySelected(resolved, False)
            End If
            Return
        End If

        _inner.Select(item, bFlag)
    End Sub

    Public Overrides Sub [Select](items As IEnumerable, bFlag As Boolean)
        BeginSelect()
        Try
            If items Is Nothing Then
                Clear()
                Return
            End If
            For Each item As Object In items
                [Select](item, bFlag)
            Next
        Finally
            EndSelect()
        End Try
    End Sub

    Public Overrides Iterator Function GetObjectGrips(obj As Object) As IEnumerable(Of IGrip)
        If TypeOf obj Is BridgePresentationObject Then Return
        For Each grip As IGrip In _inner.GetObjectGrips(obj)
            Yield grip
        Next
    End Function

    Public Overrides Function GetTransformData() As Object
        If Not SupportTransform Then Throw New NotSupportedException()
        If _selected.Count = 0 Then Return _inner.GetTransformData()

        Dim entities As New List(Of DwgEntity)()
        Dim yielded As New HashSet(Of DwgEntity)()
        For Each item As Object In Me
            Dim entity As DwgEntity = SelectedEntity(item)
            If entity IsNot Nothing AndAlso yielded.Add(entity) Then entities.Add(entity)
        Next
        Return New BridgeTransformData(entities)
    End Function

    Public Overrides Sub PaintTransformData(data As Object, pen As CadPen)
        Dim bridgeData As BridgeTransformData = TryCast(data, BridgeTransformData)
        If bridgeData Is Nothing Then
            _inner.PaintTransformData(data, pen)
            Return
        End If

        Dim args As New PaintEntityEventArgs(True, pen)
        If CadView IsNot Nothing Then
            args.AnnotationScale = CadView.AnnotationScale
            args.AnnotationRotation = CadView.ScreenRotation
        End If
        PaintEntityEventArgs.PaintEntities(Nothing, bridgeData.Entities, Vector3D.Empty, Vector3D.One, 0.0, args)
    End Sub

    Public Overrides Sub Transform(data As Object, transformMatrix As Matrix, copy As Boolean)
        Dim bridgeData As BridgeTransformData = TryCast(data, BridgeTransformData)
        If bridgeData Is Nothing Then
            _inner.Transform(data, transformMatrix, copy)
            Return
        End If
        If _layer.Drawing Is Nothing Then Return

        _layer.Drawing.BeginUpdate()
        Try
            For Each entity As DwgEntity In bridgeData.Entities
                If copy Then
                    Dim clone As DwgEntity = DirectCast(entity.Clone(), DwgEntity)
                    clone.Transform(transformMatrix)
                    _layer.Drawing.ActiveSpace.Add(clone)
                Else
                    entity.Transform(transformMatrix)
                End If
            Next
        Finally
            _layer.Drawing.EndUpdate()
        End Try
    End Sub

    Protected Overrides Function GetClipboardData() As IStgSerializable
        If Not SupportClipboard Then Throw New NotSupportedException()
        Return DirectCast(InvokeInnerClipboardMethod("GetClipboardData", Type.EmptyTypes, Nothing), IStgSerializable)
    End Function

    Protected Overrides Sub PaintClipboardData(data As IStgSerializable, pen As CadPen)
        InvokeInnerClipboardMethod(
            "PaintClipboardData",
            {GetType(IStgSerializable), GetType(CadPen)},
            {data, pen})
    End Sub

    Protected Overrides Sub AddClipboardData(data As IStgSerializable,
                                              transform As Matrix,
                                              basePoint As Vector3D,
                                              [select] As Boolean)
        InvokeInnerClipboardMethod(
            "AddClipboardData",
            {GetType(IStgSerializable), GetType(Matrix), GetType(Vector3D), GetType(Boolean)},
            {data, transform, basePoint, [select]})
    End Sub

    Private Function InvokeInnerClipboardMethod(name As String,
                                                parameterTypes As Type(),
                                                arguments As Object()) As Object
        Dim method As MethodInfo = _inner.GetType().GetMethod(
            name,
            BindingFlags.Instance Or BindingFlags.NonPublic,
            Nothing,
            parameterTypes,
            Nothing)
        If method Is Nothing Then Throw New MissingMethodException(_inner.GetType().FullName, name)
        Return method.Invoke(_inner, arguments)
    End Function

    Private Function MapPointResults(hits As IEnumerable(Of KeyValuePair(Of Double, Object)),
                                     match As Predicate(Of Object)) As IEnumerable(Of KeyValuePair(Of Double, Object))
        Dim distances As New Dictionary(Of Object, Double)(ReferenceObjectComparer.Instance)
        For Each hit As KeyValuePair(Of Double, Object) In hits
            Dim resolved As Object = _layer.ResolveSelectionObject(hit.Value)
            If resolved Is Nothing OrElse (match IsNot Nothing AndAlso Not match(resolved)) Then Continue For

            Dim currentDistance As Double
            If Not distances.TryGetValue(resolved, currentDistance) OrElse hit.Key < currentDistance Then
                distances(resolved) = hit.Key
            End If
        Next

        Return distances.
            Select(Function(item) New KeyValuePair(Of Double, Object)(item.Value, item.Key)).
            OrderBy(Function(item) item.Key).
            ToArray()
    End Function

    Private Sub EmitFrameResults(mode As FrameSelectType,
                                 hits As IEnumerable(Of Object),
                                 match As Predicate(Of Object),
                                 action As Action(Of Object))
        If action Is Nothing Then Return

        Dim rawHits As New HashSet(Of Object)(ReferenceObjectComparer.Instance)
        Dim mapped As New List(Of Object)()
        For Each hit As Object In hits
            rawHits.Add(hit)
            Dim subEntity As DwgSubEntity = TryCast(hit, DwgSubEntity)
            If subEntity IsNot Nothing Then rawHits.Add(subEntity.Owner)
            mapped.Add(_layer.ResolveSelectionObject(hit))
        Next

        Dim yielded As New HashSet(Of Object)(ReferenceObjectComparer.Instance)
        For Each item As Object In mapped
            If item Is Nothing OrElse Not yielded.Add(item) Then Continue For
            Dim presentationObject As BridgePresentationObject = TryCast(item, BridgePresentationObject)
            If mode = FrameSelectType.Contains AndAlso
               presentationObject IsNot Nothing AndAlso
               Not _layer.IsWholeBridgeInFrame(presentationObject, rawHits) Then
                Continue For
            End If
            If match Is Nothing OrElse match(item) Then action(item)
        Next
    End Sub

    Private Shared Sub SetEntitySelected(item As Object, selected As Boolean)
        Dim entity As DwgEntity = TryCast(item, DwgEntity)
        If entity IsNot Nothing Then
            entity.IsSelected = selected
            Return
        End If

        Dim subEntity As DwgSubEntity = TryCast(item, DwgSubEntity)
        If subEntity IsNot Nothing AndAlso subEntity.Owner IsNot Nothing Then
            subEntity.Owner.IsSelected = selected
        End If
    End Sub

    Private Shared Function SelectedEntity(item As Object) As DwgEntity
        Dim entity As DwgEntity = TryCast(item, DwgEntity)
        If entity IsNot Nothing Then Return entity

        Dim subEntity As DwgSubEntity = TryCast(item, DwgSubEntity)
        Return If(subEntity Is Nothing, Nothing, subEntity.Owner)
    End Function

    Private NotInheritable Class BridgeTransformData
        Public Sub New(entities As List(Of DwgEntity))
            Me.Entities = entities
        End Sub

        Public ReadOnly Entities As List(Of DwgEntity)
    End Class
End Class

Friend NotInheritable Class ReferenceObjectComparer
    Implements IEqualityComparer(Of Object)

    Public Shared ReadOnly Instance As New ReferenceObjectComparer()

    Private Sub New()
    End Sub

    Public Overloads Function Equals(x As Object, y As Object) As Boolean Implements IEqualityComparer(Of Object).Equals
        Return Object.ReferenceEquals(x, y)
    End Function

    Public Overloads Function GetHashCode(obj As Object) As Integer Implements IEqualityComparer(Of Object).GetHashCode
        Return If(obj Is Nothing, 0, RuntimeHelpers.GetHashCode(obj))
    End Function
End Class
