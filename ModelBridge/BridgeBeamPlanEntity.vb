Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports Topomatic.Cad.Foundation
Imports Topomatic.ComponentModel
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Design
Imports Topomatic.Dwg.Entities
Imports Topomatic.Stg

<DesignAlias("INFRA_BRIDGE_BEAM_PLAN")>
<EntityController(GetType(BridgeBeamPlanEntityController))>
Public NotInheritable Class BridgeBeamPlanEntity
    Inherits DwgEntity

    Public Const EntityAlias As String = "INFRA_BRIDGE_BEAM_PLAN"
    Public Const DefaultPlanLayerName As String = "ИССО-П_Балка (контур)"
    Public Const DefaultHatchPatternName As String = "SOLID"
    Public Const DefaultFillColorIndex As Integer = 9
    Public Const DefaultPerimeterColorIndex As Integer = 8

    Private _outline As Vector2D() = Array.Empty(Of Vector2D)()
    Private _fillColor As CadColor = New CadColor(DefaultFillColorIndex)
    Private _hatchPatternName As String = DefaultHatchPatternName
    Private _hatchScale As Double = 1.0R
    Private _hatchAngle As Double

    Public Shared Sub RegisterActivator()
        Drawing.RegisterActivator(EntityAlias, Function() New BridgeBeamPlanEntity())
    End Sub

    <Browsable(False)>
    Public ReadOnly Property Outline As Vector2D()
        Get
            Return DirectCast(_outline.Clone(), Vector2D())
        End Get
    End Property

    <Category("Оформление")>
    <DisplayName("Цвет заливки")>
    Public Property FillColor As CadColor
        Get
            Return _fillColor
        End Get
        Set(value As CadColor)
            If _fillColor = value Then Return
            BeginUpdate()
            Try
                _fillColor = value
                Regen(EventArgs.Empty)
            Finally
                EndUpdate()
            End Try
        End Set
    End Property

    <Category("Оформление")>
    <DisplayName("Штриховка")>
    Public Property HatchPatternName As String
        Get
            Return _hatchPatternName
        End Get
        Set(value As String)
            If String.IsNullOrWhiteSpace(value) Then Throw New ArgumentException("Имя штриховки не задано.", NameOf(value))
            Dim normalized As String = value.Trim().ToUpperInvariant()
            If normalized <> "SOLID" AndAlso normalized <> "ANSI31" Then
                Throw New ArgumentException("Поддерживаются штриховки SOLID и ANSI31.", NameOf(value))
            End If
            If String.Equals(_hatchPatternName, normalized, StringComparison.Ordinal) Then Return
            BeginUpdate()
            Try
                _hatchPatternName = normalized
                Regen(EventArgs.Empty)
            Finally
                EndUpdate()
            End Try
        End Set
    End Property

    <Category("Оформление")>
    <DisplayName("Масштаб штриховки")>
    Public Property HatchScale As Double
        Get
            Return _hatchScale
        End Get
        Set(value As Double)
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) OrElse value <= 0.0R Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "Масштаб штриховки должен быть положительным конечным числом.")
            End If
            If _hatchScale = value Then Return
            BeginUpdate()
            Try
                _hatchScale = value
                Regen(EventArgs.Empty)
            Finally
                EndUpdate()
            End Try
        End Set
    End Property

    <Category("Оформление")>
    <DisplayName("Угол штриховки")>
    Public Property HatchAngle As Double
        Get
            Return _hatchAngle
        End Get
        Set(value As Double)
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "Угол штриховки должен быть конечным числом.")
            End If
            If _hatchAngle = value Then Return
            BeginUpdate()
            Try
                _hatchAngle = value
                Regen(EventArgs.Empty)
            Finally
                EndUpdate()
            End Try
        End Set
    End Property

    Public Overrides ReadOnly Property EntityName As String
        Get
            Return "Контур балки моста"
        End Get
    End Property

    Public Overrides ReadOnly Property IsBackgroud As Boolean
        Get
            Return True
        End Get
    End Property

    Public Sub SetOutline(points As IEnumerable(Of Vector2D))
        If points Is Nothing Then Throw New ArgumentNullException(NameOf(points))
        Dim candidate As Vector2D() = points.ToArray()
        If candidate.Length < 4 OrElse Not SamePoint(candidate(0), candidate(candidate.Length - 1)) Then
            Throw New ArgumentException("Контур балки должен быть замкнут и содержать не менее трёх вершин.", NameOf(points))
        End If
        For Each pointValue As Vector2D In candidate
            If Double.IsNaN(pointValue.X) OrElse Double.IsInfinity(pointValue.X) OrElse
               Double.IsNaN(pointValue.Y) OrElse Double.IsInfinity(pointValue.Y) Then
                Throw New ArgumentException("Координаты контура балки должны быть конечными.", NameOf(points))
            End If
        Next
        BeginUpdate()
        Try
            _outline = DirectCast(candidate.Clone(), Vector2D())
            Regen(EventArgs.Empty)
        Finally
            EndUpdate()
        End Try
    End Sub

    Public Overrides Sub Layout(entities As IList(Of DwgEntity), e As LayoutEntityEventArgs)
        If entities Is Nothing Then Throw New ArgumentNullException(NameOf(entities))
        If _outline.Length < 4 Then Return

        Dim hatch As DwgHatch = CreateHatch()
        Dim perimeter As DwgPolyline = CreatePerimeter()
        entities.Add(hatch)
        entities.Add(perimeter)
    End Sub

    Public Overrides Function IntersectWith(boundsValue As BoundingBox2D,
                                            annotationScale As Double) As Boolean
        If _outline.Length < 4 OrElse Not BoundsOverlap(Bounds, boundsValue) Then Return False

        For index As Integer = 0 To _outline.Length - 2
            If PointInBounds(_outline(index), boundsValue) Then Return True
            If SegmentIntersectsBounds(_outline(index), _outline(index + 1), boundsValue) Then Return True
        Next
        For Each corner As Vector2D In boundsValue.GetCorners()
            If ContainsPoint(corner) Then Return True
        Next
        Return False
    End Function

    Friend Function CreateLayoutEntities(highlight As Boolean) As List(Of DwgEntity)
        Dim result As New List(Of DwgEntity)()
        Layout(result, New LayoutEntityEventArgs(1.0R, 0.0R))
        If highlight Then
            For Each item As DwgEntity In result
                If TypeOf item Is DwgHatch Then
                    item.Color = New CadColor(2)
                    item.Transparency = New Transparency(100)
                Else
                    item.Color = New CadColor(1)
                    item.Lineweight = CType(100, Lineweight)
                End If
            Next
        End If
        Return result
    End Function

    Friend Function ContainsPoint(pointValue As Vector2D) As Boolean
        If _outline.Length < 4 Then Return False
        Dim inside As Boolean = False
        For index As Integer = 0 To _outline.Length - 2
            Dim first As Vector2D = _outline(index)
            Dim second As Vector2D = _outline(index + 1)
            If PointOnSegment(pointValue, first, second) Then Return True
            If (first.Y > pointValue.Y) <> (second.Y > pointValue.Y) Then
                Dim crossingX As Double = first.X +
                    (pointValue.Y - first.Y) * (second.X - first.X) / (second.Y - first.Y)
                If crossingX >= pointValue.X Then inside = Not inside
            End If
        Next
        Return inside
    End Function

    Protected Overrides Sub OnRegen(e As EventArgs)
        If _outline.Length < 4 Then
            Bounds = BoundingBox2D.Empty
            Return
        End If
        Dim minimum As New Vector2D(_outline.Min(Function(pointValue) pointValue.X),
                                    _outline.Min(Function(pointValue) pointValue.Y))
        Dim maximum As New Vector2D(_outline.Max(Function(pointValue) pointValue.X),
                                    _outline.Max(Function(pointValue) pointValue.Y))
        Bounds = New BoundingBox2D(minimum, maximum)
    End Sub

    Protected Overrides Sub OnAssign(entity As DwgEntity)
        Dim source As BridgeBeamPlanEntity = TryCast(entity, BridgeBeamPlanEntity)
        If source Is Nothing Then Throw New ArgumentException("Нельзя назначить данные объекта другого типа.", NameOf(entity))
        _outline = DirectCast(source._outline.Clone(), Vector2D())
        _fillColor = source._fillColor
        _hatchPatternName = source._hatchPatternName
        _hatchScale = source._hatchScale
        _hatchAngle = source._hatchAngle
        OnRegen(EventArgs.Empty)
    End Sub

    Protected Overrides Sub OnTransform(matrixValue As Matrix)
        For index As Integer = 0 To _outline.Length - 1
            _outline(index) = Vector2D.Transform(_outline(index), matrixValue)
        Next
        OnRegen(EventArgs.Empty)
    End Sub

    Protected Overrides Sub OnSaveToStg(node As StgNode)
        Dim pointsArray = node.AddArray("Outline", StgType.Node)
        For Each pointValue As Vector2D In _outline
            pointValue.SaveToStg(pointsArray.AddNode())
        Next
        node.AddInt32("FillColor", _fillColor.ToCompressValue())
        node.AddString("HatchPatternName", _hatchPatternName)
        node.AddDouble("HatchScale", _hatchScale)
        node.AddDouble("HatchAngle", _hatchAngle)
    End Sub

    Protected Overrides Sub OnLoadFromStg(node As StgNode)
        Dim points As New List(Of Vector2D)()
        If node.IsExists("Outline") Then
            Dim pointsArray = node.GetArray("Outline", StgType.Node)
            For index As Integer = 0 To pointsArray.Count - 1
                points.Add(Vector2D.LoadFromStg(pointsArray.GetNode(index)))
            Next
        End If
        _outline = points.ToArray()
        _fillColor = CadColor.FromCompressValue(node.GetInt32("FillColor", New CadColor(DefaultFillColorIndex).ToCompressValue()))
        _hatchPatternName = node.GetString("HatchPatternName", DefaultHatchPatternName)
        _hatchScale = node.GetDouble("HatchScale", 1.0R)
        _hatchAngle = node.GetDouble("HatchAngle", 0.0R)
        If String.IsNullOrWhiteSpace(_hatchPatternName) OrElse
           (_hatchPatternName.ToUpperInvariant() <> "SOLID" AndAlso _hatchPatternName.ToUpperInvariant() <> "ANSI31") Then
            _hatchPatternName = DefaultHatchPatternName
        Else
            _hatchPatternName = _hatchPatternName.ToUpperInvariant()
        End If
        If Double.IsNaN(_hatchScale) OrElse Double.IsInfinity(_hatchScale) OrElse _hatchScale <= 0.0R Then _hatchScale = 1.0R
        If Double.IsNaN(_hatchAngle) OrElse Double.IsInfinity(_hatchAngle) Then _hatchAngle = 0.0R
        OnRegen(EventArgs.Empty)
    End Sub

    Private Function CreateHatch() As DwgHatch
        Dim hatch As New DwgHatch With {
            .PatternType = AcPatternType.PreDefined,
            .PatternName = _hatchPatternName,
            .PatternAngle = _hatchAngle,
            .PatternScale = _hatchScale,
            .PatternDouble = False,
            .HatchStyle = AcHatchStyle.Normal,
            .Color = _fillColor,
            .Layer = Layer,
            .Elevation = 0.0R
        }
        Dim boundary As New PolylineBoundaryPath(_outline.Length - 1)
        For index As Integer = 0 To _outline.Length - 2
            boundary.Add(New BugleVector2D(_outline(index)))
        Next
        boundary.IsClosed = True
        hatch.BoundaryPath.Add(boundary)
        Return hatch
    End Function

    Private Function CreatePerimeter() As DwgPolyline
        Dim perimeter As New DwgPolyline With {
            .Closed = True,
            .Color = Color,
            .Layer = Layer,
            .Linetype = Linetype,
            .LinetypeScale = LinetypeScale,
            .Lineweight = Lineweight,
            .Elevation = 0.0R
        }
        For index As Integer = 0 To _outline.Length - 2
            perimeter.Add(New BugleVector2D(_outline(index)))
        Next
        Return perimeter
    End Function

    Private Shared Function BoundsOverlap(first As BoundingBox2D, second As BoundingBox2D) As Boolean
        Return first.Left <= second.Right AndAlso first.Right >= second.Left AndAlso
               first.Bottom <= second.Top AndAlso first.Top >= second.Bottom
    End Function

    Private Shared Function PointInBounds(pointValue As Vector2D, boundsValue As BoundingBox2D) As Boolean
        Return pointValue.X >= boundsValue.Left AndAlso pointValue.X <= boundsValue.Right AndAlso
               pointValue.Y >= boundsValue.Bottom AndAlso pointValue.Y <= boundsValue.Top
    End Function

    Private Shared Function SegmentIntersectsBounds(first As Vector2D,
                                                     second As Vector2D,
                                                     boundsValue As BoundingBox2D) As Boolean
        Dim corners As Vector2D() = boundsValue.GetCorners()
        For index As Integer = 0 To corners.Length - 1
            If SegmentsIntersect(first, second, corners(index), corners((index + 1) Mod corners.Length)) Then Return True
        Next
        Return False
    End Function

    Private Shared Function SegmentsIntersect(firstStart As Vector2D,
                                              firstEnd As Vector2D,
                                              secondStart As Vector2D,
                                              secondEnd As Vector2D) As Boolean
        Dim firstSide As Double = Cross(firstStart, firstEnd, secondStart)
        Dim secondSide As Double = Cross(firstStart, firstEnd, secondEnd)
        Dim thirdSide As Double = Cross(secondStart, secondEnd, firstStart)
        Dim fourthSide As Double = Cross(secondStart, secondEnd, firstEnd)
        If firstSide = 0.0R AndAlso PointOnSegment(secondStart, firstStart, firstEnd) Then Return True
        If secondSide = 0.0R AndAlso PointOnSegment(secondEnd, firstStart, firstEnd) Then Return True
        If thirdSide = 0.0R AndAlso PointOnSegment(firstStart, secondStart, secondEnd) Then Return True
        If fourthSide = 0.0R AndAlso PointOnSegment(firstEnd, secondStart, secondEnd) Then Return True
        Return (firstSide < 0.0R) <> (secondSide < 0.0R) AndAlso
               (thirdSide < 0.0R) <> (fourthSide < 0.0R)
    End Function

    Private Shared Function PointOnSegment(pointValue As Vector2D,
                                           first As Vector2D,
                                           second As Vector2D) As Boolean
        Dim crossValue As Double = Cross(first, second, pointValue)
        Dim tolerance As Double = Math.Max(1.0R, Vector2D.Distance(first, second)) * 0.000000001R
        If Math.Abs(crossValue) > tolerance Then Return False
        Return pointValue.X >= Math.Min(first.X, second.X) - tolerance AndAlso
               pointValue.X <= Math.Max(first.X, second.X) + tolerance AndAlso
               pointValue.Y >= Math.Min(first.Y, second.Y) - tolerance AndAlso
               pointValue.Y <= Math.Max(first.Y, second.Y) + tolerance
    End Function

    Private Shared Function Cross(first As Vector2D, second As Vector2D, pointValue As Vector2D) As Double
        Return (second.X - first.X) * (pointValue.Y - first.Y) -
               (second.Y - first.Y) * (pointValue.X - first.X)
    End Function

    Private Shared Function SamePoint(first As Vector2D, second As Vector2D) As Boolean
        Return first.X = second.X AndAlso first.Y = second.Y
    End Function
End Class

Public NotInheritable Class BridgeBeamPlanEntityController
    Inherits DwgEntityController

    Public Overrides ReadOnly Property SupportPaint3d As Boolean
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub OnPaintEntity(entity As DwgEntity, e As PaintEntityEventArgs)
        Dim planEntity As BridgeBeamPlanEntity = TryCast(entity, BridgeBeamPlanEntity)
        If planEntity Is Nothing OrElse e Is Nothing Then Return
        Dim highlight As Boolean = e.Pen.DrawingMode = DrawingMode.Highlight
        Dim layoutEntities As List(Of DwgEntity) = planEntity.CreateLayoutEntities(highlight)
        Dim savedDrawingMode As DrawingMode = e.Pen.DrawingMode
        Dim savedHighlightMode As HighlightMode = e.Pen.HighlightMode
        Try
            If highlight Then
                e.Pen.DrawingMode = DrawingMode.Show
                e.Pen.HighlightMode = HighlightMode.Color
            End If
            PaintEntityEventArgs.PaintEntities(
                planEntity,
                layoutEntities,
                Vector3D.Empty,
                Vector3D.One,
                0.0R,
                e)
        Finally
            e.Pen.DrawingMode = savedDrawingMode
            e.Pen.HighlightMode = savedHighlightMode
            For Each layoutEntity As DwgEntity In layoutEntities
                layoutEntity.Dispose()
            Next
        End Try
    End Sub

End Class

Public NotInheritable Class BridgeBeamPlanException
    Inherits InvalidOperationException

    Public Sub New(spanNumber As Integer, rowNumber As Integer, innerException As Exception)
        MyBase.New("Не удалось построить плановый контур балки: пролёт " & spanNumber &
                   ", ряд " & rowNumber & ". " & innerException.Message,
                   innerException)
    End Sub
End Class
