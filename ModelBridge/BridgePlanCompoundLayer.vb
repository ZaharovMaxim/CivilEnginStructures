Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Layer
Imports Topomatic.FoundationClasses
Imports Topomatic.Rsf.Layers.Layers
Imports Topomatic.Sfc
Imports Topomatic.Sfc.Layer

Public NotInheritable Class BridgePlanCompoundLayer
    Inherits CompoundLayer
    Implements IDrawingContainer, ISurfaceContainer, ILayerActivityController

    Public Shared ReadOnly Guid As New Guid("{1E15E635-4188-463E-B54B-328846FC26D4}")

    Private ReadOnly _surfaceLayer As SurfaceLayer
    Private ReadOnly _drawingLayer As BridgeDrawingLayer
    Private ReadOnly _preliminaryLayer As PreliminaryFencesPlanLayer
    Private ReadOnly _roundedLayer As RoundedFencesPlanLayer
    Private ReadOnly _modelPathId As String
    Private _bridgeModel As InfrastradaBridgesModel

    Public Sub New(modelPathId As String)
        MyBase.New(modelPathId, Guid)
        _modelPathId = modelPathId
        _surfaceLayer = New SurfaceLayer("Поверхность")
        _drawingLayer = New BridgeDrawingLayer("Ситуация")
        _preliminaryLayer = New PreliminaryFencesPlanLayer()
        _roundedLayer = New RoundedFencesPlanLayer()
        Add(_surfaceLayer)
        Add(_drawingLayer)
        Add(_preliminaryLayer)
        Add(_roundedLayer)
    End Sub

    Friend ReadOnly Property DrawingLayer As BridgeDrawingLayer
        Get
            Return _drawingLayer
        End Get
    End Property

    Friend ReadOnly Property ModelPathId As String
        Get
            Return _modelPathId
        End Get
    End Property

    Public Property BridgeModel As InfrastradaBridgesModel
        Get
            Return _bridgeModel
        End Get
        Set(value As InfrastradaBridgesModel)
            _bridgeModel = value
            _surfaceLayer.Surface = If(value Is Nothing, Nothing, value.Surface)
            _drawingLayer.Drawing = If(value Is Nothing, Nothing, value.Drawing)
            _preliminaryLayer.Rsf = If(value Is Nothing, Nothing, value.Arrangement.Rsf)
            _roundedLayer.Rsf = If(value Is Nothing, Nothing, value.Arrangement.Rsf)
            If value IsNot Nothing Then value.Arrangement.LayerLinks.RefreshDrawingLayers()
        End Set
    End Property

    Public ReadOnly Property Drawing As Drawing Implements IDrawingContainer.Drawing
        Get
            Return If(_bridgeModel Is Nothing, Nothing, _bridgeModel.Drawing)
        End Get
    End Property

    Public ReadOnly Property Surface As Surface Implements ISurfaceContainer.Surface
        Get
            Return If(_bridgeModel Is Nothing, Nothing, _bridgeModel.Surface)
        End Get
    End Property

    Public Overrides Function GetSubLayers() As IEnumerable(Of ILayer)
        Return _drawingLayer.GetSubLayers()
    End Function

    Public Property ActiveLayer As ILayer Implements ILayerActivityController.ActiveLayer
        Get
            Return _drawingLayer.ActiveLayer
        End Get
        Set(value As ILayer)
            _drawingLayer.ActiveLayer = value
        End Set
    End Property

    Public Function RemoveLayer(layer As ILayer) As Boolean Implements ILayerActivityController.RemoveLayer
        Return _drawingLayer.RemoveLayer(layer)
    End Function
End Class
