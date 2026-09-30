Imports System.ComponentModel

Public NotInheritable Class BridgeContainerState
    Private ReadOnly _presentationObjects As New Dictionary(Of String, BridgePresentationObject)(StringComparer.Ordinal)
    Private _openBridgeId As String

    Public ReadOnly Property OpenBridgeId As String
        Get
            Return _openBridgeId
        End Get
    End Property

    Public Sub EnterBridge(idStructure As String)
        Dim canonicalId As String = NormalizeId(idStructure)
        If canonicalId.Length = 0 Then Return
        _openBridgeId = canonicalId
    End Sub

    Public Sub ExitBridge()
        _openBridgeId = Nothing
    End Sub

    Public Function IsBridgeOpen(idStructure As String) As Boolean
        Dim canonicalId As String = NormalizeId(idStructure)
        Return canonicalId.Length > 0 AndAlso
               String.Equals(_openBridgeId, canonicalId, StringComparison.Ordinal)
    End Function

    Public Function ResolveSelectionObject(item As Object, idStructure As String) As Object
        If item Is Nothing Then Return Nothing

        Dim canonicalId As String = NormalizeId(idStructure)
        If canonicalId.Length = 0 OrElse IsBridgeOpen(idStructure) Then Return item

        Dim presentationObject As BridgePresentationObject = Nothing
        If Not _presentationObjects.TryGetValue(canonicalId, presentationObject) Then
            presentationObject = New BridgePresentationObject(idStructure.Trim())
            _presentationObjects.Add(canonicalId, presentationObject)
        End If
        Return presentationObject
    End Function

    Friend Function Owns(presentationObject As BridgePresentationObject) As Boolean
        If presentationObject Is Nothing Then Return False

        Dim cachedObject As BridgePresentationObject = Nothing
        Return _presentationObjects.TryGetValue(NormalizeId(presentationObject.IdStructure), cachedObject) AndAlso
               Object.ReferenceEquals(cachedObject, presentationObject)
    End Function

    Friend Sub Reset()
        _openBridgeId = Nothing
        _presentationObjects.Clear()
    End Sub

    Private Shared Function NormalizeId(idStructure As String) As String
        Return BridgeDrawingGroupManager.GetGroupName(idStructure)
    End Function
End Class

Public NotInheritable Class BridgePresentationObject
    Private ReadOnly _idStructure As String
    Private _nameBridge As String
    Private _proletCount As Integer
    Private _leftStructureWidth As Double
    Private _rightStructureWidth As Double
    Private _startPlacementPosition As Double
    Private _alignmentName As String

    Friend Sub New(idStructure As String)
        _idStructure = If(idStructure, String.Empty)
    End Sub

    <Browsable(False)>
    Public ReadOnly Property IdStructure As String
        Get
            Return _idStructure
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Наименование сооружения"),
     Description("Наименование мостового сооружения")>
    Public ReadOnly Property NameBridge As String
        Get
            Return _nameBridge
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Число пролётов"),
     Description("Число пролётов мостового сооружения")>
    Public ReadOnly Property ProletCount As Integer
        Get
            Return _proletCount
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Ширина сооружения слева"),
     Description("Ширина сооружения слева от оси, м")>
    Public ReadOnly Property LeftStructureWidth As Double
        Get
            Return _leftStructureWidth
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Ширина сооружения справа"),
     Description("Ширина сооружения справа от оси, м")>
    Public ReadOnly Property RightStructureWidth As Double
        Get
            Return _rightStructureWidth
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Начальный пикет"),
     Description("Начальный пикет раскладки балок, ПК+")>
    Public ReadOnly Property startPlacementPosition As Double
        Get
            Return _startPlacementPosition
        End Get
    End Property

    <Browsable(True),
     Category("Мост"),
     DisplayName("Ось трассы"),
     Description("Имя проектной трассы")>
    Public ReadOnly Property AlignmentName As String
        Get
            Return _alignmentName
        End Get
    End Property

    Friend Sub Update(bridge As Bridges)
        If bridge Is Nothing Then
            _nameBridge = String.Empty
            _proletCount = 0
            _leftStructureWidth = 0.0
            _rightStructureWidth = 0.0
            _startPlacementPosition = 0.0
            _alignmentName = String.Empty
            Return
        End If

        _nameBridge = bridge.NameBridge
        _proletCount = bridge.ProletCount
        _leftStructureWidth = bridge.LeftStructureWidth
        _rightStructureWidth = bridge.RightStructureWidth
        _startPlacementPosition = bridge.startPlacementPosition
        _alignmentName = bridge.AlignmentName
    End Sub

    Public Overrides Function ToString() As String
        Dim displayName As String = If(_nameBridge, String.Empty).Trim()
        If displayName.Length = 0 Then displayName = _idStructure
        Return "Мост «" & displayName & "»"
    End Function
End Class

Public NotInheritable Class BridgePlanVisibilityPolicy
    Private Sub New()
    End Sub

    Public Shared Function ShouldDrawInPlan(isBridgeEntity As Boolean, isModel3d As Boolean) As Boolean
        Return Not (isBridgeEntity AndAlso isModel3d)
    End Function
End Class
