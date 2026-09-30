Imports Newtonsoft.Json.Linq
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization

Friend NotInheritable Class BridgeAppearanceRebuildScope
    Implements IDisposable

    Private Shared ReadOnly SyncRoot As New Object()
    Private Shared ReadOnly ActiveStates As New Dictionary(Of Drawing, Dictionary(Of String, ScopeState))()

    Private ReadOnly _state As ScopeState
    Private _disposed As Boolean

    Private Sub New(state As ScopeState)
        _state = state
    End Sub

    Public Shared Function Begin(entity As DwgEntity) As BridgeAppearanceRebuildScope
        If entity Is Nothing Then Return New BridgeAppearanceRebuildScope(Nothing)
        Dim data As New StructureElement()
        If Not FuncXRecords.getXRecords(entity, data) Then Return New BridgeAppearanceRebuildScope(Nothing)
        Return Begin(entity.Drawing, data.IdStructure)
    End Function

    Public Shared Function Begin(drawing As Drawing, bridgeId As String) As BridgeAppearanceRebuildScope
        If drawing Is Nothing OrElse String.IsNullOrWhiteSpace(bridgeId) Then
            Return New BridgeAppearanceRebuildScope(Nothing)
        End If
        Dim bridgeKey As String = BridgeDrawingGroupManager.GetGroupName(bridgeId)
        If String.IsNullOrEmpty(bridgeKey) Then Return New BridgeAppearanceRebuildScope(Nothing)

        SyncLock SyncRoot
            Dim drawingStates As Dictionary(Of String, ScopeState) = Nothing
            ActiveStates.TryGetValue(drawing, drawingStates)

            Dim state As ScopeState = Nothing
            If drawingStates IsNot Nothing AndAlso drawingStates.TryGetValue(bridgeKey, state) Then
                state.Depth += 1
                Return New BridgeAppearanceRebuildScope(state)
            End If

            Dim updateStarted As Boolean
            Dim registered As Boolean
            Try
                drawing.BeginUpdate("Перестроение и оформление моста")
                updateStarted = True
                state = New ScopeState With {
                    .Drawing = drawing,
                    .BridgeId = bridgeKey,
                    .Depth = 1,
                    .Snapshots = Capture(drawing, bridgeKey)
                }
                If drawingStates Is Nothing Then
                    drawingStates = New Dictionary(Of String, ScopeState)(StringComparer.Ordinal)
                    ActiveStates.Add(drawing, drawingStates)
                End If
                drawingStates.Add(bridgeKey, state)
                registered = True
                Return New BridgeAppearanceRebuildScope(state)
            Catch primaryFailure As Exception
                Dim cleanupFailure As Exception = Nothing
                If updateStarted Then
                    Try
                        drawing.EndUpdate()
                    Catch ex As Exception
                        cleanupFailure = ex
                    End Try
                End If
                If cleanupFailure IsNot Nothing Then
                    Throw New AggregateException(
                        "Не удалось подготовить сохранение оформления и завершить изменение чертежа.",
                        primaryFailure,
                        cleanupFailure)
                End If
                Throw
            Finally
                If Not registered AndAlso drawingStates IsNot Nothing AndAlso drawingStates.Count = 0 Then
                    ActiveStates.Remove(drawing)
                End If
            End Try
        End SyncLock
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True
        If _state Is Nothing Then Return

        Dim shouldRestore As Boolean
        SyncLock SyncRoot
            _state.Depth -= 1
            shouldRestore = _state.Depth = 0
        End SyncLock
        If Not shouldRestore Then Return

        Dim failure As Exception = Nothing
        Try
            Try
                Restore(_state)
            Catch ex As Exception
                failure = ex
            End Try
            Try
                _state.Drawing.EndUpdate()
            Catch ex As Exception
                If failure Is Nothing Then
                    failure = ex
                Else
                    failure = New AggregateException(
                        "Не удалось восстановить оформление и завершить изменение чертежа.",
                        failure,
                        ex)
                End If
            End Try
        Finally
            SyncLock SyncRoot
                Dim drawingStates As Dictionary(Of String, ScopeState) = Nothing
                If ActiveStates.TryGetValue(_state.Drawing, drawingStates) Then
                    drawingStates.Remove(_state.BridgeId)
                    If drawingStates.Count = 0 Then ActiveStates.Remove(_state.Drawing)
                End If
            End SyncLock
        End Try
        If failure IsNot Nothing Then ReportRestoreFailure(failure)
    End Sub

    Private Shared Function Capture(drawing As Drawing, bridgeKey As String) As List(Of AppearanceSnapshot)
        Dim result As New List(Of AppearanceSnapshot)()
        For Each target As BridgeAppearanceNativeTarget In BridgeAppearanceNativeRuntime.Collect(drawing)
            If Not String.Equals(CanonicalBridgeKey(target.Info.BridgeId), bridgeKey, StringComparison.Ordinal) Then Continue For
            result.Add(New AppearanceSnapshot With {
                .ExactKey = ExactKey(target.Data),
                .SemanticKey = SemanticKey(target.Data),
                .Values = target.Info.Values,
                .BodyColors = BridgeAppearanceNativeRuntime.CaptureBodyColors(target.Entity)
            })
        Next
        Return result
    End Function

    Private Shared Sub Restore(state As ScopeState)
        Dim current = BridgeAppearanceNativeRuntime.Collect(state.Drawing).
            Where(Function(target) String.Equals(CanonicalBridgeKey(target.Info.BridgeId), state.BridgeId, StringComparison.Ordinal)).
            ToList()
        If current.Count = 0 OrElse state.Snapshots.Count = 0 Then Return

        Dim patches As New Dictionary(Of String, BridgeAppearancePatch)(StringComparer.Ordinal)
        Dim bodyRestores As New Dictionary(Of String, ImElement)(StringComparer.Ordinal)
        Dim restoreFailures As New List(Of Exception)()
        Dim remaining As New HashSet(Of BridgeAppearanceNativeTarget)(current)
        Dim oldExact = state.Snapshots.Where(Function(item) Not String.IsNullOrWhiteSpace(item.ExactKey)).
            GroupBy(Function(item) item.ExactKey, StringComparer.Ordinal).
            Where(Function(group) group.Count() = 1).
            ToDictionary(Function(group) group.Key, Function(group) group.Single(), StringComparer.Ordinal)
        Dim newExact = current.Select(Function(item) New With {.Target = item, .Key = ExactKey(item.Data)}).
            Where(Function(item) Not String.IsNullOrWhiteSpace(item.Key)).
            GroupBy(Function(item) item.Key, StringComparer.Ordinal).
            Where(Function(group) group.Count() = 1).
            ToDictionary(Function(group) group.Key, Function(group) group.Single().Target, StringComparer.Ordinal)

        For Each pair In oldExact
            Dim target As BridgeAppearanceNativeTarget = Nothing
            If newExact.TryGetValue(pair.Key, target) Then
                Dim patch As BridgeAppearancePatch = ToPatch(pair.Value.Values, target)
                If Not patch.IsEmpty Then patches(target.Info.SelectionId) = patch
                Try
                    PrepareBodyRestore(pair.Value, target, bodyRestores)
                Catch ex As Exception
                    restoreFailures.Add(ex)
                End Try
                remaining.Remove(target)
            End If
        Next

        Dim usedSnapshots As New HashSet(Of AppearanceSnapshot)(oldExact.Values.Where(
            Function(snapshot) newExact.ContainsKey(snapshot.ExactKey)))
        Dim oldSemantic = state.Snapshots.Where(
            Function(item) Not usedSnapshots.Contains(item) AndAlso Not String.IsNullOrEmpty(item.SemanticKey)).
            GroupBy(Function(item) item.SemanticKey, StringComparer.Ordinal).
            Where(Function(group) group.Count() = 1).
            ToDictionary(Function(group) group.Key, Function(group) group.Single(), StringComparer.Ordinal)
        Dim newSemantic = remaining.Select(Function(target) New With {.Target = target, .Key = SemanticKey(target.Data)}).
            Where(Function(item) Not String.IsNullOrEmpty(item.Key)).
            GroupBy(Function(item) item.Key, StringComparer.Ordinal).
            Where(Function(group) group.Count() = 1).
            ToDictionary(Function(group) group.Key, Function(group) group.Single().Target, StringComparer.Ordinal)
        For Each pair In oldSemantic
            Dim target As BridgeAppearanceNativeTarget = Nothing
            If newSemantic.TryGetValue(pair.Key, target) Then
                Dim patch As BridgeAppearancePatch = ToPatch(pair.Value.Values, target)
                If Not patch.IsEmpty Then patches(target.Info.SelectionId) = patch
                Try
                    PrepareBodyRestore(pair.Value, target, bodyRestores)
                Catch ex As Exception
                    restoreFailures.Add(ex)
                End Try
            End If
        Next

        If patches.Count > 0 Then
            BridgeAppearanceNativeRuntime.Apply(
                state.Drawing,
                current.Where(Function(target) patches.ContainsKey(target.Info.SelectionId)),
                patches)
        End If
        For Each target As BridgeAppearanceNativeTarget In current
            Dim clone As ImElement = Nothing
            If bodyRestores.TryGetValue(target.Info.SelectionId, clone) Then
                Try
                    BridgeAppearanceNativeRuntime.AssignBodyColorClone(target.Entity, clone)
                Catch ex As Exception
                    restoreFailures.Add(ex)
                End Try
            End If
        Next
        If restoreFailures.Count > 0 Then
            Throw New AggregateException("Не удалось восстановить цвет части 3D-элементов.", restoreFailures)
        End If
    End Sub

    Private Shared Function ToPatch(values As BridgeAppearanceValues,
                                    target As BridgeAppearanceNativeTarget) As BridgeAppearancePatch
        Dim current As BridgeAppearanceValues = target.Info.Values
        Dim patch As New BridgeAppearancePatch()
        If values.CadColorValue <> current.CadColorValue Then patch.CadColorValue = values.CadColorValue
        If target.Data.Name <> StructureElement.typeObject.counterTopBeam AndAlso
           target.Data.Name <> StructureElement.typeObject.counterBottomBeam AndAlso
           target.Data.Name <> StructureElement.typeObject.counterSubFermentersTop AndAlso
           target.Data.Name <> StructureElement.typeObject.counterSubFermentersBottom AndAlso
           target.Data.Name <> StructureElement.typeObject.counterSubFermentersUTop AndAlso
           target.Data.Name <> StructureElement.typeObject.counterSubFermentersUBottom AndAlso
           Not String.Equals(values.LayerName, current.LayerName, StringComparison.Ordinal) Then
            patch.LayerName = values.LayerName
        End If
        If Not String.Equals(values.LinetypeName, current.LinetypeName, StringComparison.Ordinal) Then
            patch.LinetypeName = values.LinetypeName
        End If
        If values.LinetypeScale <> current.LinetypeScale Then patch.LinetypeScale = values.LinetypeScale
        If values.Lineweight <> current.Lineweight Then patch.Lineweight = values.Lineweight
        If target.Info.SupportsWidth AndAlso values.Width.HasValue AndAlso Not values.Width.Equals(current.Width) Then
            patch.Width = values.Width
        End If
        If target.Info.SupportsFill Then
            If values.FillCadColorValue.HasValue AndAlso Not values.FillCadColorValue.Equals(current.FillCadColorValue) Then
                patch.FillCadColorValue = values.FillCadColorValue
            End If
            If values.HatchPatternName IsNot Nothing AndAlso
               Not String.Equals(values.HatchPatternName, current.HatchPatternName, StringComparison.OrdinalIgnoreCase) Then
                patch.HatchPatternName = values.HatchPatternName
            End If
            If values.HatchScale.HasValue AndAlso Not values.HatchScale.Equals(current.HatchScale) Then
                patch.HatchScale = values.HatchScale
            End If
            If values.HatchAngle.HasValue AndAlso Not values.HatchAngle.Equals(current.HatchAngle) Then
                patch.HatchAngle = values.HatchAngle
            End If
        End If
        Return patch
    End Function

    Private Shared Sub PrepareBodyRestore(snapshot As AppearanceSnapshot,
                                          target As BridgeAppearanceNativeTarget,
                                          bodyRestores As Dictionary(Of String, ImElement))
        If snapshot.BodyColors Is Nothing OrElse snapshot.BodyColors.Length = 0 Then Return
        Dim bodyArgb As Integer = snapshot.BodyColors(0)
        If snapshot.BodyColors.Any(Function(value) value <> bodyArgb) Then
            Throw New InvalidOperationException(
                "Различное оформление частей 3D-элемента пропущено: части перестроенной модели нельзя сопоставить однозначно.")
        End If
        Dim currentColors As Integer() = BridgeAppearanceNativeRuntime.CaptureBodyColors(target.Entity)
        If currentColors IsNot Nothing AndAlso
           currentColors.Length > 0 AndAlso
           currentColors.All(Function(value) value = bodyArgb) Then
            Return
        End If
        bodyRestores(target.Info.SelectionId) =
            BridgeAppearanceNativeRuntime.CreateBodyColorClone(target.Entity, snapshot.BodyColors)
    End Sub

    Private Shared Function SemanticKey(data As StructureElement) As String
        If data Is Nothing Then Return Nothing
        Dim baseKey As String = BuildBaseKey(data)
        If String.IsNullOrEmpty(baseKey) Then Return Nothing
        Dim names As String() = SemanticFieldNames(CInt(data.Name))
        If names Is Nothing Then Return Nothing
        If names.Length = 0 Then Return baseKey & "|единственный"
        If String.IsNullOrWhiteSpace(data.KeyParameter) Then Return Nothing

        Try
            Dim json As JObject = JObject.Parse(data.KeyParameter)
            Dim values As New List(Of String)()
            For Each name As String In names
                Dim propertyValue = json.Properties().FirstOrDefault(
                    Function(item) String.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
                If propertyValue Is Nothing OrElse propertyValue.Value.Type = JTokenType.Null Then Return Nothing
                values.Add(propertyValue.Value.ToString())
            Next
            Return baseKey & "|" & String.Join("|", values)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function ExactKey(data As StructureElement) As String
        If data Is Nothing OrElse String.IsNullOrWhiteSpace(data.IdElement) Then Return Nothing
        Dim baseKey As String = BuildBaseKey(data)
        If String.IsNullOrEmpty(baseKey) Then Return Nothing
        Return baseKey & "|id|" & data.IdElement
    End Function

    Private Shared Function BuildBaseKey(data As StructureElement) As String
        Dim bridgeKey As String = CanonicalBridgeKey(data.IdStructure)
        If String.IsNullOrEmpty(bridgeKey) Then Return Nothing
        Return String.Join("|", {
            bridgeKey,
            CInt(data.ClassBridgeObject).ToString(),
            CInt(data.ClassObject).ToString(),
            CInt(data.Name).ToString()
        })
    End Function

    Private Shared Function CanonicalBridgeKey(bridgeId As String) As String
        Return BridgeDrawingGroupManager.GetGroupName(bridgeId)
    End Function

    Private Shared Function SemanticFieldNames(typeValue As Integer) As String()
        Select Case typeValue
            Case 100, 182
                Return Array.Empty(Of String)()
            Case 103, 155, 156, 163, 184
                Return {"numberProlet", "numberRow"}
            Case 102
                Return {"numberPillar", "numberSubPillar", "numberProlet"}
            Case 104
                Return {"Number"}
            Case 105
                Return {"NumberPillar", "Number"}
            Case 120, 121, 122, 123, 124, 164
                Return {"NumberPillar", "NumberSubPillar"}
            Case 106
                Return {"NumberPillar", "Number"}
            Case 145, 146, 165, 200
                Return {"NumberPillar", "NumberSubPillar"}
            Case 108
                Return {"NumberPillar", "NumberSubPillars", "Number"}
            Case 149, 150
                Return {"NumberPillar", "NumberSubPillar", "NumberRack", "TypeContour"}
            Case 170
                Return {"NumberPillar", "NumberSubPillar", "NumberRack"}
            Case 112
                Return {"NumberPillar", "NumberSubPillars", "NumberColumn", "NumberRow"}
            Case 143, 144, 169
                Return {"NumberPillar", "NumberSubPillar", "NumberColumn", "NumberRow"}
            Case 118
                Return {"numberProlet", "numberLeftBeam", "numberRightBeam"}
            Case 157, 158
                Return {"numberProlet", "numberRow", "TypeCounter"}
            Case 161, 180
                Return {"numberProlet", "numberRow"}
            Case Else
                Return Nothing
        End Select
    End Function

    Private Shared Sub ReportRestoreFailure(failure As Exception)
        Dim report As New BuildFailureReport(
            "Восстановление оформления моста",
            "Завершение перестроения",
            "Не удалось восстановить оформление части перестроенных элементов.",
            "Проверьте оформление перестроенных элементов и при необходимости повторите его в редакторе.",
            "BridgeAppearanceRebuildScope.Dispose",
            failure)
        BuildFailurePresenter.Show(report)
    End Sub

    Private NotInheritable Class ScopeState
        Public Drawing As Drawing
        Public BridgeId As String
        Public Depth As Integer
        Public Snapshots As List(Of AppearanceSnapshot)
    End Class

    Private NotInheritable Class AppearanceSnapshot
        Public ExactKey As String
        Public SemanticKey As String
        Public Values As BridgeAppearanceValues
        Public BodyColors As Integer()
    End Class
End Class
