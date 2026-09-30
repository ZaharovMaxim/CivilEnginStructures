Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Runtime.CompilerServices
Imports Newtonsoft.Json.Linq
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Stg
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.Runtime

Public NotInheritable Class BridgeAppearanceItem
    Public Sub New(selectionId As String,
                   bridgeId As String,
                   bridgeLabel As String,
                   classLabel As String,
                   typeLabel As String,
                   itemLabel As String,
                   values As BridgeAppearanceValues,
                   supportsWidth As Boolean,
                   supportsBody As Boolean,
                   Optional isModel3D As Boolean = False)
        Me.New(selectionId, bridgeId, bridgeLabel, classLabel, typeLabel, itemLabel,
               values, supportsWidth, supportsBody, isModel3D, False)
    End Sub

    Public Sub New(selectionId As String,
                   bridgeId As String,
                   bridgeLabel As String,
                   classLabel As String,
                   typeLabel As String,
                   itemLabel As String,
                   values As BridgeAppearanceValues,
                   supportsWidth As Boolean,
                   supportsBody As Boolean,
                   isModel3D As Boolean,
                   supportsFill As Boolean)
        Me.SelectionId = selectionId
        Me.BridgeId = bridgeId
        Me.BridgeLabel = bridgeLabel
        Me.ClassLabel = classLabel
        Me.TypeLabel = typeLabel
        Me.ItemLabel = itemLabel
        Me.Values = values
        Me.SupportsWidth = supportsWidth
        Me.SupportsBody = supportsBody
        Me.IsModel3D = isModel3D
        Me.SupportsFill = supportsFill
    End Sub

    Public ReadOnly Property SelectionId As String
    Public ReadOnly Property BridgeId As String
    Public ReadOnly Property BridgeLabel As String
    Public ReadOnly Property ClassLabel As String
    Public ReadOnly Property TypeLabel As String
    Public ReadOnly Property ItemLabel As String
    Public Property Values As BridgeAppearanceValues
    Public ReadOnly Property SupportsWidth As Boolean
    Public ReadOnly Property SupportsBody As Boolean
    Public ReadOnly Property IsModel3D As Boolean
    Public ReadOnly Property SupportsFill As Boolean
End Class

Friend NotInheritable Class BridgeAppearanceNativeTarget
    Public Property Info As BridgeAppearanceItem
    Public Property Data As StructureElement
    Public Property Entity As DwgEntity
End Class

Friend NotInheritable Class BridgeAppearanceApplyResult
    Public Sub New(appliedCount As Integer, skippedBodyLabels As IEnumerable(Of String))
        Me.AppliedCount = appliedCount
        Me.SkippedBodyLabels = If(skippedBodyLabels, Enumerable.Empty(Of String)()).ToArray()
    End Sub

    Public ReadOnly Property AppliedCount As Integer
    Public ReadOnly Property SkippedBodyLabels As IReadOnlyList(Of String)
End Class

Friend NotInheritable Class BridgeAppearanceNativeRuntime
    Private Sub New()
    End Sub

    Public Shared Function Collect(drawing As Drawing) As List(Of BridgeAppearanceNativeTarget)
        Dim result As New List(Of BridgeAppearanceNativeTarget)()
        If drawing Is Nothing OrElse drawing.ActiveSpace Is Nothing Then Return result

        Dim bridgeIndex As Dictionary(Of String, BridgeTaggedEntityRecord) =
            BridgeTaggedEntityIndex.Build(drawing.ActiveSpace.Entities.OfType(Of DwgEntity)())
        For Each record As BridgeTaggedEntityRecord In bridgeIndex.Values
            Dim bridgeLabel As String = "Мост " & ShortId(record.IdStructure)
            If record.Bridge IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(record.Bridge.NameBridge) Then
                bridgeLabel = record.Bridge.NameBridge
            End If

            For Each entity As DwgEntity In record.Entities
                Dim data As StructureElement = Nothing
                If Not BridgeTaggedEntityIndex.TryReadStructureData(entity, data) Then Continue For

                Dim supportsBody As Boolean = False
                Dim values As BridgeAppearanceValues = CaptureValues(entity, supportsBody)
                Dim selectionId As String = Guid.NewGuid().ToString("N")
                Dim info As New BridgeAppearanceItem(
                    selectionId,
                    record.IdStructure,
                    bridgeLabel,
                    StructureElement.GetDescription(data.ClassObject),
                    StructureElement.GetDescription(data.Name),
                    GetItemLabel(data),
                    values,
                    TypeOf entity Is DwgPolyline,
                    supportsBody,
                    TypeOf entity Is DwgModel3DElement,
                    TypeOf entity Is BridgeBeamPlanEntity)
                result.Add(New BridgeAppearanceNativeTarget With {
                    .Info = info,
                    .Data = data,
                    .Entity = entity
                })
            Next
        Next
        Return result
    End Function

    Public Shared Function CaptureValues(entity As DwgEntity,
                                         ByRef supportsBody As Boolean) As BridgeAppearanceValues
        If entity Is Nothing Then Throw New ArgumentNullException(NameOf(entity))

        Dim width As Double? = Nothing
        Dim polyline As DwgPolyline = TryCast(entity, DwgPolyline)
        If polyline IsNot Nothing Then width = polyline.Width

        Dim bodyArgb As Integer? = Nothing
        Dim bodyColors As Integer() = CaptureBodyColors(entity)
        supportsBody = bodyColors IsNot Nothing AndAlso bodyColors.Length > 0
        If supportsBody Then
            Dim firstArgb As Integer = bodyColors(0)
            If bodyColors.All(Function(argb) argb = firstArgb) Then bodyArgb = firstArgb
        End If

        Dim planEntity As BridgeBeamPlanEntity = TryCast(entity, BridgeBeamPlanEntity)
        Return New BridgeAppearanceValues(
            entity.Color.ToCompressValue(),
            If(entity.Layer Is Nothing, Nothing, entity.Layer.Name),
            If(entity.Linetype Is Nothing, Nothing, entity.Linetype.Name),
            entity.LinetypeScale,
            CInt(entity.Lineweight),
            width,
            bodyArgb,
            If(planEntity Is Nothing, DirectCast(Nothing, Integer?), planEntity.FillColor.ToCompressValue()),
            If(planEntity Is Nothing, Nothing, planEntity.HatchPatternName),
            If(planEntity Is Nothing, DirectCast(Nothing, Double?), planEntity.HatchScale),
            If(planEntity Is Nothing, DirectCast(Nothing, Double?), planEntity.HatchAngle))
    End Function

    Public Shared Function Apply(drawing As Drawing,
                                 targets As IEnumerable(Of BridgeAppearanceNativeTarget),
                                 patches As IDictionary(Of String, BridgeAppearancePatch)) As Integer
        Return ApplyDetailed(drawing, targets, patches).AppliedCount
    End Function

    Public Shared Function ApplyDetailed(drawing As Drawing,
                                         targets As IEnumerable(Of BridgeAppearanceNativeTarget),
                                         patches As IDictionary(Of String, BridgeAppearancePatch)) As BridgeAppearanceApplyResult
        If drawing Is Nothing Then Throw New ArgumentNullException(NameOf(drawing))
        If targets Is Nothing Then Throw New ArgumentNullException(NameOf(targets))
        If patches Is Nothing Then Throw New ArgumentNullException(NameOf(patches))

        Dim changes As New List(Of PreparedChange)()
        Dim skippedBodyLabels As New List(Of String)()
        For Each target As BridgeAppearanceNativeTarget In targets
            Dim patch As BridgeAppearancePatch = Nothing
            If target Is Nothing OrElse target.Info Is Nothing OrElse
               Not patches.TryGetValue(target.Info.SelectionId, patch) OrElse patch Is Nothing OrElse patch.IsEmpty Then
                Continue For
            End If

            If target.Entity Is Nothing OrElse target.Entity.Drawing IsNot drawing Then
                Throw New InvalidOperationException("Выбранный элемент больше не принадлежит текущему чертежу.")
            End If

            Dim supportsBody As Boolean = False
            Dim original As BridgeAppearanceValues = CaptureValues(target.Entity, supportsBody)
            Dim bodyRequested As Boolean = patch.BodyArgb.HasValue
            patch = patch.ForTarget(target.Info.SupportsWidth, supportsBody, target.Info.SupportsFill)
            If bodyRequested AndAlso Not supportsBody AndAlso TypeOf target.Entity Is DwgModel3DElement Then
                skippedBodyLabels.Add(AppearanceLabel(target.Info))
            End If
            If patch.IsEmpty Then Continue For

            Dim reason As String = Nothing
            If Not patch.TryValidate(reason) Then Throw New ArgumentException(reason, NameOf(patches))
            Dim updated As BridgeAppearanceValues = patch.ApplyTo(original)
            If IsNoOp(patch, original) Then Continue For
            Dim prepared As New PreparedChange With {
                .Target = target,
                .Patch = patch,
                .Original = original,
                .Updated = updated
            }

            If patch.LayerName IsNot Nothing Then
                If Not drawing.Layers.Names.ContainsKey(patch.LayerName) Then
                    Throw New InvalidOperationException("Слой «" & patch.LayerName & "» отсутствует в текущем чертеже.")
                End If
                prepared.Layer = drawing.Layers(patch.LayerName)
            End If
            If patch.LinetypeName IsNot Nothing Then
                If Not drawing.Linetypes.Names.ContainsKey(patch.LinetypeName) Then
                    Throw New InvalidOperationException("Тип линии «" & patch.LinetypeName & "» отсутствует в текущем чертеже.")
                End If
                prepared.Linetype = drawing.Linetypes(patch.LinetypeName)
            End If
            If patch.Lineweight.HasValue AndAlso
               Not [Enum].IsDefined(GetType(Lineweight), patch.Lineweight.Value) Then
                Throw New InvalidOperationException("Выбрана неподдерживаемая толщина линии.")
            End If

            If patch.BodyArgb.HasValue Then
                Dim model As DwgModel3DElement = TryCast(target.Entity, DwgModel3DElement)
                prepared.OriginalModel = CloneElementForAppearance(model.Element)
                prepared.UpdatedModel = CloneElementForAppearance(model.Element)
                SetBodyColor(prepared.UpdatedModel, Color.FromArgb(patch.BodyArgb.Value))
            End If
            changes.Add(prepared)
        Next
        If changes.Count = 0 Then Return New BridgeAppearanceApplyResult(0, skippedBodyLabels)

        Dim applied As New List(Of PreparedChange)()
        Dim applyFailure As Exception = Nothing
        drawing.BeginUpdate("Оформление моста")
        Try
            For Each change As PreparedChange In changes
                applied.Add(change)
                ApplyPrepared(change)
            Next
        Catch ex As Exception
            Dim rollbackFailures As New List(Of Exception)()
            For index As Integer = applied.Count - 1 To 0 Step -1
                Try
                    RollbackPrepared(applied(index), drawing)
                Catch rollbackFailure As Exception
                    rollbackFailures.Add(rollbackFailure)
                End Try
            Next
            If rollbackFailures.Count > 0 Then
                rollbackFailures.Insert(0, ex)
                applyFailure = New AggregateException(
                    "Не удалось применить оформление; при восстановлении части элементов возникла дополнительная ошибка.",
                    rollbackFailures)
            Else
                applyFailure = ex
            End If
        End Try
        Try
            drawing.EndUpdate()
        Catch endFailure As Exception
            If applyFailure Is Nothing Then
                applyFailure = endFailure
            Else
                applyFailure = New AggregateException(
                    "Не удалось завершить изменение оформления.",
                    applyFailure,
                    endFailure)
            End If
        End Try
        If applyFailure IsNot Nothing Then Throw applyFailure

        For Each change As PreparedChange In changes
            Dim supportsBody As Boolean = False
            change.Target.Info.Values = CaptureValues(change.Target.Entity, supportsBody)
        Next
        Return New BridgeAppearanceApplyResult(changes.Count, skippedBodyLabels)
    End Function

    Private Shared Function IsNoOp(patch As BridgeAppearancePatch,
                                   original As BridgeAppearanceValues) As Boolean
        Return (Not patch.CadColorValue.HasValue OrElse patch.CadColorValue.Value = original.CadColorValue) AndAlso
               (patch.LayerName Is Nothing OrElse String.Equals(patch.LayerName, original.LayerName, StringComparison.Ordinal)) AndAlso
               (patch.LinetypeName Is Nothing OrElse String.Equals(patch.LinetypeName, original.LinetypeName, StringComparison.Ordinal)) AndAlso
               (Not patch.LinetypeScale.HasValue OrElse patch.LinetypeScale.Value = original.LinetypeScale) AndAlso
               (Not patch.Lineweight.HasValue OrElse patch.Lineweight.Value = original.Lineweight) AndAlso
               (Not patch.Width.HasValue OrElse patch.Width.Equals(original.Width)) AndAlso
               (Not patch.BodyArgb.HasValue OrElse patch.BodyArgb.Equals(original.BodyArgb)) AndAlso
               (Not patch.FillCadColorValue.HasValue OrElse patch.FillCadColorValue.Equals(original.FillCadColorValue)) AndAlso
               (patch.HatchPatternName Is Nothing OrElse String.Equals(patch.HatchPatternName, original.HatchPatternName, StringComparison.OrdinalIgnoreCase)) AndAlso
               (Not patch.HatchScale.HasValue OrElse patch.HatchScale.Equals(original.HatchScale)) AndAlso
               (Not patch.HatchAngle.HasValue OrElse patch.HatchAngle.Equals(original.HatchAngle))
    End Function

    Public Shared Function CloneElementForAppearance(source As ImElement) As ImElement
        If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))

        Dim document As New StgDocument()
        Dim saveContext As New SerializationContext(source)
        source.SaveToStg(document.Body, saveContext)
        saveContext.SaveToStg(document.Header)

        Using stream As New MemoryStream()
            document.SaveToStreamAsBinary(stream)
            stream.Position = 0
            Dim loaded As New StgDocument()
            loaded.LoadFromStreamAsBinary(stream)
            Dim loadContext As New SerializationContext(Nothing)
            loadContext.LoadFromStg(loaded.Header)
            Return DirectCast(TypedObject.LoadFromStg(loaded.Body, loadContext), ImElement)
        End Using
    End Function

    Public Shared Function CaptureBodyColors(entity As DwgEntity) As Integer()
        Dim model As DwgModel3DElement = TryCast(entity, DwgModel3DElement)
        If model Is Nothing OrElse model.Element Is Nothing Then Return Nothing
        Dim solids As List(Of StaticSolidElement) = GetSolidElements(model.Element)
        Dim materials As List(Of PhongMaterial) = GetPhongMaterials(model.Element)
        If solids.Count = 0 AndAlso materials.Count = 0 Then Return Nothing
        Return solids.Select(Function(solid) solid.Color.ToArgb()).
            Concat(materials.Select(Function(material) PhongDiffuseToColor(material).ToArgb())).
            ToArray()
    End Function

    Public Shared Function CreateBodyColorClone(entity As DwgEntity,
                                                colors As Integer()) As ImElement
        If colors Is Nothing OrElse colors.Length = 0 Then Return Nothing
        Dim model As DwgModel3DElement = TryCast(entity, DwgModel3DElement)
        If model Is Nothing OrElse model.Element Is Nothing Then
            Throw New InvalidOperationException("Для перестроенного 3D-элемента недоступно восстановление цвета тела.")
        End If
        Dim clone As ImElement = CloneElementForAppearance(model.Element)
        Dim solids As List(Of StaticSolidElement) = GetSolidElements(clone)
        Dim materials As List(Of PhongMaterial) = GetPhongMaterials(clone)
        Dim bodyArgb As Integer = colors(0)
        If colors.Any(Function(value) value <> bodyArgb) Then
            Throw New InvalidOperationException("Различные цвета частей 3D-элемента нельзя восстановить без однозначного соответствия частей.")
        End If
        If solids.Count = 0 AndAlso materials.Count = 0 Then
            Throw New InvalidOperationException("Для перестроенного 3D-элемента недоступно восстановление цвета тела.")
        End If
        For Each solid As StaticSolidElement In solids
            solid.Color = Color.FromArgb(bodyArgb)
        Next
        For Each material As PhongMaterial In materials
            SetPhongDiffuse(material, Color.FromArgb(bodyArgb))
        Next
        Return clone
    End Function

    Public Shared Sub AssignBodyColorClone(entity As DwgEntity, clone As ImElement)
        AssignModel(TryCast(entity, DwgModel3DElement), clone)
    End Sub

    Private Shared Sub ApplyPrepared(change As PreparedChange)
        Dim entity As DwgEntity = change.Target.Entity
        If change.Patch.CadColorValue.HasValue Then entity.Color = CadColor.FromCompressValue(change.Patch.CadColorValue.Value)
        If change.Layer IsNot Nothing Then entity.Layer = change.Layer
        If change.Linetype IsNot Nothing Then entity.Linetype = change.Linetype
        If change.Patch.LinetypeScale.HasValue Then entity.LinetypeScale = change.Patch.LinetypeScale.Value
        If change.Patch.Lineweight.HasValue Then entity.Lineweight = CType(change.Patch.Lineweight.Value, Lineweight)
        If change.Patch.Width.HasValue Then
            Dim polyline As DwgPolyline = TryCast(entity, DwgPolyline)
            If polyline IsNot Nothing Then polyline.Width = change.Patch.Width.Value
        End If
        Dim planEntity As BridgeBeamPlanEntity = TryCast(entity, BridgeBeamPlanEntity)
        If planEntity IsNot Nothing Then
            If change.Patch.FillCadColorValue.HasValue Then
                planEntity.FillColor = CadColor.FromCompressValue(change.Patch.FillCadColorValue.Value)
            End If
            If change.Patch.HatchPatternName IsNot Nothing Then planEntity.HatchPatternName = change.Patch.HatchPatternName
            If change.Patch.HatchScale.HasValue Then planEntity.HatchScale = change.Patch.HatchScale.Value
            If change.Patch.HatchAngle.HasValue Then planEntity.HatchAngle = change.Patch.HatchAngle.Value
        End If
        If change.UpdatedModel IsNot Nothing Then AssignModel(TryCast(entity, DwgModel3DElement), change.UpdatedModel)
    End Sub

    Private Shared Sub RollbackPrepared(change As PreparedChange, drawing As Drawing)
        Dim entity As DwgEntity = change.Target.Entity
        entity.Color = CadColor.FromCompressValue(change.Original.CadColorValue)
        If Not String.IsNullOrEmpty(change.Original.LayerName) AndAlso drawing.Layers.Names.ContainsKey(change.Original.LayerName) Then
            entity.Layer = drawing.Layers(change.Original.LayerName)
        End If
        If Not String.IsNullOrEmpty(change.Original.LinetypeName) AndAlso drawing.Linetypes.Names.ContainsKey(change.Original.LinetypeName) Then
            entity.Linetype = drawing.Linetypes(change.Original.LinetypeName)
        End If
        entity.LinetypeScale = change.Original.LinetypeScale
        entity.Lineweight = CType(change.Original.Lineweight, Lineweight)
        Dim polyline As DwgPolyline = TryCast(entity, DwgPolyline)
        If polyline IsNot Nothing AndAlso change.Original.Width.HasValue Then polyline.Width = change.Original.Width.Value
        Dim planEntity As BridgeBeamPlanEntity = TryCast(entity, BridgeBeamPlanEntity)
        If planEntity IsNot Nothing Then
            If change.Original.FillCadColorValue.HasValue Then
                planEntity.FillColor = CadColor.FromCompressValue(change.Original.FillCadColorValue.Value)
            End If
            If change.Original.HatchPatternName IsNot Nothing Then planEntity.HatchPatternName = change.Original.HatchPatternName
            If change.Original.HatchScale.HasValue Then planEntity.HatchScale = change.Original.HatchScale.Value
            If change.Original.HatchAngle.HasValue Then planEntity.HatchAngle = change.Original.HatchAngle.Value
        End If
        If change.OriginalModel IsNot Nothing Then AssignModel(TryCast(entity, DwgModel3DElement), change.OriginalModel)
    End Sub

    Private Shared Sub AssignModel(model As DwgModel3DElement, element As ImElement)
        If model Is Nothing OrElse element Is Nothing Then Return
        model.BeginChange()
        Try
            model.Element = element
        Finally
            model.EndChange()
        End Try
    End Sub

    Private Shared Sub SetBodyColor(element As ImElement, color As Color)
        For Each solid As StaticSolidElement In GetSolidElements(element)
            solid.Color = color
        Next
        For Each material As PhongMaterial In GetPhongMaterials(element)
            SetPhongDiffuse(material, color)
        Next
    End Sub

    Private Shared Function GetPhongMaterials(root As ImElement) As List(Of PhongMaterial)
        Dim result As New List(Of PhongMaterial)()
        Dim visitedElements As New HashSet(Of ImElement)(ReferenceComparer(Of ImElement).Instance)
        Dim visitedMaterials As New HashSet(Of PhongMaterial)(ReferenceComparer(Of PhongMaterial).Instance)
        CollectPhongMaterials(root, visitedElements, visitedMaterials, result)
        Return result
    End Function

    Private Shared Sub CollectPhongMaterials(element As ImElement,
                                              visitedElements As HashSet(Of ImElement),
                                              visitedMaterials As HashSet(Of PhongMaterial),
                                              result As List(Of PhongMaterial))
        If element Is Nothing OrElse Not visitedElements.Add(element) Then Return
        Dim staticElement As Static3DElement = TryCast(element, Static3DElement)
        If staticElement IsNot Nothing Then
            Dim model As GeometryModel3D = staticElement.GetModel()
            If model IsNot Nothing AndAlso model.Materials IsNot Nothing Then
                For Each material As PhongMaterial In model.Materials.Values.OfType(Of PhongMaterial)()
                    If visitedMaterials.Add(material) Then result.Add(material)
                Next
            End If
        End If
        Dim wrapper As ImElementWrapper = TryCast(element, ImElementWrapper)
        If wrapper IsNot Nothing Then CollectPhongMaterials(wrapper.Element, visitedElements, visitedMaterials, result)
        For Each reference As IImElementReference In element.GetReferences()
            If reference IsNot Nothing Then
                CollectPhongMaterials(reference.Element, visitedElements, visitedMaterials, result)
            End If
        Next
    End Sub

    Private Shared Function PhongDiffuseToColor(material As PhongMaterial) As Color
        Dim diffuse As Vector3F = material.Diffuse
        Return Color.FromArgb(255,
                              ColorComponent(diffuse.X),
                              ColorComponent(diffuse.Y),
                              ColorComponent(diffuse.Z))
    End Function

    Private Shared Function ColorComponent(value As Single) As Integer
        Return CInt(Math.Round(Math.Max(0.0F, Math.Min(1.0F, value)) * 255.0F))
    End Function

    Private Shared Sub SetPhongDiffuse(material As PhongMaterial, color As Color)
        material.Diffuse = New Vector3F(color.R / 255.0F,
                                        color.G / 255.0F,
                                        color.B / 255.0F)
    End Sub

    Private Shared Function GetSolidElements(root As ImElement) As List(Of StaticSolidElement)
        Dim result As New List(Of StaticSolidElement)()
        Dim visited As New HashSet(Of ImElement)(ReferenceComparer(Of ImElement).Instance)
        CollectSolidElements(root, visited, result)
        Return result
    End Function

    Private Shared Sub CollectSolidElements(element As ImElement,
                                            visited As HashSet(Of ImElement),
                                            result As List(Of StaticSolidElement))
        If element Is Nothing OrElse Not visited.Add(element) Then Return
        Dim solid As StaticSolidElement = TryCast(element, StaticSolidElement)
        If solid IsNot Nothing Then result.Add(solid)
        Dim wrapper As ImElementWrapper = TryCast(element, ImElementWrapper)
        If wrapper IsNot Nothing Then CollectSolidElements(wrapper.Element, visited, result)
        For Each reference As IImElementReference In element.GetReferences()
            If reference IsNot Nothing Then CollectSolidElements(reference.Element, visited, result)
        Next
    End Sub

    Private Shared Function GetItemLabel(data As StructureElement) As String
        Dim semanticLabel As String = GetSemanticLabel(data)
        If Not String.IsNullOrEmpty(semanticLabel) Then Return semanticLabel
        If Not String.IsNullOrWhiteSpace(data.Description) Then Return data.Description
        Return StructureElement.GetDescription(data.Name) & " · " & ShortId(data.IdElement)
    End Function

    Private Shared Function AppearanceLabel(info As BridgeAppearanceItem) As String
        If info Is Nothing Then Return "неизвестный элемент"
        If String.IsNullOrWhiteSpace(info.ItemLabel) Then Return info.TypeLabel
        If String.IsNullOrWhiteSpace(info.TypeLabel) Then Return info.ItemLabel
        Return info.TypeLabel & " — " & info.ItemLabel
    End Function

    Private Shared Function GetSemanticLabel(data As StructureElement) As String
        If String.IsNullOrWhiteSpace(data.KeyParameter) Then Return Nothing
        Try
            Dim json As JObject = JObject.Parse(data.KeyParameter)
            If data.ClassObject = StructureElement.classStructure.BeamI Then
                Return NumberLabel(json, "Пролёт", "numberProlet", "ряд", "numberRow")
            End If
            If data.ClassObject = StructureElement.classStructure.PilePillar Then
                Dim pillar As String = JsonValue(json, "NumberPillar")
                Dim row As String = JsonValue(json, "NumberRow")
                Dim column As String = JsonValue(json, "NumberColumn")
                If pillar IsNot Nothing AndAlso row IsNot Nothing AndAlso column IsNot Nothing Then
                    Return "Опора " & pillar & ", ряд " & row & ", столбец " & column
                End If
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function NumberLabel(json As JObject,
                                        firstCaption As String,
                                        firstName As String,
                                        secondCaption As String,
                                        secondName As String) As String
        Dim first As String = JsonValue(json, firstName)
        Dim second As String = JsonValue(json, secondName)
        If first Is Nothing OrElse second Is Nothing Then Return Nothing
        Return firstCaption & " " & first & ", " & secondCaption & " " & second
    End Function

    Private Shared Function JsonValue(json As JObject, name As String) As String
        Dim propertyValue = json.Properties().FirstOrDefault(
            Function(item) String.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
        If propertyValue Is Nothing OrElse propertyValue.Value.Type = JTokenType.Null Then Return Nothing
        Return propertyValue.Value.ToString()
    End Function

    Private Shared Function ShortId(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return "без идентификатора"
        Return value.Substring(0, Math.Min(8, value.Length))
    End Function

    Private NotInheritable Class PreparedChange
        Public Target As BridgeAppearanceNativeTarget
        Public Patch As BridgeAppearancePatch
        Public Original As BridgeAppearanceValues
        Public Updated As BridgeAppearanceValues
        Public Layer As DwgLayer
        Public Linetype As DwgLinetype
        Public OriginalModel As ImElement
        Public UpdatedModel As ImElement
    End Class

    Private NotInheritable Class ReferenceComparer(Of T As Class)
        Implements IEqualityComparer(Of T)

        Public Shared ReadOnly Instance As New ReferenceComparer(Of T)()

        Public Overloads Function Equals(x As T, y As T) As Boolean Implements IEqualityComparer(Of T).Equals
            Return Object.ReferenceEquals(x, y)
        End Function

        Public Overloads Function GetHashCode(value As T) As Integer Implements IEqualityComparer(Of T).GetHashCode
            Return RuntimeHelpers.GetHashCode(value)
        End Function
    End Class
End Class
