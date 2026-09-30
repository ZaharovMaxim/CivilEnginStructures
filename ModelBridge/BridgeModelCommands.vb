Imports System.Linq
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements

Namespace RopExample1
    Partial Public Class RopExample1Module
        <cmd("CreateInfrastradaBridgeModel")>
        Public Function CreateInfrastradaBridgeModel() As BridgesModelEditor
            Return New BridgesModelEditor()
        End Function

        <cmd("edit_infrastrada_bridge_settings")>
        Public Sub EditInfrastradaBridgeSettings(pathId As String)
            Dim operation As New BuildOperationContext(
                "Настройка модели мостов",
                "Проверьте, что модель мостов открыта и доступна для редактирования.",
                "RopExample1Module.EditInfrastradaBridgeSettings")
            Dim projectModel As IProjectModel = Nothing
            Dim writeLocked As Boolean = False
            Dim failure As Exception = Nothing
            Try
                operation.Stage = "Поиск модели мостов"
                If String.IsNullOrWhiteSpace(pathId) Then
                    operation.Fail("Не передан путь к модели мостов.")
                End If
                projectModel = PluginCoreOps.FindModel(pathId)
                ValidateBridgeProjectModel(projectModel, operation)

                operation.Stage = "Блокировка модели для записи"
                projectModel.LockWrite()
                writeLocked = True

                operation.Stage = "Чтение настроек модели"
                Dim model As ArrangementModel = BridgeModelRuntime.GetArrangement(projectModel.Model)
                If model Is Nothing Then
                    operation.Fail("В выбранной модели нет данных мостов.")
                End If
                Dim settings As BridgeModelSettings = BridgeModelSettingsStore.GetSettings(model)
                operation.Stage = "Открытие окна настроек"
                Using dialog As New BridgeModelSettingsDialog(projectModel, settings)
                    If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        operation.Stage = "Сохранение настроек поверхностей"
                        operation.MarkModelMutationStarted()
                        BridgeModelSettingsStore.UpdateSurfaceReferences(model,
                                                                         dialog.ProjectSurfaceRelativePath,
                                                                         dialog.EarthSurfaceRelativePath)
                        projectModel.Modified = True
                        projectModel.ReferencesModified = True
                    End If
                End Using
            Catch ex As Exception
                failure = ex
            Finally
                UnlockProjectModel(projectModel, writeLocked, operation, failure)
            End Try
            If failure IsNot Nothing Then operation.Report(failure)
        End Sub

        <cmd("GroupInfrastradaBridges")>
        Public Sub GroupInfrastradaBridges()
            Dim operation As New BuildOperationContext(
                "Объединение элементов моста",
                "Сделайте активной модель «Мосты» и повторите операцию.",
                "RopExample1Module.GroupInfrastradaBridges")
            Dim projectModel As IProjectModel = Nothing
            Dim writeLocked As Boolean = False
            Dim failure As Exception = Nothing
            Try
                Dim layer As BridgePlanCompoundLayer = ResolveActiveBridgePlan(operation, projectModel)
                operation.Stage = "Блокировка модели для записи"
                projectModel.LockWrite()
                writeLocked = True

                operation.Stage = "Объединение элементов по мостам"
                operation.MarkModelMutationStarted()
                Dim changed As Integer = BridgeDrawingGroupManager.Synchronize(layer.DrawingLayer.Drawing)
                If changed > 0 Then projectModel.Modified = True
            Catch ex As Exception
                failure = ex
            Finally
                UnlockProjectModel(projectModel, writeLocked, operation, failure)
            End Try
            If failure IsNot Nothing Then operation.Report(failure)
        End Sub

        <cmd("EditInfrastradaBridgeAppearance")>
        Public Sub EditInfrastradaBridgeAppearance()
            Dim operation As New BuildOperationContext(
                "Изменение оформления моста",
                "Проверьте выбранные элементы, слои и типы линий, затем повторите операцию.",
                "RopExample1Module.EditInfrastradaBridgeAppearance")
            Dim projectModel As IProjectModel = Nothing
            Dim writeLocked As Boolean = False
            Dim failure As Exception = Nothing
            Try
                Dim layer As BridgePlanCompoundLayer = ResolveActiveBridgePlan(operation, projectModel)
                operation.Stage = "Блокировка модели для записи"
                projectModel.LockWrite()
                writeLocked = True

                operation.Stage = "Сбор элементов мостов"
                Dim targets As List(Of BridgeAppearanceNativeTarget) =
                    BridgeAppearanceNativeRuntime.Collect(layer.DrawingLayer.Drawing)
                If targets.Count = 0 Then
                    Windows.Forms.MessageBox.Show(
                        "В текущей модели не найдены элементы мостов с доступным оформлением.",
                        "Оформление моста",
                        Windows.Forms.MessageBoxButtons.OK,
                        Windows.Forms.MessageBoxIcon.Information)
                Else
                    operation.Stage = "Открытие редактора оформления"
                    Using dialog As New BridgeAppearanceDialog(
                        targets.Select(Function(target) target.Info),
                        layer.DrawingLayer.Drawing.Layers.Names.Keys,
                        layer.DrawingLayer.Drawing.Linetypes.Names.Keys)
                        AddHandler dialog.ApplyRequested,
                            Sub(sender As Object, args As BridgeAppearanceApplyEventArgs)
                                operation.Stage = "Применение оформления"
                                Try
                                    Dim selected = New HashSet(Of String)(args.SelectionIds, StringComparer.Ordinal)
                                    Dim selectedTargets = targets.Where(Function(target) selected.Contains(target.Info.SelectionId)).ToList()
                                    Dim patches As New Dictionary(Of String, BridgeAppearancePatch)(StringComparer.Ordinal)
                                    For Each target As BridgeAppearanceNativeTarget In selectedTargets
                                        patches(target.Info.SelectionId) = args.Patch
                                    Next
                                    Dim result As BridgeAppearanceApplyResult = BridgeAppearanceNativeRuntime.ApplyDetailed(
                                        layer.DrawingLayer.Drawing,
                                        selectedTargets,
                                        patches)
                                    args.AppliedCount = result.AppliedCount
                                    args.SkippedBodyLabels = result.SkippedBodyLabels
                                    If args.AppliedCount > 0 Then projectModel.Modified = True
                                Catch ex As Exception
                                    args.ErrorMessage = ex.Message
                                    operation.Report(ex)
                                End Try
                            End Sub
                        dialog.ShowDialog()
                    End Using
                End If
            Catch ex As Exception
                failure = ex
            Finally
                UnlockProjectModel(projectModel, writeLocked, operation, failure)
            End Try
            If failure IsNot Nothing Then operation.Report(failure)
        End Sub

        Private Function ResolveActiveBridgePlan(operation As BuildOperationContext,
                                                 ByRef projectModel As IProjectModel) As BridgePlanCompoundLayer
            operation.Stage = "Поиск активной модели мостов"
            If Me.CadView Is Nothing Then
                operation.Fail("Нет активного вида плана.")
            End If

            Dim layer As BridgePlanCompoundLayer = BridgePlanContextResolver.Resolve(Me.CadView)
            If layer Is Nothing Then
                operation.Fail("Активная модель не является моделью «Мосты» или заблокирована.")
            End If
            If layer.DrawingLayer Is Nothing OrElse layer.DrawingLayer.Drawing Is Nothing Then
                operation.Fail("В активной модели мостов не загружен чертёж.")
            End If
            If String.IsNullOrWhiteSpace(layer.ModelPathId) Then
                operation.Fail("Для активной модели мостов не определён путь в проекте.")
            End If

            projectModel = PluginCoreOps.FindModel(layer.ModelPathId)
            ValidateBridgeProjectModel(projectModel, operation)
            Return layer
        End Function

        Private Shared Sub ValidateBridgeProjectModel(projectModel As IProjectModel,
                                                      operation As BuildOperationContext)
            If projectModel Is Nothing Then
                operation.Fail("Модель мостов не найдена в текущем проекте.")
            End If
            If Not String.Equals(projectModel.ModelType, "infrastrada_bridges", StringComparison.OrdinalIgnoreCase) Then
                operation.Fail("Выбранная модель не является моделью «Мосты».")
            End If
        End Sub

        Private Shared Sub UnlockProjectModel(projectModel As IProjectModel,
                                              ByRef writeLocked As Boolean,
                                              operation As BuildOperationContext,
                                              ByRef failure As Exception)
            If Not writeLocked OrElse projectModel Is Nothing Then Return
            Try
                projectModel.UnlockWrite()
            Catch cleanupFailure As Exception
                If failure Is Nothing Then
                    failure = New BuildStageException(
                        "Разблокировка модели",
                        cleanupFailure.Message,
                        operation.SuggestedAction,
                        operation.SourceMethod,
                        cleanupFailure)
                Else
                    Dim stageFailure As BuildStageException = TryCast(failure, BuildStageException)
                    failure = New BuildStageException(
                        If(stageFailure Is Nothing, operation.Stage, stageFailure.Stage),
                        If(stageFailure Is Nothing, failure.Message, stageFailure.Reason),
                        If(stageFailure Is Nothing, operation.SuggestedAction, stageFailure.SuggestedAction),
                        If(stageFailure Is Nothing, operation.SourceMethod, stageFailure.SourceMethod),
                        New AggregateException(failure, cleanupFailure))
                End If
            Finally
                writeLocked = False
            End Try
        End Sub
    End Class
End Namespace
