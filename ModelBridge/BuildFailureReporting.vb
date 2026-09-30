Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Windows.Forms

Public NotInheritable Class BuildFailureReport
    Private ReadOnly _summary As String
    Private ReadOnly _technicalDetails As String

    Public Sub New(operation As String,
                   stage As String,
                   reason As String,
                   suggestedAction As String,
                   sourceMethod As String,
                   failure As Exception)
        Dim operationText As String = SafeText(operation, "Построение")
        Dim stageText As String = SafeText(stage, "Неизвестный этап")
        Dim reasonText As String = SafeText(reason, "Операция не завершена")
        Dim actionText As String = SafeText(suggestedAction, "Проверьте исходные данные и повторите операцию")

        _summary = String.Join(
            Environment.NewLine,
            "Операция: " & operationText,
            "Этап: " & stageText,
            "Причина: " & reasonText,
            "Что сделать: " & actionText)

        Dim assembly As Assembly = GetType(BuildFailureReport).Assembly
        Dim details As New StringBuilder(_summary)
        details.AppendLine().AppendLine().Append("Время: ").Append(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss zzz"))
        details.AppendLine().Append("Сборка: ").Append(assembly.FullName)
        details.AppendLine().Append("MVID: ").Append(GetType(BuildFailureReport).Module.ModuleVersionId.ToString())
        If Not String.IsNullOrWhiteSpace(sourceMethod) Then
            details.AppendLine().AppendLine().Append("Метод: ").Append(sourceMethod.Trim())
        End If
        If failure IsNot Nothing Then
            details.AppendLine().AppendLine().AppendLine("Исключение:").Append(failure.ToString())
            AppendSourceLocations(details, failure)
        End If
        _technicalDetails = details.ToString()
    End Sub

    Public ReadOnly Property Summary As String
        Get
            Return _summary
        End Get
    End Property

    Public ReadOnly Property TechnicalDetails As String
        Get
            Return _technicalDetails
        End Get
    End Property

    Private Shared Function SafeText(value As String, fallback As String) As String
        Return If(String.IsNullOrWhiteSpace(value), fallback, value.Trim())
    End Function

    Private Shared Sub AppendSourceLocations(details As StringBuilder, failure As Exception)
        Try
            Dim frames As StackFrame() = New StackTrace(failure, True).GetFrames()
            If frames Is Nothing Then Return

            Dim locations As New List(Of String)()
            For Each frame As StackFrame In frames
                Dim fileName As String = frame.GetFileName()
                If String.IsNullOrWhiteSpace(fileName) Then Continue For
                Dim location As String = fileName
                If frame.GetFileLineNumber() > 0 Then location &= ":" & frame.GetFileLineNumber().ToString()
                If Not locations.Contains(location) Then locations.Add(location)
            Next
            If locations.Count > 0 Then
                details.AppendLine().AppendLine().AppendLine("Исходный код:").Append(String.Join(Environment.NewLine, locations))
            End If
        Catch
            ' Exception.ToString above remains the authoritative diagnostic text.
        End Try
    End Sub
End Class

Public NotInheritable Class BuildValueValidation
    Private Sub New()
    End Sub

    Public Shared Function TryRequireFinite(reader As Func(Of Double?),
                                            valueName As String,
                                            ByRef value As Double,
                                            ByRef failureReason As String) As Boolean
        If reader Is Nothing Then
            failureReason = SafeValueName(valueName) & ": не задан способ чтения значения."
            Return False
        End If

        Dim candidate As Double? = reader()
        If Not candidate.HasValue OrElse Double.IsNaN(candidate.Value) OrElse Double.IsInfinity(candidate.Value) Then
            failureReason = SafeValueName(valueName) & ": отметка отсутствует или не является конечным числом."
            Return False
        End If

        value = candidate.Value
        failureReason = Nothing
        Return True
    End Function

    Private Shared Function SafeValueName(valueName As String) As String
        Return If(String.IsNullOrWhiteSpace(valueName), "Значение", valueName.Trim())
    End Function
End Class

Friend NotInheritable Class BuildStageException
    Inherits InvalidOperationException

    Public Sub New(stage As String,
                   reason As String,
                   suggestedAction As String,
                   sourceMethod As String,
                   Optional innerException As Exception = Nothing)
        MyBase.New(reason, innerException)
        Me.Stage = stage
        Me.Reason = reason
        Me.SuggestedAction = suggestedAction
        Me.SourceMethod = sourceMethod
    End Sub

    Public ReadOnly Stage As String
    Public ReadOnly Reason As String
    Public ReadOnly SuggestedAction As String
    Public ReadOnly SourceMethod As String
End Class

Friend NotInheritable Class BuildOperationContext
    Private _modelMutationStarted As Boolean

    Public Sub New(operation As String, suggestedAction As String, sourceMethod As String)
        Me.Operation = operation
        Me.SuggestedAction = suggestedAction
        Me.SourceMethod = sourceMethod
        Me.Stage = "Подготовка"
    End Sub

    Public ReadOnly Operation As String
    Public ReadOnly SuggestedAction As String
    Public ReadOnly SourceMethod As String
    Public Property Stage As String

    Public Sub MarkModelMutationStarted()
        _modelMutationStarted = True
    End Sub

    Public Sub Fail(reason As String, Optional suggestedActionOverride As String = Nothing)
        Throw New BuildStageException(
            Stage,
            reason,
            If(String.IsNullOrWhiteSpace(suggestedActionOverride), SuggestedAction, suggestedActionOverride),
            SourceMethod)
    End Sub

    Public Sub Report(failure As Exception)
        Dim stageFailure As BuildStageException = TryCast(failure, BuildStageException)
        Dim effectiveAction As String = If(
            _modelMutationStarted,
            "Операция могла выполниться частично. Проверьте модель и отмените изменения командой Undo перед повтором.",
            SuggestedAction)
        Dim report As BuildFailureReport
        If stageFailure Is Nothing Then
            report = New BuildFailureReport(
                Operation,
                Stage,
                If(failure Is Nothing, "Операция не завершена", failure.Message),
                effectiveAction,
                SourceMethod,
                failure)
        Else
            report = New BuildFailureReport(
                Operation,
                stageFailure.Stage,
                stageFailure.Reason,
                If(_modelMutationStarted, effectiveAction, stageFailure.SuggestedAction),
                stageFailure.SourceMethod,
                stageFailure)
        End If
        BuildFailurePresenter.Show(report)
    End Sub
End Class

Friend NotInheritable Class BuildFailurePresenter
    Private Sub New()
    End Sub

    Public Shared Sub Show(report As BuildFailureReport)
        If report Is Nothing Then Return
        WriteLog(report)
        Try
            Using dialog As New BuildFailureDialog(report)
                dialog.ShowDialog()
            End Using
        Catch
            MessageBox.Show(report.Summary, "Ошибка построения", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared ReadOnly Property LogPath As String
        Get
            Return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Infrastrada",
                "Logs",
                "bridge-build-errors.txt")
        End Get
    End Property

    Private Shared Sub WriteLog(report As BuildFailureReport)
        Try
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath))
            Dim entry As String = String.Join(
                Environment.NewLine,
                New String("="c, 80),
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                report.TechnicalDetails,
                String.Empty)
            File.AppendAllText(LogPath, entry, New UTF8Encoding(False))
        Catch
            ' Logging must never hide the original build failure.
        End Try
    End Sub
End Class

Friend NotInheritable Class BuildFailureDialog
    Inherits Form

    Private ReadOnly _report As BuildFailureReport
    Private ReadOnly _detailsCard As Panel
    Private ReadOnly _detailsText As TextBox
    Private ReadOnly _toggleButton As Button

    Public Sub New(report As BuildFailureReport)
        _report = report
        Font = New Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
        BackColor = Color.FromArgb(243, 246, 250)
        ForeColor = Color.FromArgb(31, 41, 55)
        Text = "Ошибка построения"
        StartPosition = FormStartPosition.CenterParent
        MinimizeBox = False
        MaximizeBox = False
        ShowIcon = False
        MinimumSize = New Size(680, 330)
        ClientSize = New Size(760, 360)

        Dim accent As New Panel With {
            .BackColor = Color.FromArgb(37, 99, 235),
            .Dock = DockStyle.Top,
            .Height = 5
        }
        Controls.Add(accent)

        Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(18),
            .ColumnCount = 1,
            .RowCount = 5
        }
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Controls.Add(layout)
        layout.BringToFront()

        Dim titleLabel As New Label With {
            .AutoSize = True,
            .Font = New Font(Font.FontFamily, 14.0F, FontStyle.Bold),
            .Text = "Не удалось завершить построение",
            .Margin = New Padding(0, 0, 0, 12)
        }
        layout.Controls.Add(titleLabel, 0, 0)

        Dim summaryCard As New Panel With {
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Dock = DockStyle.Fill,
            .Padding = New Padding(14),
            .Margin = New Padding(0)
        }
        Dim summaryText As New TextBox With {
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .Dock = DockStyle.Fill,
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Text = report.Summary
        }
        summaryCard.Controls.Add(summaryText)
        layout.Controls.Add(summaryCard, 0, 1)

        _detailsCard = New Panel With {
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Dock = DockStyle.Fill,
            .Height = 220,
            .Padding = New Padding(10),
            .Margin = New Padding(0, 12, 0, 0),
            .Visible = False
        }
        _detailsText = New TextBox With {
            .BackColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .Dock = DockStyle.Fill,
            .Font = New Font("Consolas", 9.0F),
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .WordWrap = False,
            .Text = report.TechnicalDetails
        }
        _detailsCard.Controls.Add(_detailsText)
        layout.Controls.Add(_detailsCard, 0, 2)

        Dim logLabel As New Label With {
            .AutoEllipsis = True,
            .Dock = DockStyle.Fill,
            .Text = "Журнал: " & BuildFailurePresenter.LogPath,
            .Margin = New Padding(0, 10, 0, 0)
        }
        layout.Controls.Add(logLabel, 0, 3)

        Dim buttons As New FlowLayoutPanel With {
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft,
            .Margin = New Padding(0, 14, 0, 0),
            .WrapContents = False
        }
        Dim closeButton As New Button With {.Text = "Закрыть", .AutoSize = True, .DialogResult = DialogResult.OK}
        Dim copyButton As New Button With {.Text = "Копировать детали", .AutoSize = True}
        _toggleButton = New Button With {.Text = "Показать детали", .AutoSize = True}
        AddHandler copyButton.Click, AddressOf CopyDetails
        AddHandler _toggleButton.Click, AddressOf ToggleDetails
        buttons.Controls.Add(closeButton)
        buttons.Controls.Add(copyButton)
        buttons.Controls.Add(_toggleButton)
        layout.Controls.Add(buttons, 0, 4)
        AcceptButton = closeButton
        CancelButton = closeButton
    End Sub

    Private Sub ToggleDetails(sender As Object, e As EventArgs)
        _detailsCard.Visible = Not _detailsCard.Visible
        _toggleButton.Text = If(_detailsCard.Visible, "Скрыть детали", "Показать детали")
        ClientSize = New Size(ClientSize.Width, If(_detailsCard.Visible, 620, 360))
    End Sub

    Private Sub CopyDetails(sender As Object, e As EventArgs)
        Try
            Clipboard.SetText(_report.TechnicalDetails)
        Catch
        End Try
    End Sub
End Class
