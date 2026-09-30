Imports System.IO
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Sfc

Public NotInheritable Class BridgeModelSettingsDialog
    Inherits Form

    Private ReadOnly ProjectSurfaceCombo As New ComboBox()
    Private ReadOnly EarthSurfaceCombo As New ComboBox()

    Public Sub New(ownerModel As IProjectModel, settings As BridgeModelSettings)
        Text = "Настройки модели «Мосты»"
        FormBorderStyle = FormBorderStyle.FixedDialog
        StartPosition = FormStartPosition.CenterParent
        MinimizeBox = False
        MaximizeBox = False
        ClientSize = New Drawing.Size(520, 170)

        Dim choices As List(Of SurfaceReferenceChoice) = GetSurfaceChoices(ownerModel)
        ConfigureCombo(ProjectSurfaceCombo, choices, settings.ProjectSurfaceRelativePath, 15, 30)
        ConfigureCombo(EarthSurfaceCombo, choices, settings.EarthSurfaceRelativePath, 15, 80)
        Controls.Add(New Label() With {.Text = "Проектная поверхность", .AutoSize = True, .Left = 15, .Top = 10})
        Controls.Add(New Label() With {.Text = "Поверхность земли", .AutoSize = True, .Left = 15, .Top = 60})
        Controls.Add(ProjectSurfaceCombo)
        Controls.Add(EarthSurfaceCombo)

        Dim okButton As New Button() With {.Text = "ОК", .DialogResult = DialogResult.OK, .Left = 355, .Top = 132, .Width = 70}
        Dim cancelButton As New Button() With {.Text = "Отмена", .DialogResult = DialogResult.Cancel, .Left = 435, .Top = 132, .Width = 70}
        Controls.Add(okButton)
        Controls.Add(cancelButton)
        AcceptButton = okButton
        Me.CancelButton = cancelButton
    End Sub

    Public ReadOnly Property ProjectSurfaceRelativePath As String
        Get
            Return DirectCast(ProjectSurfaceCombo.SelectedItem, SurfaceReferenceChoice).RelativePath
        End Get
    End Property

    Public ReadOnly Property EarthSurfaceRelativePath As String
        Get
            Return DirectCast(EarthSurfaceCombo.SelectedItem, SurfaceReferenceChoice).RelativePath
        End Get
    End Property

    Private Shared Sub ConfigureCombo(combo As ComboBox, choices As List(Of SurfaceReferenceChoice), selectedPath As String, left As Integer, top As Integer)
        Dim comboChoices As New List(Of SurfaceReferenceChoice)(choices)
        Dim selectedIndex As Integer = comboChoices.FindIndex(Function(item) String.Equals(item.RelativePath, selectedPath, StringComparison.OrdinalIgnoreCase))
        If selectedIndex < 0 AndAlso Not String.IsNullOrWhiteSpace(selectedPath) Then
            comboChoices.Add(New SurfaceReferenceChoice("(недоступна) " & selectedPath, selectedPath))
            selectedIndex = comboChoices.Count - 1
        End If
        combo.DropDownStyle = ComboBoxStyle.DropDownList
        combo.DisplayMember = NameOf(SurfaceReferenceChoice.DisplayName)
        combo.BindingContext = New BindingContext()
        combo.DataSource = comboChoices
        combo.Left = left
        combo.Top = top
        combo.Width = 490
        combo.SelectedIndex = Math.Max(0, selectedIndex)
    End Sub

    Private Shared Function GetSurfaceChoices(ownerModel As IProjectModel) As List(Of SurfaceReferenceChoice)
        Dim result As New List(Of SurfaceReferenceChoice) From {
            New SurfaceReferenceChoice("(не назначена)", String.Empty)
        }
        PluginCoreOps.FilterModels(
            Function(candidate As IProjectModel) As Boolean
                If candidate Is Nothing OrElse Object.ReferenceEquals(candidate, ownerModel) OrElse
                   candidate.Uri Is Nothing OrElse
                   Not SurfaceModelChoice.IsSupportedSurfacePath(candidate.Uri.LastPathComponent) Then Return False
                Dim relativePath As String = PluginCoreOps.FindModelRelativePath(ownerModel, candidate)
                If String.IsNullOrWhiteSpace(relativePath) Then Return False
                Dim displayName As String = Path.GetFileNameWithoutExtension(candidate.Uri.LastPathComponent)
                result.Add(New SurfaceReferenceChoice(displayName & " [" & candidate.ModelType & "] — " & relativePath, relativePath))
                Return False
            End Function)
        Return result
    End Function

    Private NotInheritable Class SurfaceReferenceChoice
        Public Sub New(displayName As String, relativePath As String)
            Me.DisplayName = displayName
            Me.RelativePath = relativePath
        End Sub

        Public ReadOnly Property DisplayName As String
        Public ReadOnly Property RelativePath As String
    End Class
End Class
