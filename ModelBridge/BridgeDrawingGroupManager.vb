Imports System.Linq
Imports System.Runtime.CompilerServices
Imports System.Security.Cryptography
Imports System.Text
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Public NotInheritable Class BridgeDrawingGroupManager
    Private Const GroupPrefix As String = "INFRA_BRIDGE_"
    Private Shared ReadOnly States As New ConditionalWeakTable(Of Drawing, DrawingState)()

    Private Sub New()
    End Sub

    Public Shared Function GetGroupName(idStructure As String) As String
        If String.IsNullOrWhiteSpace(idStructure) Then Return String.Empty

        Dim value As String = idStructure.Trim()
        Dim bridgeId As Guid
        If Guid.TryParse(value, bridgeId) Then
            Return GroupPrefix & bridgeId.ToString("N").ToLowerInvariant()
        End If

        Using algorithm As SHA256 = SHA256.Create()
            Dim hash As Byte() = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value))
            Dim result As New StringBuilder(GroupPrefix.Length + 7 + hash.Length * 2)
            result.Append(GroupPrefix).Append("legacy_")
            For Each item As Byte In hash
                result.Append(item.ToString("x2"))
            Next
            Return result.ToString()
        End Using
    End Function

    Public Shared Sub Attach(drawing As Drawing)
        If drawing Is Nothing Then Return

        Dim state As DrawingState = Nothing
        If States.TryGetValue(drawing, state) Then Return

        state = New DrawingState()
        Try
            States.Add(drawing, state)
        Catch ex As ArgumentException
            Return
        End Try

        AddHandler drawing.AfterAddEntity, AddressOf OnEntityChanged
        AddHandler drawing.ModifyEntity, AddressOf OnEntityChanged
        AddHandler drawing.BeforeRemoveEntity, AddressOf OnBeforeRemoveEntity
        AddHandler drawing.AfterUpdate, AddressOf OnAfterUpdate
    End Sub

    Public Shared Sub Detach(drawing As Drawing)
        If drawing Is Nothing Then Return

        Dim state As DrawingState = Nothing
        If Not States.TryGetValue(drawing, state) Then Return

        RemoveHandler drawing.AfterAddEntity, AddressOf OnEntityChanged
        RemoveHandler drawing.ModifyEntity, AddressOf OnEntityChanged
        RemoveHandler drawing.BeforeRemoveEntity, AddressOf OnBeforeRemoveEntity
        RemoveHandler drawing.AfterUpdate, AddressOf OnAfterUpdate
        States.Remove(drawing)
    End Sub

    Public Shared Sub TryGroupEntity(entity As DwgEntity)
        If entity Is Nothing OrElse entity.Drawing Is Nothing Then Return

        Dim state As DrawingState = Nothing
        If Not States.TryGetValue(entity.Drawing, state) OrElse state.IsHandling Then Return

        state.IsHandling = True
        Try
            If ApplyEntityGroup(entity) Then
                RemoveEmptyOwnedGroups(entity.Drawing)
            End If
        Catch
            ' Grouping must never change the result of the drawing operation that created the entity.
        Finally
            state.IsHandling = False
        End Try
    End Sub

    Public Shared Function Synchronize(drawing As Drawing) As Integer
        If drawing Is Nothing Then Return 0

        Dim state As DrawingState = Nothing
        States.TryGetValue(drawing, state)
        If state IsNot Nothing AndAlso state.IsHandling Then Return 0

        If state IsNot Nothing Then state.IsHandling = True
        Dim changed As Integer = 0
        Dim updateStarted As Boolean = False
        Try
            Dim assignments As New List(Of GroupAssignment)()
            For Each entity As DwgEntity In drawing.Blocks.SelectMany(Function(block As DwgBlock) block.Entities).ToArray()
                Dim expectedName As String = ReadExpectedGroupName(entity)
                Dim currentGroup As DwgGroup = entity.Group

                If expectedName.Length = 0 Then
                    If IsOwnedGroup(currentGroup) Then assignments.Add(New GroupAssignment(entity, String.Empty))
                ElseIf currentGroup Is Nothing Then
                    assignments.Add(New GroupAssignment(entity, expectedName))
                ElseIf IsOwnedGroup(currentGroup) AndAlso
                       Not String.Equals(currentGroup.Name, expectedName, StringComparison.Ordinal) Then
                    assignments.Add(New GroupAssignment(entity, expectedName))
                End If
            Next

            Dim hasEmptyGroups As Boolean = HasEmptyOwnedGroups(drawing)
            If assignments.Count = 0 AndAlso Not hasEmptyGroups Then Return 0

            drawing.BeginUpdate("Объединение элементов мостов")
            updateStarted = True
            For Each assignment As GroupAssignment In assignments
                If assignment.GroupName.Length = 0 Then
                    assignment.Entity.Group = Nothing
                Else
                    assignment.Entity.Group = GetOrCreateGroup(drawing, assignment.GroupName)
                End If
                changed += 1
            Next
            RemoveEmptyOwnedGroups(drawing)
            Return changed
        Finally
            If updateStarted Then drawing.EndUpdate()
            If state IsNot Nothing Then state.IsHandling = False
        End Try
    End Function

    Private Shared Function ApplyEntityGroup(entity As DwgEntity) As Boolean
        Dim expectedName As String = ReadExpectedGroupName(entity)
        Dim currentGroup As DwgGroup = entity.Group

        If expectedName.Length = 0 Then
            If Not IsOwnedGroup(currentGroup) Then Return False
            entity.Group = Nothing
            Return True
        End If

        If currentGroup IsNot Nothing Then
            If Not IsOwnedGroup(currentGroup) Then Return False
            If String.Equals(currentGroup.Name, expectedName, StringComparison.Ordinal) Then Return False
        End If

        entity.Group = GetOrCreateGroup(entity.Drawing, expectedName)
        Return True
    End Function

    Private Shared Function ReadExpectedGroupName(entity As DwgEntity) As String
        Dim data As New StructureElement()
        If Not FuncXRecords.getXRecords(entity, data, StructureElement.tableXRecords.PROJECT_STRUCTURES) Then
            Return String.Empty
        End If
        Return GetGroupName(data.IdStructure)
    End Function

    Private Shared Function GetOrCreateGroup(drawing As Drawing, name As String) As DwgGroup
        If drawing.Groups.IsExists(name) Then Return drawing.Groups(name)
        Return drawing.Groups.Add(name)
    End Function

    Private Shared Function IsOwnedGroup(group As DwgGroup) As Boolean
        Return group IsNot Nothing AndAlso IsOwnedGroupName(group.Name)
    End Function

    Private Shared Function IsOwnedGroupName(groupName As String) As Boolean
        If String.IsNullOrEmpty(groupName) OrElse
           Not groupName.StartsWith(GroupPrefix, StringComparison.Ordinal) Then Return False

        Dim suffix As String = groupName.Substring(GroupPrefix.Length)
        If suffix.Length = 32 Then Return IsLowerHex(suffix)

        Const LegacyPrefix As String = "legacy_"
        Return suffix.Length = LegacyPrefix.Length + 64 AndAlso
               suffix.StartsWith(LegacyPrefix, StringComparison.Ordinal) AndAlso
               IsLowerHex(suffix.Substring(LegacyPrefix.Length))
    End Function

    Private Shared Function IsLowerHex(value As String) As Boolean
        For Each item As Char In value
            If Not ((item >= "0"c AndAlso item <= "9"c) OrElse
                    (item >= "a"c AndAlso item <= "f"c)) Then Return False
        Next
        Return value.Length > 0
    End Function

    Private Shared Function HasEmptyOwnedGroups(drawing As Drawing) As Boolean
        For Each group As DwgGroup In drawing.Groups
            If IsOwnedGroup(group) AndAlso group.EntitysCount = 0 Then Return True
        Next
        Return False
    End Function

    Private Shared Sub RemoveEmptyOwnedGroups(drawing As Drawing)
        For index As Integer = drawing.Groups.Count - 1 To 0 Step -1
            Dim group As DwgGroup = drawing.Groups(index)
            If IsOwnedGroup(group) AndAlso group.EntitysCount = 0 Then
                drawing.Groups.Remove(group)
            End If
        Next
    End Sub

    Private Shared Sub OnEntityChanged(sender As Object, e As EventArgs)
        TryGroupEntity(TryCast(sender, DwgEntity))
    End Sub

    Private Shared Sub OnBeforeRemoveEntity(sender As Object, e As EventArgs)
        Dim entity As DwgEntity = TryCast(sender, DwgEntity)
        If entity Is Nothing OrElse entity.Drawing Is Nothing Then Return

        Dim state As DrawingState = Nothing
        If States.TryGetValue(entity.Drawing, state) Then state.NeedsCleanup = True
    End Sub

    Private Shared Sub OnAfterUpdate(sender As Object, e As EventArgs)
        Dim drawing As Drawing = TryCast(sender, Drawing)
        If drawing Is Nothing Then Return

        Dim state As DrawingState = Nothing
        If Not States.TryGetValue(drawing, state) OrElse state.IsHandling OrElse Not state.NeedsCleanup Then Return

        state.IsHandling = True
        Try
            RemoveEmptyOwnedGroups(drawing)
            state.NeedsCleanup = False
        Catch
            ' Cleanup can be retried after the next drawing update.
        Finally
            state.IsHandling = False
        End Try
    End Sub

    Private NotInheritable Class DrawingState
        Public IsHandling As Boolean
        Public NeedsCleanup As Boolean
    End Class

    Private NotInheritable Class GroupAssignment
        Public Sub New(entity As DwgEntity, groupName As String)
            Me.Entity = entity
            Me.GroupName = groupName
        End Sub

        Public ReadOnly Entity As DwgEntity
        Public ReadOnly GroupName As String
    End Class
End Class
