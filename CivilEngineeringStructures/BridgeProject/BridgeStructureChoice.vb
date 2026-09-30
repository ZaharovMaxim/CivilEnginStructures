Imports System.Text.RegularExpressions

Public NotInheritable Class BridgeStructureChoice
    Private Const NewStructureCaption As String = "Новое искусственное сооружение"
    Private Const GeneratedNamePrefix As String = "Искусственное сооружение "

    Private Sub New(id As String, displayName As String, isNew As Boolean)
        Me.Id = If(id, String.Empty)
        Me.DisplayName = If(displayName, String.Empty)
        Me.IsNew = isNew
    End Sub

    Public ReadOnly Property Id As String
    Public ReadOnly Property DisplayName As String
    Public ReadOnly Property IsNew As Boolean

    Public Shared Function CreateChoices(names As IEnumerable(Of KeyValuePair(Of String, String)),
                                         newBridgeId As String) As List(Of BridgeStructureChoice)
        Dim source As List(Of KeyValuePair(Of String, String)) = If(names Is Nothing,
                                                                    New List(Of KeyValuePair(Of String, String))(),
                                                                    names.ToList())
        Dim reservedNumbers As New HashSet(Of Integer)()
        For Each item As KeyValuePair(Of String, String) In source
            Dim number As Integer
            If TryGetGeneratedNumber(item.Value, number) Then reservedNumbers.Add(number)
        Next

        Dim result As New List(Of BridgeStructureChoice)()
        Dim nextNumber As Integer = 1
        For Each item As KeyValuePair(Of String, String) In source
            Dim isNew As Boolean = Not String.IsNullOrWhiteSpace(newBridgeId) AndAlso
                                   String.Equals(item.Key, newBridgeId, StringComparison.Ordinal)
            Dim displayName As String = If(item.Value, String.Empty)
            If isNew Then
                displayName = NewStructureCaption
            ElseIf String.IsNullOrWhiteSpace(displayName) OrElse
                   String.Equals(displayName, NewStructureCaption, StringComparison.Ordinal) Then
                While reservedNumbers.Contains(nextNumber)
                    nextNumber += 1
                End While
                displayName = GeneratedNamePrefix & nextNumber.ToString()
                reservedNumbers.Add(nextNumber)
                nextNumber += 1
            End If
            result.Add(New BridgeStructureChoice(item.Key, displayName, isNew))
        Next
        Return result
    End Function

    Public Shared Function ResolveSavedName(inputText As String,
                                            selected As BridgeStructureChoice,
                                            choices As IEnumerable(Of BridgeStructureChoice)) As String
        Dim value As String = If(inputText, String.Empty).Trim()
        If value.Length > 0 AndAlso Not String.Equals(value, NewStructureCaption, StringComparison.Ordinal) Then Return value
        If selected IsNot Nothing AndAlso Not selected.IsNew Then Return selected.DisplayName

        Dim reservedNumbers As New HashSet(Of Integer)()
        If choices IsNot Nothing Then
            For Each choice As BridgeStructureChoice In choices
                If choice Is Nothing OrElse choice.IsNew Then Continue For
                Dim number As Integer
                If TryGetGeneratedNumber(choice.DisplayName, number) Then reservedNumbers.Add(number)
            Next
        End If
        Dim nextNumber As Integer = 1
        While reservedNumbers.Contains(nextNumber)
            nextNumber += 1
        End While
        Return GeneratedNamePrefix & nextNumber.ToString()
    End Function

    Private Shared Function TryGetGeneratedNumber(value As String, ByRef number As Integer) As Boolean
        If value Is Nothing Then Return False
        Dim match As Match = Regex.Match(value,
                                         "^" & Regex.Escape(GeneratedNamePrefix) & "([1-9][0-9]*)$",
                                         RegexOptions.CultureInvariant Or RegexOptions.IgnoreCase)
        Return match.Success AndAlso Integer.TryParse(match.Groups(1).Value, number)
    End Function

    Public Overrides Function ToString() As String
        Return DisplayName
    End Function
End Class
