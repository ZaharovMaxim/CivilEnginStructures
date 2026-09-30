Public NotInheritable Class SurfaceModelChoice
    Private Shared ReadOnly SupportedExtensions As String() = {".sfcx", ".roadx", ".site", ".algx"}

    Public Sub New(modelName As String, modelType As String, relativePath As String)
        Me.ModelName = If(modelName, String.Empty)
        Me.ModelType = If(modelType, String.Empty)
        Me.RelativePath = If(relativePath, String.Empty)
    End Sub

    Public ReadOnly Property ModelName As String
    Public ReadOnly Property ModelType As String
    Public ReadOnly Property RelativePath As String

    Public ReadOnly Property DisplayName As String
        Get
            If String.IsNullOrWhiteSpace(ModelType) AndAlso String.IsNullOrWhiteSpace(RelativePath) Then Return ModelName
            If String.IsNullOrWhiteSpace(ModelType) Then Return ModelName & " — " & RelativePath
            Return ModelName & " [" & ModelType & "] — " & RelativePath
        End Get
    End Property

    Public Shared Function FindByStoredValue(choices As IEnumerable(Of SurfaceModelChoice), storedValue As String) As SurfaceModelChoice
        If choices Is Nothing OrElse String.IsNullOrWhiteSpace(storedValue) Then Return Nothing
        For Each choice As SurfaceModelChoice In choices
            If String.Equals(choice.RelativePath, storedValue, StringComparison.OrdinalIgnoreCase) Then Return choice
        Next
        For Each choice As SurfaceModelChoice In choices
            If String.Equals(choice.ModelName, storedValue, StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(choice.DisplayName, storedValue, StringComparison.Ordinal) Then Return choice
        Next
        Return Nothing
    End Function

    Public Shared Function IsSupportedSurfacePath(path As String) As Boolean
        If String.IsNullOrWhiteSpace(path) Then Return False
        Dim extension As String = IO.Path.GetExtension(path)
        Return SupportedExtensions.Any(Function(item) String.Equals(item, extension, StringComparison.OrdinalIgnoreCase))
    End Function

    Public Overrides Function ToString() As String
        Return DisplayName
    End Function
End Class
