Public NotInheritable Class BridgeAppearancePatch
    Public Property CadColorValue As Integer?
    Public Property LayerName As String
    Public Property LinetypeName As String
    Public Property LinetypeScale As Double?
    Public Property Lineweight As Integer?
    Public Property Width As Double?
    Public Property BodyArgb As Integer?
    Public Property FillCadColorValue As Integer?
    Public Property HatchPatternName As String
    Public Property HatchScale As Double?
    Public Property HatchAngle As Double?

    Public Function TryValidate(ByRef reason As String) As Boolean
        reason = Nothing

        If LinetypeScale.HasValue AndAlso
           (Double.IsNaN(LinetypeScale.Value) OrElse
            Double.IsInfinity(LinetypeScale.Value) OrElse
            LinetypeScale.Value <= 0.0) Then
            reason = "Масштаб типа линии должен быть положительным конечным числом."
            Return False
        End If

        If Width.HasValue AndAlso
           (Double.IsNaN(Width.Value) OrElse
            Double.IsInfinity(Width.Value) OrElse
            Width.Value < 0.0) Then
            reason = "Ширина полилинии должна быть неотрицательным конечным числом."
            Return False
        End If

        If HatchPatternName IsNot Nothing Then
            Dim normalized As String = HatchPatternName.Trim().ToUpperInvariant()
            If normalized <> "SOLID" AndAlso normalized <> "ANSI31" Then
                reason = "Поддерживаются штриховки SOLID и ANSI31."
                Return False
            End If
        End If

        If HatchScale.HasValue AndAlso
           (Double.IsNaN(HatchScale.Value) OrElse
            Double.IsInfinity(HatchScale.Value) OrElse
            HatchScale.Value <= 0.0R) Then
            reason = "Масштаб штриховки должен быть положительным конечным числом."
            Return False
        End If

        If HatchAngle.HasValue AndAlso
           (Double.IsNaN(HatchAngle.Value) OrElse Double.IsInfinity(HatchAngle.Value)) Then
            reason = "Угол штриховки должен быть конечным числом."
            Return False
        End If

        Return True
    End Function

    Public Function ApplyTo(source As BridgeAppearanceValues) As BridgeAppearanceValues
        If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))

        Dim reason As String = Nothing
        If Not TryValidate(reason) Then Throw New ArgumentException(reason, NameOf(source))

        Return New BridgeAppearanceValues(
            If(CadColorValue.HasValue, CadColorValue.Value, source.CadColorValue),
            If(LayerName, source.LayerName),
            If(LinetypeName, source.LinetypeName),
            If(LinetypeScale.HasValue, LinetypeScale.Value, source.LinetypeScale),
            If(Lineweight.HasValue, Lineweight.Value, source.Lineweight),
            If(Width.HasValue, Width, source.Width),
            If(BodyArgb.HasValue, BodyArgb, source.BodyArgb),
            If(FillCadColorValue.HasValue, FillCadColorValue, source.FillCadColorValue),
            If(HatchPatternName, source.HatchPatternName),
            If(HatchScale.HasValue, HatchScale, source.HatchScale),
            If(HatchAngle.HasValue, HatchAngle, source.HatchAngle))
    End Function

    Public Function ForTarget(supportsWidth As Boolean,
                              isModel3D As Boolean) As BridgeAppearancePatch
        Return ForTarget(supportsWidth, isModel3D, False)
    End Function

    Public Function ForTarget(supportsWidth As Boolean,
                              isModel3D As Boolean,
                              supportsFill As Boolean) As BridgeAppearancePatch
        Return New BridgeAppearancePatch With {
            .CadColorValue = CadColorValue,
            .LayerName = LayerName,
            .LinetypeName = LinetypeName,
            .LinetypeScale = LinetypeScale,
            .Lineweight = Lineweight,
            .Width = If(supportsWidth, Width, Nothing),
            .BodyArgb = If(isModel3D, BodyArgb, Nothing),
            .FillCadColorValue = If(supportsFill, FillCadColorValue, Nothing),
            .HatchPatternName = If(supportsFill, HatchPatternName, Nothing),
            .HatchScale = If(supportsFill, HatchScale, Nothing),
            .HatchAngle = If(supportsFill, HatchAngle, Nothing)
        }
    End Function

    Friend ReadOnly Property IsEmpty As Boolean
        Get
            Return Not CadColorValue.HasValue AndAlso
                   LayerName Is Nothing AndAlso
                   LinetypeName Is Nothing AndAlso
                   Not LinetypeScale.HasValue AndAlso
                   Not Lineweight.HasValue AndAlso
                   Not Width.HasValue AndAlso
                   Not BodyArgb.HasValue AndAlso
                   Not FillCadColorValue.HasValue AndAlso
                   HatchPatternName Is Nothing AndAlso
                   Not HatchScale.HasValue AndAlso
                   Not HatchAngle.HasValue
        End Get
    End Property
End Class
