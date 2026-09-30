Public NotInheritable Class BridgeAppearanceValues
    Public Sub New(cadColorValue As Integer,
                   layerName As String,
                   linetypeName As String,
                   linetypeScale As Double,
                   lineweight As Integer,
                   width As Double?,
                   bodyArgb As Integer?)
        Me.New(cadColorValue, layerName, linetypeName, linetypeScale, lineweight,
               width, bodyArgb, Nothing, Nothing, Nothing, Nothing)
    End Sub

    Public Sub New(cadColorValue As Integer,
                   layerName As String,
                   linetypeName As String,
                   linetypeScale As Double,
                   lineweight As Integer,
                   width As Double?,
                   bodyArgb As Integer?,
                   fillCadColorValue As Integer?,
                   hatchPatternName As String,
                   hatchScale As Double?,
                   hatchAngle As Double?)
        Me.CadColorValue = cadColorValue
        Me.LayerName = layerName
        Me.LinetypeName = linetypeName
        Me.LinetypeScale = linetypeScale
        Me.Lineweight = lineweight
        Me.Width = width
        Me.BodyArgb = bodyArgb
        Me.FillCadColorValue = fillCadColorValue
        Me.HatchPatternName = hatchPatternName
        Me.HatchScale = hatchScale
        Me.HatchAngle = hatchAngle
    End Sub

    Public ReadOnly Property CadColorValue As Integer
    Public ReadOnly Property LayerName As String
    Public ReadOnly Property LinetypeName As String
    Public ReadOnly Property LinetypeScale As Double
    Public ReadOnly Property Lineweight As Integer
    Public ReadOnly Property Width As Double?
    Public ReadOnly Property BodyArgb As Integer?
    Public ReadOnly Property FillCadColorValue As Integer?
    Public ReadOnly Property HatchPatternName As String
    Public ReadOnly Property HatchScale As Double?
    Public ReadOnly Property HatchAngle As Double?
End Class
