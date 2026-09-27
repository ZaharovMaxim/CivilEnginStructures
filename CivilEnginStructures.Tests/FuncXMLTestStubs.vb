' Test-only DTO required to compile the linked production FuncXML.vb in isolation.
Public Class BeamI
    Public Property model As String
    Public Property lenght As Double
    Public Property widthTopPlateLeft As Double
    Public Property widthTopPlateRight As Double
    Public Property heightTopPlate As Double
    Public Property WidthTop As Double
    Public Property widthBottom As Double
    Public Property height As Double
    Public Property a As Double
    Public Property b As Double
    Public Property VerticalOffsetRibZone As Double
    Public Property HorizontalOffsetRibZone As Double
    Public Property DeltaRib As Double
    Public Property radiusTop As Double
    Public Property radiusBottom As Double
    Public Property gWidth As Double
    Public Property mass As Double
    Public Property volume As Double
    Public Property modelTLS As String
    Public Property nameAlbum As String
End Class

Namespace Global.Microsoft.Office.Interop.Excel
    ' Compile-only stubs for the unrelated Excel import/export method in FuncXML.vb.
    Public Class Application
        Public ReadOnly Property Workbooks As New Workbooks

        Public Sub Quit()
        End Sub
    End Class

    Public Class Workbooks
        Public Function Open(
            fileName As String,
            Optional updateLinks As Object = Nothing,
            Optional [readOnly] As Object = Nothing) As Workbook

            Return New Workbook()
        End Function
    End Class

    Public Class Workbook
        Public ReadOnly Property Worksheets As New Worksheets

        Public Sub Close(Optional saveChanges As Object = Nothing)
        End Sub
    End Class

    Public Class Worksheets
        Default Public ReadOnly Property Item(index As Object) As Worksheet
            Get
                Return New Worksheet()
            End Get
        End Property
    End Class

    Public Class Worksheet
        Public ReadOnly Property Rows As New Rows
        Public ReadOnly Property Cells As New Cells
    End Class

    Public Class Rows
        Public ReadOnly Property Count As Integer
            Get
                Return 0
            End Get
        End Property
    End Class

    Public Class Cells
        Default Public ReadOnly Property Item(row As Integer, column As Integer) As Cell
            Get
                Return New Cell()
            End Get
        End Property
    End Class

    Public Class Cell
        Public Property value As Object
    End Class
End Namespace
