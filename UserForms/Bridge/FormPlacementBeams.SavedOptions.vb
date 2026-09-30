Imports System.Globalization
Imports System.Windows.Forms

Partial Public Class FormPlacementBeams
    Friend Sub RestoreSavedPlacementOptions(typeBridge As Integer,
                                            transverseOffset As Double,
                                            verticalOffset As Double,
                                            startPlacementPosition As Double,
                                            spanCount As Integer,
                                            leftRowsCount As Integer,
                                            rightRowsCount As Integer,
                                            leftWidth As Double,
                                            rightWidth As Double)
        If CBox_ListPlacementBeams.Items.Count > 0 Then
            CBox_ListPlacementBeams.SelectedIndex = Math.Max(0, Math.Min(typeBridge, CBox_ListPlacementBeams.Items.Count - 1))
        End If

        NUpD_VerticalOffset.Minimum = -999999999999D
        NUpD_VerticalOffset.Maximum = 999999999999D
        SetNumericValue(NUpD_TraverseOffset, transverseOffset * 1000.0R)
        SetNumericValue(NUpD_VerticalOffset, verticalOffset * 1000.0R)
        SetNumericValue(NUpD_CountProlet, spanCount)
        SetNumericValue(NUpD_CountLeftRows, leftRowsCount)
        SetNumericValue(NUpD_CountRightRows, rightRowsCount)
        SetNumericValue(NUpD_dimLeftBridge, leftWidth * 1000.0R)
        SetNumericValue(NUpD_dimRightBridge, rightWidth * 1000.0R)

        Dim formattedText As String = Nothing
        Dim dynamicMask As String = Nothing
        Dim hasStartPosition As Boolean = startPlacementPosition <> 0.0R AndAlso
            Not Double.IsNaN(startPlacementPosition) AndAlso
            Not Double.IsInfinity(startPlacementPosition) AndAlso
            FuncFormatZn.TryFormatPKText(startPlacementPosition.ToString("0.000", CultureInfo.InvariantCulture),
                                         formattedText,
                                         dynamicMask)
        If Not hasStartPosition Then
            FuncFormatZn.TryFormatPKText("0.000", formattedText, dynamicMask)
        End If

        MaskTB_PK.Mask = String.Empty
        MaskTB_PK.Mask = dynamicMask
        MaskTB_PK.TextMaskFormat = MaskFormat.IncludeLiterals
        MaskTB_PK.Text = formattedText
        CheckBox3.Checked = hasStartPosition
        MaskTB_PK.Enabled = hasStartPosition
    End Sub

    Friend Sub RestoreSavedStartStation(value As Double)
        Dim formattedText As String = Nothing
        Dim dynamicMask As String = Nothing
        Dim hasStartPosition As Boolean = Not Double.IsNaN(value) AndAlso
            Not Double.IsInfinity(value) AndAlso
            FuncFormatZn.TryFormatPKText(value.ToString("0.000", CultureInfo.InvariantCulture),
                                         formattedText,
                                         dynamicMask)
        If Not hasStartPosition Then
            FuncFormatZn.TryFormatPKText("0.000", formattedText, dynamicMask)
        End If

        MaskTB_PK.Mask = String.Empty
        MaskTB_PK.Mask = dynamicMask
        MaskTB_PK.TextMaskFormat = MaskFormat.IncludeLiterals
        MaskTB_PK.Text = formattedText
        CheckBox3.Checked = hasStartPosition
        MaskTB_PK.Enabled = hasStartPosition
    End Sub

    Friend Sub RestoreSavedLongitudinalOffset(value As Double)
        SetNumericValue(NUpD_LongitudinalOffset, value * 1000.0R)
    End Sub

    Friend Shared Function GetSavedRowValues(axisOffset As Double,
                                              offsetSurface As Double,
                                              transverseOffset As Double,
                                              verticalOffset As Double) As Double()
        Return {
            Math.Round(Math.Abs(axisOffset - transverseOffset) * 1000.0R, 3),
            Math.Round((offsetSurface - verticalOffset) * 1000.0R, 3)
        }
    End Function

    Private Shared Sub SetNumericValue(control As NumericUpDown, value As Double)
        If Double.IsNaN(value) Then value = 0.0R
        If Double.IsPositiveInfinity(value) OrElse value > Decimal.ToDouble(control.Maximum) Then
            control.Value = control.Maximum
        ElseIf Double.IsNegativeInfinity(value) OrElse value < Decimal.ToDouble(control.Minimum) Then
            control.Value = control.Minimum
        Else
            control.Value = CDec(value)
        End If
    End Sub
End Class
