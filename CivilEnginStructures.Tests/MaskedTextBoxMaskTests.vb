Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class MaskedTextBoxMaskTests
        <TestCase("123.123", "1+23.123", "0+00\.000")>
        <TestCase("123312.111", "1233+12.111", "0000+00\.000")>
        Public Sub TryFormatPkTextBuildsMaskForAllStationDigits(input As String,
                                                                 expectedText As String,
                                                                 expectedMask As String)
            Dim formattedText As String = Nothing
            Dim mask As String = Nothing

            Dim success As Boolean = FuncFormatZn.TryFormatPKText(input, formattedText, mask)

            Assert.Multiple(
                Sub()
                    Assert.That(success, [Is].True)
                    Assert.That(formattedText, [Is].EqualTo(expectedText))
                    Assert.That(mask, [Is].EqualTo(expectedMask))
                End Sub)
        End Sub

        <TestCase("1+23.123", 123.123R)>
        <TestCase("1233+12.111", 123312.111R)>
        <SetCulture("ru-RU")>
        Public Sub TryParsePkTextUsesLiteralDotUnderRussianCulture(input As String,
                                                                    expectedValue As Double)
            Dim decimalPK As Double

            Dim success As Boolean = FuncFormatZn.TryParsePKText(input, decimalPK)

            Assert.Multiple(
                Sub()
                    Assert.That(success, [Is].True)
                    Assert.That(decimalPK, [Is].EqualTo(expectedValue).Within(0.0000001R))
                End Sub)
        End Sub

        <Test, SetCulture("ru-RU"), Apartment(ApartmentState.STA)>
        Public Sub DynamicMaskKeepsLiteralDotUnderRussianCulture()
            Dim formattedText As String = Nothing
            Dim mask As String = Nothing

            Dim success As Boolean = FuncFormatZn.TryFormatPKText("123312.111", formattedText, mask)

            Using textBox As New MaskedTextBox()
                textBox.Mask = mask
                textBox.TextMaskFormat = MaskFormat.IncludeLiterals
                textBox.Text = formattedText

                Assert.Multiple(
                    Sub()
                        Assert.That(success, [Is].True)
                        Assert.That(mask, [Is].EqualTo("0000+00\.000"))
                        Assert.That(textBox.Text, [Is].EqualTo("1233+12.111"))
                    End Sub)
            End Using
        End Sub

        <Test, Apartment(ApartmentState.STA)>
        Public Sub ReapplyingDynamicMaskAfterClearingMaskPreservesEveryDigit()
            Dim formattedText As String = Nothing
            Dim mask As String = Nothing
            Assert.That(FuncFormatZn.TryFormatPKText("123312.111", formattedText, mask), [Is].True)

            Using textBox As New MaskedTextBox()
                textBox.Mask = mask
                textBox.TextMaskFormat = MaskFormat.IncludeLiterals
                textBox.Text = formattedText

                Dim displayedText As String = textBox.Text
                Dim rawText As String = displayedText.Replace("+", "")
                textBox.Mask = String.Empty
                textBox.Text = rawText

                Dim reappliedText As String = Nothing
                Dim reappliedMask As String = Nothing
                Dim success As Boolean = FuncFormatZn.TryFormatPKText(textBox.Text, reappliedText, reappliedMask)

                Assert.Multiple(
                    Sub()
                        Assert.That(displayedText, [Is].EqualTo("1233+12.111"))
                        Assert.That(rawText, [Is].EqualTo("123312.111"))
                        Assert.That(success, [Is].True)
                        Assert.That(reappliedText, [Is].EqualTo("1233+12.111"))
                        Assert.That(reappliedMask, [Is].EqualTo("0000+00\.000"))
                    End Sub)
            End Using
        End Sub
    End Class
End Namespace
