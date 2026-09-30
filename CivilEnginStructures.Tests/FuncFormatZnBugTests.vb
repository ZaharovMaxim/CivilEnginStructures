Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    <SetCulture("en-US")>
    Public Class FuncFormatZnBugTests
        <Test>
        Public Sub FuncFormatPkHonorsRequestedDecimalPlaces()
            Dim result As String = FuncFormatZn.FuncFormatPK(123.456, 3)

            Assert.That(result, [Is].EqualTo("1+23.456"))
        End Sub

        <Test>
        Public Sub FuncFormatPkDefaultPrecisionRemainsTwoDecimals()
            Dim result As String = FuncFormatZn.FuncFormatPK(123.456)

            Assert.That(result, [Is].EqualTo("1+23.46"))
        End Sub

        <Test>
        Public Sub FuncFormatKmHonorsRequestedDecimalPlaces()
            Dim result As String = FuncFormatZn.FuncFormatKM(1234.567, 3)

            Assert.That(result, [Is].EqualTo("1+234.567"))
        End Sub

        <Test>
        Public Sub FuncFormatPkNormalizesRoundedHundredToNextStation()
            Dim result As String = FuncFormatZn.FuncFormatPK(99.999)

            Assert.That(result, [Is].EqualTo("1+00.00"))
        End Sub

        <Test>
        Public Sub FuncUnicodDecodeDecodesUnicodeFullStop()
            Dim result As String = FuncFormatZn.FuncUnicodDecode("\u002E")

            Assert.That(result, [Is].EqualTo("."))
        End Sub

        <Test>
        Public Sub FuncFormatDegToDmmssPreservesNegativeSign()
            Dim positive As String = FuncFormatZn.FuncFormatDEGToDMMSS(12.5)
            Dim negative As String = FuncFormatZn.FuncFormatDEGToDMMSS(-12.5)

            Assert.That(negative, [Is].EqualTo("-" & positive))
        End Sub
    End Class
End Namespace
