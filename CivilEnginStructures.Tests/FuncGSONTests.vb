Imports System
Imports System.Collections.Generic
Imports NUnit.Framework

Namespace Tests
    <TestFixture>
    Public Class FuncGSONTests
        <Test>
        Public Sub ParseJsonToDictionaryReturnsSupportedPrimitiveTypes()
            Dim json As String =
                "{""text"":""bridge""," &
                """integer"":42," &
                """number"":1.25," &
                """enabled"":true," &
                """created"":""2026-08-01T12:34:56Z""}"

            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary(json)

            Assert.Multiple(
                Sub()
                    Assert.That(result, Has.Count.EqualTo(5))
                    Assert.That(result("text"), [Is].TypeOf(Of String)())
                    Assert.That(result("text"), [Is].EqualTo("bridge"))
                    Assert.That(result("integer"), [Is].TypeOf(Of Integer)())
                    Assert.That(result("integer"), [Is].EqualTo(42))
                    Assert.That(result("number"), [Is].TypeOf(Of Double)())
                    Assert.That(result("number"), [Is].EqualTo(1.25R))
                    Assert.That(result("enabled"), [Is].TypeOf(Of Boolean)())
                    Assert.That(result("enabled"), [Is].True)
                    Assert.That(result("created"), [Is].TypeOf(Of DateTime)())
                    Assert.That(
                        DirectCast(result("created"), DateTime),
                        [Is].EqualTo(New DateTime(2026, 8, 1, 12, 34, 56, DateTimeKind.Utc)))
                End Sub)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryKeepsNumericAndTimeLikeStringsAsStrings()
            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary("{""numeric"":""1.2"",""time"":""12:30""}")

            Assert.Multiple(
                Sub()
                    Assert.That(result("numeric"), [Is].TypeOf(Of String)())
                    Assert.That(result("numeric"), [Is].EqualTo("1.2"))
                    Assert.That(result("time"), [Is].TypeOf(Of String)())
                    Assert.That(result("time"), [Is].EqualTo("12:30"))
                End Sub)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryKeepsInt32BoundaryValuesAsIntegers()
            Dim json As String =
                "{""minimum"":" & Integer.MinValue.ToString(Globalization.CultureInfo.InvariantCulture) &
                ",""maximum"":" & Integer.MaxValue.ToString(Globalization.CultureInfo.InvariantCulture) & "}"

            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary(json)

            Assert.Multiple(
                Sub()
                    Assert.That(result("minimum"), [Is].TypeOf(Of Integer)())
                    Assert.That(result("minimum"), [Is].EqualTo(Integer.MinValue))
                    Assert.That(result("maximum"), [Is].TypeOf(Of Integer)())
                    Assert.That(result("maximum"), [Is].EqualTo(Integer.MaxValue))
                End Sub)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionarySkipsIntegersOutsideInt32Range()
            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary(
                    "{""below"":-2147483649,""above"":2147483648,""valid"":7}")

            Assert.Multiple(
                Sub()
                    Assert.That(result.ContainsKey("below"), [Is].False)
                    Assert.That(result.ContainsKey("above"), [Is].False)
                    Assert.That(result("valid"), [Is].TypeOf(Of Integer)())
                    Assert.That(result("valid"), [Is].EqualTo(7))
                End Sub)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryIgnoresNestedObjectsArraysAndNulls()
            Dim json As String =
                "{""name"":""root""," &
                """nested"":{""child"":1}," &
                """items"":[1,2,3]," &
                """missing"":null}"

            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary(json)

            Assert.Multiple(
                Sub()
                    Assert.That(result, Has.Count.EqualTo(1))
                    Assert.That(result("name"), [Is].EqualTo("root"))
                    Assert.That(result.ContainsKey("child"), [Is].False,
                        "Nested properties must not be flattened into the result.")
                    Assert.That(result.ContainsKey("nested"), [Is].False)
                    Assert.That(result.ContainsKey("items"), [Is].False)
                    Assert.That(result.ContainsKey("missing"), [Is].False)
                End Sub)
        End Sub

        <TestCase("")>
        <TestCase("   ")>
        Public Sub ParseJsonToDictionaryReturnsEmptyDictionaryForBlankInput(json As String)
            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary(json)

            Assert.That(result, [Is].Empty)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryReturnsEmptyDictionaryForEmptyObject()
            Dim result As Dictionary(Of String, Object) =
                FuncGSON.ParseJsonToDictionary("{}")

            Assert.That(result, [Is].Empty)
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryThrowsArgumentExceptionForMalformedJson()
            Assert.Throws(Of ArgumentException)(
                Sub() FuncGSON.ParseJsonToDictionary("{""value"":1"))
        End Sub

        <Test>
        Public Sub ParseJsonToDictionaryThrowsArgumentExceptionForRootArray()
            Assert.Throws(Of ArgumentException)(
                Sub() FuncGSON.ParseJsonToDictionary("[1,2,3]"))
        End Sub

        <Test>
        Public Sub SetTypedValueUpdatesExistingPrimitiveFieldsAndPreservesTheirTokenTypes()
            Dim json As String =
                "{""Text"":""old""," &
                """Integer"":1," &
                """Double"":1.5," &
                """Boolean"":false," &
                """Date"":""2026-08-01T12:34:56Z""}"
            Dim originalCulture As Globalization.CultureInfo =
                Threading.Thread.CurrentThread.CurrentCulture

            Try
                Threading.Thread.CurrentThread.CurrentCulture =
                    Globalization.CultureInfo.GetCultureInfo("ru-RU")

                Dim textUpdated As Boolean = FuncGSON.setTypedValue(json, "text", "12:30")
                Dim integerUpdated As Boolean = FuncGSON.setTypedValue(json, "INTEGER", "-42")
                Dim doubleUpdated As Boolean = FuncGSON.setTypedValue(json, "double", "1234.5")
                Dim booleanUpdated As Boolean = FuncGSON.setTypedValue(json, "BOOLEAN", "True")
                Dim dateUpdated As Boolean =
                    FuncGSON.setTypedValue(json, "date", "2027-02-03T04:05:06Z")
                Dim parsed As Newtonsoft.Json.Linq.JObject =
                    Newtonsoft.Json.Linq.JObject.Parse(json)

                Assert.Multiple(
                    Sub()
                        Assert.That(textUpdated, [Is].True)
                        Assert.That(integerUpdated, [Is].True)
                        Assert.That(doubleUpdated, [Is].True)
                        Assert.That(booleanUpdated, [Is].True)
                        Assert.That(dateUpdated, [Is].True)
                        Assert.That(parsed.Properties().Count(), [Is].EqualTo(5),
                            "Case-insensitive updates must not add new properties.")
                        Assert.That(parsed("Text").Type,
                            [Is].EqualTo(Newtonsoft.Json.Linq.JTokenType.String))
                        Assert.That(parsed("Text").ToObject(Of String)(), [Is].EqualTo("12:30"))
                        Assert.That(parsed("Integer").Type,
                            [Is].EqualTo(Newtonsoft.Json.Linq.JTokenType.Integer))
                        Assert.That(parsed("Integer").ToObject(Of Integer)(), [Is].EqualTo(-42))
                        Assert.That(parsed("Double").Type,
                            [Is].EqualTo(Newtonsoft.Json.Linq.JTokenType.Float))
                        Assert.That(parsed("Double").ToObject(Of Double)(), [Is].EqualTo(1234.5R))
                        Assert.That(parsed("Boolean").Type,
                            [Is].EqualTo(Newtonsoft.Json.Linq.JTokenType.Boolean))
                        Assert.That(parsed("Boolean").ToObject(Of Boolean)(), [Is].True)
                        Assert.That(parsed("Date").Type,
                            [Is].EqualTo(Newtonsoft.Json.Linq.JTokenType.Date))
                        Assert.That(
                            parsed("Date").ToObject(Of DateTime)(),
                            [Is].EqualTo(New DateTime(2027, 2, 3, 4, 5, 6, DateTimeKind.Utc)))
                    End Sub)
            Finally
                Threading.Thread.CurrentThread.CurrentCulture = originalCulture
            End Try
        End Sub

        <Test>
        Public Sub SetTypedValueRejectsInvalidOrUnsupportedValuesWithoutChangingJson()
            Dim originalJson As String =
                "{""String"":""old""," &
                """Integer"":1," &
                """Double"":1.5," &
                """Boolean"":false," &
                """Date"":""2026-08-01T12:34:56Z""," &
                """Nested"":{""child"":1}," &
                """Items"":[1]," &
                """NullValue"":null}"
            Dim attempts As String(,) =
                {
                    {"Integer", "not-an-integer"},
                    {"Double", "not-a-double"},
                    {"Boolean", "yes"},
                    {"Date", "not-a-date"},
                    {"Nested", "replacement"},
                    {"Items", "replacement"},
                    {"NullValue", "replacement"},
                    {"Missing", "replacement"}
                }

            Assert.Multiple(
                Sub()
                    For index As Integer = 0 To attempts.GetLength(0) - 1
                        Dim candidate As String = originalJson
                        Dim updated As Boolean =
                            FuncGSON.setTypedValue(candidate, attempts(index, 0), attempts(index, 1))

                        Assert.That(updated, [Is].False,
                            $"Field '{attempts(index, 0)}' must be rejected.")
                        Assert.That(candidate, [Is].EqualTo(originalJson),
                            $"A rejected update of '{attempts(index, 0)}' must be atomic.")
                    Next
                End Sub)
        End Sub
    End Class
End Namespace
