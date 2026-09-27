Imports Newtonsoft.Json.Linq

Public Class FuncGSON
    Public Shared Function IsValidJson(jsonString As String) As Boolean
        If String.IsNullOrWhiteSpace(jsonString) Then
            Return False
        End If
        Try
            JToken.Parse(jsonString)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function getValue(ByVal strGSON As String, ByVal nameField As String) As String
        Dim result As String = ""
        If String.IsNullOrEmpty(strGSON) OrElse String.IsNullOrEmpty(nameField) Then
            Return result
        End If
        Try
            If IsValidJson(strGSON) Then
                Dim jsonObj As JObject = JObject.Parse(strGSON)
                ' Ищем ключ без учета регистра
                Dim foundKey As String = jsonObj.Properties() _
                .Where(Function(p) String.Equals(p.Name, nameField, StringComparison.OrdinalIgnoreCase)) _
                .Select(Function(p) p.Name) _
                .FirstOrDefault()
                If foundKey IsNot Nothing Then
                    Dim value As JToken = jsonObj(foundKey)
                    If value IsNot Nothing AndAlso value.Type <> JTokenType.Null Then
                        result = value.ToString()
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
        Return result
    End Function

    Public Shared Function setValue(ByRef strGSON As String, ByVal nameField As String, ByVal valueToSet As String) As Boolean
        ' Проверка входных параметров
        If String.IsNullOrEmpty(strGSON) OrElse String.IsNullOrEmpty(nameField) Then
            Return False
        End If

        Try
            If IsValidJson(strGSON) Then
                Dim jsonObj As JObject = JObject.Parse(strGSON)

                ' Ищем ключ без учета регистра
                Dim foundKey As String = jsonObj.Properties() _
                    .Where(Function(p) String.Equals(p.Name, nameField, StringComparison.OrdinalIgnoreCase)) _
                    .Select(Function(p) p.Name) _
                    .FirstOrDefault()

                If foundKey IsNot Nothing Then
                    ' Обновляем существующее значение
                    jsonObj(foundKey) = valueToSet
                Else
                    ' Добавляем новое поле (сохраняем оригинальное имя)
                    jsonObj.Add(nameField, valueToSet)
                End If

                ' Обновляем строку JSON
                strGSON = jsonObj.ToString()
                Return True
            End If
        Catch ex As Exception
            ' Логирование ошибки при необходимости
            ' Logger.Error($"Ошибка при установке значения: {ex.Message}")
            Return False
        End Try

        Return False
    End Function

    Public Shared Function setTypedValue(ByRef strGSON As String, nameField As String, valueToSet As String) As Boolean
        If String.IsNullOrWhiteSpace(strGSON) OrElse String.IsNullOrEmpty(nameField) OrElse valueToSet Is Nothing Then
            Return False
        End If

        Try
            Dim jsonObj As JObject = JObject.Parse(strGSON)
            Dim foundProperty As JProperty = jsonObj.Properties() _
                .FirstOrDefault(Function(p) String.Equals(p.Name, nameField, StringComparison.OrdinalIgnoreCase))
            If foundProperty Is Nothing Then Return False

            Dim convertedValue As JToken = Nothing
            Select Case foundProperty.Value.Type
                Case JTokenType.String
                    convertedValue = New JValue(valueToSet)
                Case JTokenType.Integer
                    Dim integerValue As Integer
                    If Not Integer.TryParse(valueToSet, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, integerValue) Then
                        Return False
                    End If
                    convertedValue = New JValue(integerValue)
                Case JTokenType.Float
                    Dim doubleValue As Double
                    If Not Double.TryParse(valueToSet, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, doubleValue) _
                        OrElse Double.IsNaN(doubleValue) OrElse Double.IsInfinity(doubleValue) Then
                        Return False
                    End If
                    convertedValue = New JValue(doubleValue)
                Case JTokenType.Boolean
                    Dim booleanValue As Boolean
                    If Not Boolean.TryParse(valueToSet, booleanValue) Then Return False
                    convertedValue = New JValue(booleanValue)
                Case JTokenType.Date
                    Dim dateValue As DateTime
                    If Not DateTime.TryParse(valueToSet, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.RoundtripKind, dateValue) Then
                        Return False
                    End If
                    convertedValue = New JValue(dateValue)
                Case Else
                    Return False
            End Select

            foundProperty.Value = convertedValue
            Dim updatedJson As String = jsonObj.ToString()
            strGSON = updatedJson
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function ParseJsonToDictionary(json As String) As Dictionary(Of String, Object)
        Dim result As New Dictionary(Of String, Object)()
        If String.IsNullOrWhiteSpace(json) Then Return result
        Try
            Dim jObject As JObject = JObject.Parse(json)
            For Each prop As JProperty In jObject.Properties()
                Dim key = prop.Name
                Dim token = prop.Value

                Select Case token.Type
                    Case JTokenType.String
                        Dim str = token.ToString()
                        result.Add(key, str)
                    Case JTokenType.Integer
                        Dim intValue As Integer
                        If Integer.TryParse(token.ToString(), Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, intValue) Then
                            result.Add(key, intValue)
                        End If
                    Case JTokenType.Float
                        Dim d As Double = token.Value(Of Double)()
                        result.Add(key, d)
                    Case JTokenType.Boolean
                        Dim b As Boolean = token.Value(Of Boolean)()
                        result.Add(key, b)
                    Case JTokenType.Date
                        Dim d As DateTime = token.Value(Of DateTime)()
                        result.Add(key, d)
                    Case Else
                        ' Игнорируем объекты, массивы, null
                        Continue For
                End Select
            Next
        Catch ex As Exception
            Throw New ArgumentException("Неверный JSON", ex)
        End Try

        Return result
    End Function

End Class
