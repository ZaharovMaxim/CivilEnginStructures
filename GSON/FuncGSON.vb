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
End Class
