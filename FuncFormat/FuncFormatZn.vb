Imports System.Globalization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Topomatic.Cad.Foundation

Public Class FuncFormatZn
    '==================================================================================================================================
    'функция форматирует длины линий под определенный формат
    Public Shared Function FuncFormatCoordinate(ByVal valueZn As Double, Optional ByVal RountZn As Integer = 2) As String
        FuncFormatCoordinate = ""
        Try
            If RountZn = 0 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0}", Math.Round(valueZn, RountZn))
            ElseIf RountZn = 1 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Round(valueZn, RountZn))
            ElseIf RountZn = 2 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Round(valueZn, RountZn))
            ElseIf RountZn = 3 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Round(valueZn, RountZn))
            ElseIf RountZn = 4 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Round(valueZn, RountZn))
            Else
                FuncFormatCoordinate = Math.Round(valueZn, RountZn)
            End If
        Catch
            Exit Function
        End Try
    End Function
    '==================================================================================================================================
    'функция форматирует десятичное число в ПК+(с учетом количества знаков после запятой
    Public Shared Function FuncFormatPK(ByVal DecimalPK As Double, Optional ByVal lenStr As Integer = 2) As String
        FuncFormatPK = "0+00.00"
        Try
            'отделяем число полных пикетов
            Dim IntDist As Integer = Int(DecimalPK / 100)
            'отделяем плюсовую часть
            Dim PKPlus As Double = DecimalPK - IntDist * 100
            If PKPlus = 100 Then
                PKPlus = 0
                IntDist = IntDist + 1
            End If
            'форматируем число (целое число пикетов)
            Dim StrPKPlus As String = String.Format("{0:00.00}", PKPlus)
            FuncFormatPK = IntDist & "+" & StrPKPlus
        Catch
            Exit Function
        End Try
    End Function
    '==================================================================================================================================
    'функция форматирует десятичное число в километр+
    Public Shared Function FuncFormatKM(ByVal DecimalKM As Double, Optional ByVal lenStr As Integer = 2) As String
        FuncFormatKM = "0+000.00"
        Try
            'отделяем число полных километров
            Dim IntDist As Integer = Int(DecimalKM / 1000)
            'отделяем плюсовую часть
            Dim KMPlus As Double = Math.Round(DecimalKM - IntDist * 1000, 2)
            If KMPlus = 1000 Then
                KMPlus = 0
                IntDist = IntDist + 1
            End If
            'преобразуем число в строку
            Dim StrKMPlus As String = String.Format("{0:000.00}", KMPlus)
            FuncFormatKM = IntDist & "+" & StrKMPlus
        Catch
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция преобразут ввденные пользователем значения градусов в десятичном формате в градусы мин сек (простое форматирование)
    Public Shared Function FuncFormatDEGToDMMSS(ByVal Gragus As Double) As String
        FuncFormatDEGToDMMSS = "0.0000"
        Try
            Dim znk As String = ""
            If Gragus < 0 Then
                znk = "-"
                Gragus = Math.Abs(Gragus)
            End If
            Dim Gr As Double = Int(Gragus)
            Dim dr As Double = Math.Round(Gragus - Gr, 4)
            Dim dr_str As String = dr.ToString
            Dim lenStr As Integer = dr_str.Trim.Length
            If lenStr = 1 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & ".0000"
            ElseIf lenStr = 3 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "000"
            ElseIf lenStr = 4 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "00"
            ElseIf lenStr = 5 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "0"
            Else
                FuncFormatDEGToDMMSS = Gragus.ToString
            End If
        Catch
            FuncFormatDEGToDMMSS = "0.0000"
            Exit Function
        End Try
    End Function

    Public Shared Function FuncConvertStrNumberPoint(ByVal xStr As String) As String
        Dim ii, i1, i2, i3, i4, aa, aaa, oStr
        aa = Split(xStr, ",") 'разбиваем строку на массив подстрок по разделителю - запятая
        aaa = Split(aa(0), "-") ' разбиваем начальный элемент этого массива подстрок по разделителю - тире. 
        i1 = CInt(aaa(0)) : i2 = CInt(aaa(UBound(aaa))) 'первое число -> i1, последнее (если оно одно, то первое) -> i2
        oStr = CStr(i1) ' запишем в результирующую строку пока самое первое число
        For ii = 1 To UBound(aa) Step 1 ' пробежимся по остальным элементам массива подстрок
            aaa = Split(aa(ii), "-") ' каждый разбиваем на массив полуподстрок - получаем 1 или 2 элемента
            i3 = CInt(aaa(0)) : i4 = CInt(aaa(UBound(aaa))) 'первое число -> i3, последнее (если оно одно - оно же и первое) -> i4
            If i3 = i2 + 1 Then ' Если это смежные диапазоны чисел 
                i2 = i4         ' - сольём их в один
            Else              ' в противном случае дозаписываем хвост предыдущего диапазона
                If i2 = i1 + 1 Then oStr = oStr & "," & CStr(i2) 'если диапазон всего из 2 чисел
                If i2 > i1 + 1 Then oStr = oStr & "-" & CStr(i2) 'если диапазон поболее
                i1 = i3 : i2 = i4 'текущий диапазон становится прежним
                oStr = oStr & "," & CStr(i1) 'Дозаписываем начало этого диапазона в результирующую строку 
            End If
        Next
        If i2 = i1 + 1 Then oStr = oStr & "," & CStr(i2) 'дозаписываем хвост последнего диапазаноа
        If i2 > i1 + 1 Then oStr = oStr & "-" & CStr(i2)
        FuncConvertStrNumberPoint = oStr ' результат возвращаем
    End Function

    '================================================================================================================================
    'функция декодирует символы Unicode
    Public Shared Function FuncUnicodDecode(ByVal str As String) As String
        Dim Tempstr As String() = str.Split("\")
        Dim FinalStr As String = ""
        For i As Integer = 0 To Tempstr.Length - 1
            Dim tstr As String = Tempstr(i)
            If Left(tstr, 2) Like "u0" Then
                Dim tstr2 As String = Mid(tstr, 2, 4)
                Select Case tstr2.ToUpper
                    Case "0410"
                        FinalStr = FinalStr & "А"
                    Case "0411"
                        FinalStr = FinalStr & "Б"
                    Case "0412"
                        FinalStr = FinalStr & "В"
                    Case "0413"
                        FinalStr = FinalStr & "Г"
                    Case "0414"
                        FinalStr = FinalStr & "Д"
                    Case "0415"
                        FinalStr = FinalStr & "Е"
                    Case "0416"
                        FinalStr = FinalStr & "Ж"
                    Case "0417"
                        FinalStr = FinalStr & "З"
                    Case "0418"
                        FinalStr = FinalStr & "И"
                    Case "0419"
                        FinalStr = FinalStr & "Й"
                    Case "041A"
                        FinalStr = FinalStr & "К"
                    Case "041B"
                        FinalStr = FinalStr & "Л"
                    Case "041C"
                        FinalStr = FinalStr & "М"
                    Case "041D"
                        FinalStr = FinalStr & "Н"
                    Case "041E"
                        FinalStr = FinalStr & "О"
                    Case "041F"
                        FinalStr = FinalStr & "П"
                    Case "0420"
                        FinalStr = FinalStr & "Р"
                    Case "0421"
                        FinalStr = FinalStr & "С"
                    Case "0422"
                        FinalStr = FinalStr & "Т"
                    Case "0423"
                        FinalStr = FinalStr & "У"
                    Case "0424"
                        FinalStr = FinalStr & "Ф"
                    Case "0425"
                        FinalStr = FinalStr & "Х"
                    Case "0426"
                        FinalStr = FinalStr & "Ц"
                    Case "0427"
                        FinalStr = FinalStr & "Ч"
                    Case "0428"
                        FinalStr = FinalStr & "Ш"
                    Case "0429"
                        FinalStr = FinalStr & "Щ"
                    Case "042A"
                        FinalStr = FinalStr & "Ъ"
                    Case "042B"
                        FinalStr = FinalStr & "Ы"
                    Case "042C"
                        FinalStr = FinalStr & "Ь"
                    Case "042D"
                        FinalStr = FinalStr & "Э"
                    Case "042E"
                        FinalStr = FinalStr & "Ю"
                    Case "042F"
                        FinalStr = FinalStr & "Я"

                    Case "0430"
                        FinalStr = FinalStr & "а"
                    Case "0431"
                        FinalStr = FinalStr & "б"
                    Case "0432"
                        FinalStr = FinalStr & "в"
                    Case "0433"
                        FinalStr = FinalStr & "г"
                    Case "0434"
                        FinalStr = FinalStr & "д"
                    Case "0435"
                        FinalStr = FinalStr & "е"
                    Case "0436"
                        FinalStr = FinalStr & "ж"
                    Case "0437"
                        FinalStr = FinalStr & "з"
                    Case "0438"
                        FinalStr = FinalStr & "и"
                    Case "0439"
                        FinalStr = FinalStr & "й"
                    Case "043A"
                        FinalStr = FinalStr & "к"
                    Case "043B"
                        FinalStr = FinalStr & "л"
                    Case "043C"
                        FinalStr = FinalStr & "м"
                    Case "043D"
                        FinalStr = FinalStr & "н"
                    Case "043E"
                        FinalStr = FinalStr & "о"
                    Case "043F"
                        FinalStr = FinalStr & "п"
                    Case "0440"
                        FinalStr = FinalStr & "р"
                    Case "0441"
                        FinalStr = FinalStr & "с"
                    Case "0442"
                        FinalStr = FinalStr & "т"
                    Case "0443"
                        FinalStr = FinalStr & "у"
                    Case "0444"
                        FinalStr = FinalStr & "ф"
                    Case "0445"
                        FinalStr = FinalStr & "х"
                    Case "0446"
                        FinalStr = FinalStr & "ц"
                    Case "0447"
                        FinalStr = FinalStr & "ч"
                    Case "0448"
                        FinalStr = FinalStr & "ш"
                    Case "0449"
                        FinalStr = FinalStr & "щ"
                    Case "044A"
                        FinalStr = FinalStr & "ъ"
                    Case "044B"
                        FinalStr = FinalStr & "ы"
                    Case "044C"
                        FinalStr = FinalStr & "ь"
                    Case "044D"
                        FinalStr = FinalStr & "э"
                    Case "044E"
                        FinalStr = FinalStr & "ю"
                    Case "044F"
                        FinalStr = FinalStr & "я"

                    Case "0022"
                        FinalStr = FinalStr & Chr(34)
                    Case "0028"
                        FinalStr = FinalStr & "("
                    Case "0029"
                        FinalStr = FinalStr & ")"
                    Case "002C"
                        FinalStr = FinalStr & ","
                    Case "002E"
                        FinalStr = FinalStr & "-"
                    Case "002F"
                        FinalStr = FinalStr & "/"
                    Case "0030"
                        FinalStr = FinalStr & "0"
                    Case "0031"
                        FinalStr = FinalStr & "1"
                    Case "0032"
                        FinalStr = FinalStr & "2"
                    Case "0033"
                        FinalStr = FinalStr & "3"
                    Case "0034"
                        FinalStr = FinalStr & "4"
                    Case "0035"
                        FinalStr = FinalStr & "5"
                    Case "0036"
                        FinalStr = FinalStr & "6"
                    Case "0037"
                        FinalStr = FinalStr & "7"
                    Case "0038"
                        FinalStr = FinalStr & "8"
                    Case "0039"
                        FinalStr = FinalStr & "9"
                    Case "003A"
                        FinalStr = FinalStr & ":"
                    Case "003B"
                        FinalStr = FinalStr & ";"
                    Case Else
                        FinalStr = FinalStr & "#"
                End Select
                If tstr.Length > 4 Then
                    FinalStr = FinalStr & Mid(tstr, 6)
                End If
            ElseIf Left(tstr, 1) Like "{" Then
                Continue For
            ElseIf Left(tstr, 1) Like "C" Then
                Dim pos As Integer = tstr.IndexOf(";")
                If pos > 0 Then
                    Dim tempArraystr As String() = tstr.Split(";")
                    If IsArray(tempArraystr) = True Then
                        If tempArraystr.Length > 0 Then
                            FinalStr = FinalStr & tempArraystr(tempArraystr.Length - 1)
                        End If
                    End If
                End If
            Else
                FinalStr = FinalStr & tstr
            End If
        Next
        FinalStr = FinalStr.Trim
        If Right(FinalStr, 1) Like "}" Then
            FinalStr = Mid(FinalStr, 1, FinalStr.Length - 1)
        End If
        FuncUnicodDecode = FinalStr
    End Function

    '============================================================================================================================
    'функция отделят от обозначения ЗУ его номер контура
    Public Shared Function FuncFindFullConditionalNumberCadastralObject(ByVal ConditionalNumber As String) As String
        FuncFindFullConditionalNumberCadastralObject = ConditionalNumber
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                Try
                    ConditionalNumber = ConditionalNumber.Trim
                    Dim pos1 As Integer = ConditionalNumber.LastIndexOf("(")
                    If pos1 > -1 Then
                        FuncFindFullConditionalNumberCadastralObject = Mid(ConditionalNumber, 1, pos1)
                    End If
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function

    '============================================================================================================================
    'функция возвращает номер контура
    Public Shared Function FuncFindNumberCounterCadastralObject(ByVal ConditionalNumber As String) As Integer
        FuncFindNumberCounterCadastralObject = 0
        Dim TempNumberStr As String = ""
        If IsNothing(ConditionalNumber) = False Then
            If ConditionalNumber.Trim.Length > 0 Then
                ConditionalNumber = ConditionalNumber.Trim
                Dim Pos1 As Integer = ConditionalNumber.LastIndexOf("(")
                Dim Pos2 As Integer = ConditionalNumber.LastIndexOf(")")
                If Pos1 > -1 And Pos2 > -1 Then
                    Dim StrCounter As String = Mid(ConditionalNumber, Pos1 + 2, Pos2 - (Pos1 + 1))
                    Dim TempNumber As String = ""
                    For i As Integer = 1 To StrCounter.Length
                        Dim sumb As String = Mid(StrCounter, i, 1)
                        If IsNumeric(sumb) Then
                            TempNumber = TempNumber & sumb
                        End If
                    Next
                    FuncFindNumberCounterCadastralObject = Val(TempNumber)
                End If
            End If
        End If
    End Function
    '================================================================================================================================
    'функция проверяет является ли строка форматом json
    Public Shared Function FuncValidateJson(ByVal jsonString As String,
                            Optional ByRef errorMessage As String = Nothing,
                            Optional ByRef jsonType As String = Nothing) As Boolean

        If String.IsNullOrWhiteSpace(jsonString) Then
            errorMessage = "Строка пустая или Nothing"
            Return False
        End If

        jsonString = jsonString.Trim()

        ' Быстрая проверка по начальным/конечным символам
        If jsonString.StartsWith("{") AndAlso jsonString.EndsWith("}") Then
            jsonType = "Object"
        ElseIf jsonString.StartsWith("[") AndAlso jsonString.EndsWith("]") Then
            jsonType = "Array"
        Else
            errorMessage = "Строка не начинается с { или ["
            Return False
        End If

        Try
            ' Пытаемся распарсить
            Dim token As JToken = JToken.Parse(jsonString)

            ' Дополнительные проверки
            If token.Type = JTokenType.Null Then
                errorMessage = "JSON содержит только null"
                Return False
            End If

            Return True

        Catch ex As JsonReaderException
            errorMessage = $"Ошибка парсинга JSON: {ex.Message}"
            Return False
        Catch ex As Exception
            errorMessage = $"Неожиданная ошибка: {ex.Message}"
            Return False
        End Try
    End Function
    '=============================================================================================================================
    'функция форматирует номер ряда балок
    Public Shared Function getConditionalRow(ByVal NumberRow As Integer) As String
        Dim result As String = ""
        If NumberRow < 0 Then
            result = "Л-" & Math.Abs(NumberRow)
        ElseIf NumberRow = 0 Then
            result = "Ось"
        Else
            result = "П-" & NumberRow
        End If
        Return result
    End Function
End Class
