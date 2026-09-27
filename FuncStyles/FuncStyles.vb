Imports System.IO
Imports Topomatic.Alg
Imports Topomatic.Alg.Offsets
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Dwg
Imports System.Text.RegularExpressions
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Undo
Imports System.Runtime.InteropServices
Imports System.Drawing.Text
Imports System.Drawing

Public Class FuncStyles
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'раблта со слоями
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция проверяет наличие типа линии на чертеже
    Public Shared Function getLineTypeDwg(ByVal ActivDocument As Drawing, ByVal NameLineType As String) As DwgLinetype
        Dim result As DwgLinetype = Nothing
        Try
            Dim DwgLineTypes As DwgLinetypes = ActivDocument.Linetypes
            If DwgLineTypes.Count > 0 Then
                For Each dwglineType As DwgLinetype In DwgLineTypes
                    If dwglineType.Name Like NameLineType.Trim Then
                        result = dwglineType
                    End If
                Next
            End If
        Catch ex As System.Exception
        End Try
        Return result
    End Function

    '========================================================================================================================================
    'функция создает тип линии по маске
    Public Shared Function FuncCreateTypeLineByPattern(ByVal ActivDocument As Drawing, ByVal NameTypeLine As String, ByVal DeskTypeLine As String, ByRef ArrayMask() As String, Optional templateXML As String = "") As DwgLinetype
        FuncCreateTypeLineByPattern = Nothing
        Dim strMess As String = ""
        Try
            'делаем проверку на наличие типа линии
            For i As Integer = 0 To ActivDocument.Linetypes.Count - 1
                If ActivDocument.Linetypes.Item(i).Name Like NameTypeLine.Trim Then
                    FuncCreateTypeLineByPattern = ActivDocument.Linetypes.Item(i)
                    Exit Function
                End If
            Next i
            Dim lenghtFirst As Double = -1
            'читаем массив
            If IsArray(ArrayMask) = True Then
                If ArrayMask.Length > 1 Then
                    Dim mask As Topomatic.Cad.Foundation.LinetypePattern = New Topomatic.Cad.Foundation.LinetypePattern
                    For i As Integer = 0 To ArrayMask.Length - 1
                        Dim TempSumb As String = ArrayMask(i).Trim
                        If TempSumb.Length = 0 Then Continue For
                        'пропускаем первый символ
                        If TempSumb Like "A" Then
                            Continue For
                            'числовое значение записываем
                        ElseIf IsNumeric(TempSumb) = True Then
                            Dim Zn As Double = Val(TempSumb)
                            mask.Add.DashDotLength = Zn
                            If lenghtFirst = -1 And Zn > 0 Then
                                lenghtFirst = Zn
                            End If
                        Else
                            '==================================================================
                            'сложный тип линии
                            If Left(TempSumb, 1) Like "[[]" Then
                                Dim sumbolElementMask As LinetypePatternItem = mask.Add 'создаем новый элемент
                                Dim nameSHPForm As String = "" 'переменная нужна для храния имени формы
                                'перебираем элементы пока не дойдем дл обратной скобки
                                For j As Integer = i To ArrayMask.Length - 1
                                    Dim tempSumb2 As String = ArrayMask(j).Trim
                                    If TempSumb.Length = 0 Then Continue For
                                    'имя формы или текста
                                    If i = j Then
                                        tempSumb2 = TempSumb.Replace("[", "") 'избавляемся от первой кавычки
                                        If Left(tempSumb2, 1) Like """" Then 'значит это обычная надпись, не форма
                                            'тексты в кавычках (убираем кавычки слева и справа)
                                            sumbolElementMask.ElementType = Topomatic.Cad.Foundation.LinetypePatternElementType.etTextString
                                            tempSumb2 = tempSumb2.Replace("""", "")
                                            sumbolElementMask.TextString = tempSumb2
                                            strMess = strMess & vbLf & "Текст линии: " & tempSumb2
                                        Else
                                            'иначе это форма (если кавычек нет)
                                            sumbolElementMask.ElementType = Topomatic.Cad.Foundation.LinetypePatternElementType.etShape
                                            tempSumb2 = tempSumb2.Replace("""", "") 'на всякий случай, кавычек вообще быть не должно
                                            nameSHPForm = tempSumb2
                                            strMess = strMess & vbLf & "Имя формы: " & nameSHPForm
                                        End If
                                        'читаем имя шрифта (второй элемент в квадратных скобках)
                                    ElseIf j = i + 1 Then
                                        'это форма
                                        If Right(tempSumb2, 4) Like ".shx" Or Right(tempSumb2, 4) Like ".shp" Then
                                            sumbolElementMask.Filename = tempSumb2
                                            Dim findFile As String = tempSumb2.Replace("shx", "shp")
                                            'ищем файл формы во вспомогательных файлах поддержки
                                            Dim fullNameFileShp As String = FuncFiles.getFileToDirectorySupport(findFile)
                                            If IO.File.Exists(fullNameFileShp) = True Then
                                                strMess = strMess & vbLf & "Файл формы (имя) " & findFile & ", найден. Полный путь: " & fullNameFileShp
                                                Dim Numbshp As Integer = -1
                                                Dim inputshp As StreamReader = New StreamReader(fullNameFileShp, True)
                                                Dim line2 As String
                                                Dim regMatch As String = "\*"
                                                Do Until inputshp.EndOfStream
                                                    line2 = inputshp.ReadLine().Trim 'считываем строку
                                                    If line2.Length > 0 Then
                                                        Dim flagread As String = Left(line2, 1)
                                                        If Regex.IsMatch(flagread, regMatch) Then 'chr(*)
                                                            Dim arrayShp As String() = line2.Split(",")
                                                            If IsArray(arrayShp) = True Then
                                                                If arrayShp.GetUpperBound(0) > 1 Then
                                                                    If arrayShp(2).Trim Like nameSHPForm Then
                                                                        Dim StrNumb As String = arrayShp(0).Trim
                                                                        If StrNumb.Length > 1 Then
                                                                            Numbshp = Val(Mid(StrNumb, 2))
                                                                            If Numbshp > 0 Then
                                                                                strMess = strMess & vbLf & "Номер формы: " & Numbshp
                                                                                Exit Do
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Loop
                                                inputshp.Close()
                                                If Numbshp > 0 Then
                                                    sumbolElementMask.ShapeNumber = Numbshp
                                                Else
                                                    MsgBox("Не удалос найти форму в указанном файле shp!!! Макро превано")
                                                    Exit Function
                                                End If
                                            Else
                                                MsgBox("Не удалось найти файл описысывающий форму данной линии!!!. Макро прервано.")
                                                Exit Function
                                            End If
                                        Else
                                            'ищем такой текстовый стиль
                                            Dim userTextStyle As DwgStyle = FuncStyles.FuncReturnTextStyle(ActivDocument, tempSumb2)
                                            If IsNothing(userTextStyle) = True Then
                                                'создаем новый текстовый стиль
                                                Dim ArrayProperties As String(,) = Nothing
                                                Dim boolFindTextStyle As Boolean = FuncXML.FuncReadTextStylePropertiesToXML(templateXML, tempSumb2, ArrayProperties)
                                                If IsArray(ArrayProperties) = True Then
                                                    userTextStyle = FuncStyles.FuncCreateTextStyleByArray(ActivDocument, tempSumb2, ArrayProperties)
                                                End If
                                            End If
                                            If IsNothing(userTextStyle) Then
                                                userTextStyle = FuncStyles.FuncCreateTextStyle(ActivDocument, tempSumb2, "Arial.ttf")
                                            End If
                                            If IsNothing(userTextStyle) Then
                                                userTextStyle = ActivDocument.ActiveStyle
                                                tempSumb2 = userTextStyle.Name
                                            End If
                                            sumbolElementMask.Style = tempSumb2
                                            strMess = strMess & vbLf & "Имя стиля текста: " & tempSumb2
                                        End If
                                    Else
                                        'ели элемент последний, то избавляемся от закрывающей последней скобки
                                        Dim flagEnd As Boolean = False
                                        If Right(tempSumb2, 1) Like "[]]" Then
                                            tempSumb2 = tempSumb2.Replace("]", "")
                                            flagEnd = True
                                        End If
                                        'разбиваем элементы по знаку равно
                                        Dim ArS As String() = tempSumb2.Split("=")
                                        Dim F As String = ""
                                        Dim v As Double = 0
                                        If IsArray(ArS) = True Then
                                            For k As Integer = 0 To ArS.Length - 1
                                                If ArS(k).Trim.Length > 0 Then
                                                    If IsNumeric(ArS(k).Trim) Then
                                                        v = Val(ArS(k).Trim)
                                                        v = Math.Round(v, 3)
                                                    Else
                                                        F = ArS(k).Trim
                                                    End If
                                                End If
                                            Next k
                                        End If
                                        If F Like "S" Or F Like "s" Then
                                            sumbolElementMask.Scale = v
                                            strMess = strMess & vbLf & "Масштаб формы: " & sumbolElementMask.Scale
                                        ElseIf F Like "R" Or F Like "r" Then
                                            sumbolElementMask.Rotation = v
                                        ElseIf F Like "U" Or F Like "u" Then
                                            sumbolElementMask.Rotation = v
                                        ElseIf F Like "X" Or F Like "x" Then
                                            sumbolElementMask.XOffset = v
                                        ElseIf F Like "Y" Or F Like "y" Then
                                            sumbolElementMask.YOffset = v
                                        End If
                                        'если элемент последний то выходим из цикла
                                        If flagEnd = True Then
                                            i = j
                                            Exit For
                                        End If
                                    End If
                                Next j
                            End If
                        End If
                    Next i
                    If mask.Item(mask.Count - 1).IsSimpleDashDot = True Then
                        Dim lenEndElement As Double = mask.Item(mask.Count - 1).DashDotLength
                        If lenEndElement < 0 Then
                            If lenghtFirst < 0.1 Then
                                lenghtFirst = lenEndElement / 2
                            ElseIf lenghtFirst < 10 Then
                                lenghtFirst = lenghtFirst / 10
                            Else
                                lenghtFirst = lenghtFirst / 20
                            End If
                            mask.Add.DashDotLength = lenghtFirst
                        End If
                    End If
                    Dim newUserLineType As DwgLinetype = ActivDocument.Linetypes.FindEqualityOrCreate(mask, NameTypeLine, DeskTypeLine)

                    Dim createPattern As LinetypePattern = newUserLineType.LinetypePattern
                    If createPattern.Count > 0 Then
                        Dim strpattern As String = "A,"
                        For i As Integer = 0 To createPattern.Count - 1
                            Dim el As LinetypePatternItem = createPattern.Item(i)
                            If el.ElementType = LinetypePatternElementType.etSimple Then
                                strpattern = strpattern & el.DashDotLength & ","
                            ElseIf el.ElementType = LinetypePatternElementType.etShape Then
                                strpattern = strpattern & "[" & el.ShapeNumber & "," & el.Filename & ",S=" & el.Scale & ",R=" & el.Rotation & ",X=" & el.XOffset & ",Y=" & el.YOffset & "]"
                            ElseIf el.ElementType = LinetypePatternElementType.etTextString Then
                                strpattern = strpattern & "[" & el.TextString & "," & el.Style & ",S=" & el.Scale & ",R=" & el.Rotation & ",X=" & el.XOffset & ",Y=" & el.YOffset & "]"
                            End If
                        Next i
                        strMess = strMess & vbLf & strpattern
                    End If
                    Return newUserLineType
                Else
                    'MsgBox("Массив описывающий выбранный тип линии имеет менее 3 элементов!!! Макро прервано!")
                    Exit Function
                End If
            Else
                'MsgBox("Массив описывающий выбранный тип линии пустой!!! Макро прервано!")
                Exit Function
            End If
        Catch ex As System.Exception
        End Try
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'раблта со слоями
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает слой по его имени
    Public Shared Function getLayerDwgByName(ByVal ActivDocument As Drawing, ByVal NameLayer As String) As DwgLayer
        Dim result As DwgLayer = Nothing
        If IsNothing(NameLayer) = False Then
            If NameLayer.Trim.Length > 0 Then
                If IsNothing(ActivDocument) = False Then
                    Try
                        Dim DwgLayers As DwgLayers = ActivDocument.Layers
                        For Each dwglayer As DwgLayer In DwgLayers
                            If dwglayer.Name Like NameLayer.Trim Then
                                result = dwglayer
                                Exit For
                            End If
                        Next
                    Catch ex As System.Exception
                    End Try
                End If
            End If
        End If
        Return result
    End Function
    'функция создает слой
    Public Shared Function CreateLayerDwg(ByVal ActivDocument As Drawing, ByVal nameLayer As String, Optional ByVal Indcolor As Integer = 7, Optional ByVal NameTypeLine As String = "", Optional ValLineWight As Integer = 20, Optional ByVal boolVisible As Boolean = True) As DwgLayer
        Dim result As DwgLayer = Nothing
        Try
            result = getLayerDwgByName(ActivDocument, nameLayer)
            If IsNothing(result) = True Then
                result = ActivDocument.Layers.Add(nameLayer.Trim)
                If IsNothing(result) = False Then
                    'назначаем цвет
                    Indcolor = Math.Abs(Indcolor)
                    If Indcolor > 256 Then
                        Indcolor = 256
                    End If
                    Dim CColor As CadColor = New CadColor(Indcolor)
                    result.Color = CColor
                    'назначаем тип линии слою
                    Dim LineTypes As DwgLinetypes = ActivDocument.Linetypes
                    Dim LineType As DwgLinetype = getLineTypeDwg(ActivDocument, NameTypeLine)
                    If IsNothing(LineType) = False Then
                        result.Linetype = LineType
                    End If
                    'назначаем вес линиям
                    result.Lineweight = ValLineWight
                    'назначаем видимость слою
                    If boolVisible = True Then
                        result.Visible = True
                    Else
                        result.Visible = False
                    End If
                End If
            End If
        Catch ex As System.ArgumentOutOfRangeException
        End Try
        Return result
    End Function




    '=======================================================================================================================================
    'функция возвращает текстовый стиль
    Public Shared Function FuncReturnTextStyle(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String) As DwgStyle
        FuncReturnTextStyle = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncReturnTextStyle = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i
        Catch ex As System.Exception
        End Try
    End Function

    '=======================================================================================================================================
    'функция возвращает все текстовые стили с чертежаь
    Public Shared Function FuncReturnTextStyles(ByVal ActivDocument As Drawing, ByRef arrayData As String()) As Boolean
        FuncReturnTextStyles = Nothing
        Dim countArray As Integer = 0
        If IsArray(arrayData) = True Then
            countArray = arrayData.Length
        End If
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                ReDim arrayData(countArray)
                arrayData(countArray) = ActivDocument.Styles.Item(i).Name
                countArray += 1
            Next i
        Catch ex As System.Exception
        End Try
    End Function

    '=======================================================================================================================================
    'функция создает текстовый стиль
    Public Shared Function FuncCreateTextStyleByArray(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String, ByVal arrayProperties As String(,)) As DwgStyle
        FuncCreateTextStyleByArray = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncCreateTextStyleByArray = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i
            Dim font As String = ""
            Dim ital As Boolean = False
            Dim anno As Boolean = False
            Dim textheight As Double = 0
            Dim rot As Double = 0
            Dim textprior As Double = 0
            If IsArray(arrayProperties) = True Then
                For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                    Dim field As String = arrayProperties(0, i)
                    Dim valText As String = arrayProperties(1, i)
                    If field Like "NameFont" Then
                        font = valText.Trim
                        Dim boolFont As Boolean = False
                        If Right(font, 4) Like ".ttf" Then
                            Dim userfontttf As CadFont = FontManager.Current.SearchTTF(font)
                            If IsNothing(userfontttf) = False Then
                                boolFont = True
                            Else
                                boolFont = False
                            End If
                            If boolFont = False Then
                                Dim col As InstalledFontCollection = New InstalledFontCollection()
                                Dim ffile As String = Path.GetFileNameWithoutExtension(font)
                                For Each fontt As FontFamily In col.Families
                                    If fontt.Name Like ffile Then
                                        boolFont = True
                                    End If
                                Next
                            End If
                        ElseIf Right(font, 4) Like ".shx" Or Right(font, 4) Like ".SHX" Then
                            Dim userfontttf As CadFont = FontManager.Current.SearchSHX(font)
                            If IsNothing(userfontttf) = True Then
                                boolFont = True
                            Else
                                boolFont = False
                            End If
                        End If
                        If boolFont = False Then
                            font = FontManager.Current.DefaultFont.FileName
                        End If
                    ElseIf field Like "Italic" Then
                        If valText Like "True" Then
                            ital = True
                        End If
                    ElseIf field Like "Annotative" Then
                        If valText Like "True" Then
                            anno = True
                        End If
                    ElseIf field Like "TextStyleHeight" Then
                        textheight = Val(valText)
                    ElseIf field Like "TextStyleRotate" Then
                        rot = Val(valText)
                    ElseIf field Like "TextStylePrior" Then
                        textprior = Val(valText)
                    End If
                Next i
            End If
            If font.Length > 0 Then
                FuncCreateTextStyleByArray = ActivDocument.Styles.Add(NameTextStyle.Trim, font)
                FuncCreateTextStyleByArray.Height = textheight
            End If
        Catch ex As System.Exception
        End Try
    End Function

    '=======================================================================================================================================
    'функция создает текстовый стиль
    Public Shared Function FuncCreateTextStyle(ByVal ActivDocument As Drawing, ByVal NameTextStyle As String, ByVal NameFonts As String, Optional HText As Double = 2.0, Optional rot As Double = 0) As DwgStyle
        FuncCreateTextStyle = Nothing
        Try
            For i As Integer = 0 To ActivDocument.Styles.Count - 1
                If ActivDocument.Styles.Item(i).Name Like NameTextStyle.Trim Then
                    FuncCreateTextStyle = ActivDocument.Styles.Item(i)
                    Exit Function
                End If
            Next i

            FuncCreateTextStyle = ActivDocument.Styles.Add(NameTextStyle.Trim, NameFonts)
            FuncCreateTextStyle.Height = HText
            FuncCreateTextStyle.Ratio = rot
        Catch ex As System.Exception
        End Try
    End Function

    '=======================================================================================================================================
    'функция возвращает код типа линии по его стилю
    Public Shared Function FuncReturnPatternLineTypeCode(ByVal userLineType As DwgLinetype) As String
        FuncReturnPatternLineTypeCode = "A,"
        If IsNothing(userLineType) = False Then
            Dim linePattern As LinetypePattern = userLineType.LinetypePattern
            If linePattern.Count > 0 Then
                Try
                    For i As Integer = 0 To linePattern.Count - 1
                        Dim element As LinetypePatternItem = linePattern.Item(i)
                        Dim typeElement As LinetypePatternElementType = element.ElementType
                        If typeElement = 0 Then
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & Math.Round(element.DashDotLength, 3)
                        ElseIf typeElement = 2 Then
                            Dim valStr As String = element.TextString
                            Dim styleName As String = element.Style
                            Dim rot As Double = Math.Round(element.Rotation, 3)
                            Dim scale As Double = Math.Round(element.Scale, 3)
                            Dim dx As Double = Math.Round(element.XOffset, 3)
                            Dim dy As Double = Math.Round(element.YOffset, 3)
                            Dim htext As Double = Math.Round(element.Height, 3)
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & "[" & Chr(34) & valStr & Chr(34) & "," & styleName & ",S=" & scale & ",R=" & rot & ",X=" & dx & ",Y=" & dy & "]"
                        ElseIf typeElement = 4 Then
                            Dim fontName As String = element.Font.FileName
                            fontName = IO.Path.GetFileNameWithoutExtension(fontName)
                            fontName = fontName & ".shp"
                            Dim valStr As String = FuncReturnNameSHPByIndex(element.ShapeNumber, fontName)
                            Dim rot As Double = Math.Round(element.Rotation, 3)
                            Dim scale As Double = Math.Round(element.Scale, 3)
                            Dim dx As Double = Math.Round(element.XOffset, 3)
                            Dim dy As Double = Math.Round(element.YOffset, 3)
                            Dim htext As Double = Math.Round(element.Height, 3)
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & "[" & valStr & "," & fontName & ",S=" & scale & ",R=" & rot & ",X=" & dx & ",Y=" & dy & "]"
                        End If
                        If i < linePattern.Count - 1 Then
                            FuncReturnPatternLineTypeCode = FuncReturnPatternLineTypeCode & ","
                        End If
                    Next i
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function

    '============================================================================================================================
    'функция возвращает имя формы по его коду
    Public Shared Function FuncReturnNameSHPByIndex(ByVal indexShape As Integer, ByVal nameFile As String) As String
        FuncReturnNameSHPByIndex = ""
        Dim putshFileSHP As String = FuncFiles.getFileToDirectorySupport(nameFile)
        If IO.File.Exists(putshFileSHP) = True Then
            'читаем файл с формой
            Dim sr As StreamReader = New StreamReader(putshFileSHP)
            Dim regMatch As String = "\*"
            While sr.Peek <> -1
                Dim st As String = sr.ReadLine().Trim
                Dim flagread As String = Left(st, 1)
                If Regex.IsMatch(flagread, regMatch) Then 'chr(42)
                    Dim rez As String() = st.Split(",")
                    If rez.GetUpperBound(0) > 1 Then
                        Dim indForm As String = Mid(rez(0), 2)
                        If indForm Like indexShape Then
                            FuncReturnNameSHPByIndex = rez(2)
                            Exit While
                        End If
                    End If
                End If
            End While
            sr.Close()
        End If
    End Function
End Class
