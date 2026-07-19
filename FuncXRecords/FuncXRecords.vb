Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Xml
Imports CivilEnginStructures.StructureElement
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Sfc
Imports Topomatic.Smt

Public Class FuncXRecords
    'чтение значений xData с примитива
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function getXRecords(ByVal entity As DwgEntity, ByRef xData As StructureElement, Optional nameTableXRecords As StructureElement.tableXRecords = StructureElement.tableXRecords.PROJECT_STRUCTURES) As Boolean
        getXRecords = False
        xData = New StructureElement()
        Try
            Dim userDictionary As DwgDictionary = entity.GetExtensionDictionary()
            If IsNothing(userDictionary) = True Then
                Return False
            Else
                Dim countRecDictionary As Integer = userDictionary.Count
                If countRecDictionary > 0 Then
                    For i As Integer = 0 To countRecDictionary - 1
                        'ищем словари
                        Dim tempUserDictionary As KeyValuePair(Of String, Object) = userDictionary.ElementAt(i)
                        Dim dictionaryName As String = tempUserDictionary.Key
                        If dictionaryName Like nameTableXRecords.ToString Then
                            'словарь найден
                            Dim tempStructureDictionary As DwgDictionary = tempUserDictionary.Value
                            Dim tempFieldDictionary As DwgDictionary = tempStructureDictionary.Item("Field")
                            'читаем имена полей
                            If tempFieldDictionary.Count > 1 Then
                                For k As Integer = 0 To tempFieldDictionary.Count - 1
                                    Dim rez As KeyValuePair(Of String, Object) = tempFieldDictionary.ElementAt(k)
                                    Dim nameF As String = rez.Key
                                    Dim valF As String = rez.Value
                                    If nameF Like "Label" Or nameF Like "LABEL" Then
                                        xData.Label = valF
                                    ElseIf nameF Like "ClassObject" Then
                                        xData.ClassObject = CType([Enum].Parse(GetType(StructureElement.classStructure), valF), StructureElement.classStructure)
                                    ElseIf nameF Like "Name" Or nameF Like "NAME" Then
                                        xData.Name = CType([Enum].Parse(GetType(StructureElement.typeObject), valF), StructureElement.typeObject)
                                    ElseIf nameF Like "Description" Then
                                        xData.Description = valF
                                    ElseIf nameF Like "KeyParameters" Then
                                        xData.KeyParameter = valF
                                    ElseIf nameF Like "IdStructure" Then
                                        xData.IdStructure = valF
                                    ElseIf nameF Like "IdElement" Then
                                        xData.IdElement = valF
                                    ElseIf nameF Like "Note" Or nameF Like "NOTE" Then
                                        xData.Note = valF
                                    End If
                                Next k
                                xData.IdObject = entity.ObjectID
                                xData.DWGEntity = entity
                                Return True
                            End If
                        ElseIf dictionaryName Like StructureElement.tableXRecords.PROJECT_BRIDGE.ToString Then
                            'читаем данные старого формата
                            'словарь найден
                            Dim arrayData As String(,) = {}
                            Dim tempDictionary2 As DwgDictionary = tempUserDictionary.Value
                            'читаем имена полей
                            If tempDictionary2.Count > 1 Then
                                Dim countArrayData As Integer = 0
                                For j As Integer = 2 To tempDictionary2.Count - 1
                                    Dim tempDictionary3 As KeyValuePair(Of String, Object) = tempDictionary2.ElementAt(j)
                                    'читаем строки таблмцы
                                    Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                                    If tempDictionary4.Count > 0 Then
                                        ReDim Preserve arrayData(5, countArrayData)
                                        For k As Integer = 0 To tempDictionary4.Count - 1
                                            Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                                            Dim nameF As String = rez.Key
                                            Dim valF As String = rez.Value
                                            If nameF Like "Name" Then
                                                arrayData(0, countArrayData) = valF
                                            ElseIf nameF Like "Value" Then
                                                arrayData(1, countArrayData) = valF
                                            ElseIf nameF Like "Description" Then
                                                arrayData(2, countArrayData) = valF
                                            ElseIf nameF Like "DefaultData" Then
                                                arrayData(3, countArrayData) = valF
                                            ElseIf nameF Like "TypeField" Then
                                                arrayData(4, countArrayData) = valF
                                            ElseIf nameF Like "DisplayOrder" Then
                                                arrayData(5, countArrayData) = valF
                                            End If
                                        Next k
                                        countArrayData += 1
                                    End If
                                Next j
                            End If
                            'анализируем структуру массива и заполняем класс
                            If arrayData.Length > 0 Then
                                For j As Integer = 0 To arrayData.GetUpperBound(1)
                                    Dim nameAF As String = arrayData(0, j)
                                    Dim valAF As String = arrayData(1, j)
                                    If nameAF Like "Label" Or nameAF Like "LABEL" Then
                                        xData.Label = valAF
                                    ElseIf nameAF Like "Name" Or nameAF Like "NAME" Then
                                        xData.Description = valAF
                                    ElseIf nameAF Like "Description" Then
                                        xData.Name = valAF
                                    ElseIf nameAF Like "KeyParameters" Then
                                        xData.KeyParameter = valAF
                                    ElseIf nameAF Like "IdStructure" Or nameAF Like "BrigeID" Then
                                        xData.IdStructure = valAF
                                    ElseIf nameAF Like "IdElement" Or nameAF Like "ElementID" Then
                                        xData.IdElement = valAF
                                    ElseIf nameAF Like "Note" Or nameAF Like "NOTE" Then
                                        xData.Note = valAF
                                    End If
                                Next j
                                xData.IdObject = entity.ObjectID
                                xData.Name = xData.getNameStructure(xData.Description)
                                'Return True
                            End If
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'запись значений xData в примитив
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function setXRecords(ByVal entity As DwgEntity, ByVal nameTable As StructureElement.tableXRecords, ByVal xData As StructureElement, Optional deskTableXRecords As String = "") As Boolean
        If entity Is Nothing OrElse xData Is Nothing Then
            Return False
        End If
        Try
            ' Получаем или создаем словарь расширений
            Dim userDictionary As DwgDictionary = entity.GetExtensionDictionary()
            If userDictionary Is Nothing Then
                userDictionary = CreateExtensionDictionary(entity)
            End If
            ' Ищем существующую таблицу
            Dim targetDictionary As DwgDictionary = FindTableDictionary(userDictionary, nameTable.ToString())
            If targetDictionary IsNot Nothing Then
                ' Обновляем существующую таблицу
                UpdateDictionaryData(targetDictionary, entity, xData)
            Else
                ' Создаем новую таблицу
                CreateNewTableDictionary(userDictionary, nameTable, deskTableXRecords, entity, xData)
            End If
            Return True
        Catch ex As System.Exception
            ' Логирование ошибки (рекомендуется добавить систему логирования)
            'Debug.WriteLine($"Ошибка при установке XRecords: {ex.Message}")
            Return False
        End Try
    End Function
    ' Вспомогательные приватные методы для разделения логики
    Private Shared Function CreateExtensionDictionary(ByVal entity As DwgEntity) As DwgDictionary
        entity.CreateExtensionDictionary()
        Return entity.GetExtensionDictionary()
    End Function

    Private Shared Function FindTableDictionary(ByVal userDictionary As DwgDictionary, ByVal tableName As String) As DwgDictionary
        If userDictionary Is Nothing OrElse userDictionary.Count = 0 Then
            Return Nothing
        End If
        ' Используем LINQ для поиска
        For Each kvp As KeyValuePair(Of String, Object) In userDictionary
            If String.Equals(kvp.Key, tableName, StringComparison.OrdinalIgnoreCase) Then
                Try
                    Return DirectCast(kvp.Value, DwgDictionary)
                Catch
                    Continue For
                End Try
            End If
        Next
        Return Nothing
    End Function

    Private Shared Sub UpdateDictionaryData(ByVal targetDictionary As DwgDictionary, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        ' Ищем словарь с данными полей
        For Each kvp As KeyValuePair(Of String, Object) In targetDictionary
            If String.Equals(kvp.Key, "Field", StringComparison.OrdinalIgnoreCase) Then
                Dim fieldDict As DwgDictionary = DirectCast(kvp.Value, DwgDictionary)
                SetFieldData(fieldDict, entity, xData)
                Exit For
            End If
        Next
    End Sub

    Private Shared Sub CreateNewTableDictionary(ByVal userDictionary As DwgDictionary, ByVal nameTable As StructureElement.tableXRecords, ByVal deskTableXRecords As String, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        Dim tableDictionary As DwgDictionary = userDictionary.AddDictionary(nameTable.ToString())
        ' Устанавливаем метаданные таблицы
        tableDictionary.SetString("TableName", nameTable.ToString())
        If Not String.IsNullOrEmpty(deskTableXRecords) Then
            tableDictionary.SetString("TableDescription", deskTableXRecords)
        End If
        ' Создаем словарь для полей данных
        Dim fieldDictionary As DwgDictionary = tableDictionary.AddDictionary("Field")
        SetFieldData(fieldDictionary, entity, xData)
    End Sub

    Private Shared Sub SetFieldData(ByVal fieldDictionary As DwgDictionary, ByVal entity As DwgEntity, ByVal xData As StructureElement)
        ' Устанавливаем данные полей с проверкой на пустые значения
        SetStringIfNotEmpty(fieldDictionary, "Label", xData.Label)
        SetStringIfNotEmpty(fieldDictionary, "ClassObject", xData.ClassObject)
        SetStringIfNotEmpty(fieldDictionary, "Name", xData.Name)
        SetStringIfNotEmpty(fieldDictionary, "Description", xData.Description)
        SetStringIfNotEmpty(fieldDictionary, "KeyParameters", xData.KeyParameter)
        SetStringIfNotEmpty(fieldDictionary, "IdStructure", xData.IdStructure)
        SetStringIfNotEmpty(fieldDictionary, "IdElement", xData.IdElement)
        SetStringIfNotEmpty(fieldDictionary, "Note", xData.Note)
        ' ObjectID всегда должен быть установлен
        fieldDictionary.SetString("IdObject", entity.ObjectID)
    End Sub
    Private Shared Sub SetStringIfNotEmpty(ByVal dictionary As DwgDictionary, ByVal key As String, ByVal value As String)
        If Not String.IsNullOrEmpty(value) Then
            dictionary.SetString(key, value)
        Else
            dictionary.SetString(key, "")
        End If
    End Sub








    'Public Shared Function setXRecords(ByVal entity As DwgEntity, ByVal nameTable As StructureElement.tableXRecords, ByVal xData As StructureElement, Optional deskTableXRecords As String = "") As Boolean
    '    setXRecords = False
    '    Dim createDictionary As DwgDictionary = Nothing
    '    Try
    '        Dim userDictionary As DwgDictionary = entity.GetExtensionDictionary()
    '        If IsNothing(userDictionary) = True Then
    '            entity.CreateExtensionDictionary()
    '            userDictionary = entity.GetExtensionDictionary()
    '            createDictionary = userDictionary.AddDictionary(nameTable.ToString)
    '            createDictionary.SetString("TableName", nameTable.ToString)
    '            createDictionary.SetString("TableDescription", deskTableXRecords)

    '            Dim newDictionary As DwgDictionary = createDictionary.AddDictionary("Field")
    '            newDictionary.SetString("Label", xData.Label)
    '            newDictionary.SetString("Name", xData.Name)
    '            newDictionary.SetString("Description", xData.Description)
    '            newDictionary.SetString("KeyParameters", xData.KeyParameter)
    '            newDictionary.SetString("IdStructure", xData.IdStructure)
    '            newDictionary.SetString("IdElement", xData.IdElement)
    '            newDictionary.SetString("Note", xData.Note)
    '            newDictionary.SetString("IdObject", entity.ObjectID)
    '            Return True
    '        Else
    '            Dim boolFindTables As Boolean = False
    '            If userDictionary.Count > 0 Then
    '                For i As Integer = 0 To userDictionary.Count - 1
    '                    'ищем словари
    '                    Dim tempDictionary As KeyValuePair(Of String, Object) = userDictionary.ElementAt(i)
    '                    Dim tempNameTable As String = tempDictionary.Key
    '                    If tempNameTable Like nameTable.ToString Then
    '                        Dim secondDict As DwgDictionary = tempDictionary.Value
    '                        Dim second2Dict As DwgDictionary = secondDict.ElementAt(2).Value
    '                        second2Dict.SetString("Label", xData.Label)
    '                        second2Dict.SetString("Name", xData.Name)
    '                        second2Dict.SetString("Description", xData.Description)
    '                        second2Dict.SetString("KeyParameters", xData.KeyParameter)
    '                        second2Dict.SetString("IdStructure", xData.IdStructure)
    '                        second2Dict.SetString("IdElement", xData.IdElement)
    '                        second2Dict.SetString("Note", xData.Note)
    '                        second2Dict.SetString("IdObject", entity.ObjectID)
    '                        Return True
    '                    End If
    '                Next
    '                If boolFindTables = False Then
    '                    createDictionary = userDictionary.AddDictionary(nameTable.ToString)
    '                    createDictionary.SetString("TableName", nameTable.ToString)
    '                    createDictionary.SetString("TableDescription", deskTableXRecords)

    '                    Dim newDictionary As DwgDictionary = createDictionary.AddDictionary("Field")
    '                    newDictionary.SetString("Label", xData.Label)
    '                    newDictionary.SetString("Name", xData.Name)
    '                    newDictionary.SetString("Description", xData.Description)
    '                    newDictionary.SetString("KeyParameters", xData.KeyParameter)
    '                    newDictionary.SetString("IdStructure", xData.IdStructure)
    '                    newDictionary.SetString("IdElement", xData.IdElement)
    '                    newDictionary.SetString("Note", xData.Note)
    '                    newDictionary.SetString("IdObject", entity.ObjectID)
    '                    Return True
    '                End If
    '            End If
    '        End If
    '    Catch ex As System.Exception
    '    End Try
    'End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читает имена и характеристики полей выбранной таблицы PS из файла XML (Результат - Имя словаря, вложенный словарь имя поля - характеристики)
    Public Shared Function FuncReadDataFieldsByTablePSToXML(ByVal putchFileXml As String, ByVal nameTablePS As String, ByRef dictionaryFields As Dictionary(Of String, String(,)), Optional ByRef deskTablePs As String = "") As Boolean
        FuncReadDataFieldsByTablePSToXML = False
        Dim boolTable As Boolean = False
        If File.Exists(putchFileXml) = True And nameTablePS.Trim.Length > 0 Then
            dictionaryFields = New Dictionary(Of String, String(,))
            'читаем xml и загружаем ветки с категориями
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            Dim countArrayTablePSFields As Integer = 0
                            If NameBlock Like "PropertyTable" Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "TableName" Then
                                            If Not (NameObject Like nameTablePS) Then
                                                boolTable = False
                                                Exit While
                                            Else
                                                boolTable = True
                                            End If
                                        ElseIf nameN Like "TableDescription" Then
                                            deskTablePs = NameObject
                                        End If
                                    End While
                                End If
                            ElseIf NameBlock Like "Field" Then
                                If boolTable = True Then
                                    Dim arrayRez As String(,) = Nothing
                                    Dim countArrayRez As Integer = 0
                                    If reader.HasAttributes = True Then
                                        Dim countAttr As Integer = reader.AttributeCount - 1
                                        Dim nameField As String = ""
                                        While (reader.MoveToNextAttribute())
                                            Dim nameF As String = reader.Name
                                            Dim valueZn As String = reader.Value
                                            'записываем в массив значения
                                            ReDim Preserve arrayRez(1, countArrayRez)
                                            arrayRez(0, countArrayRez) = nameF
                                            arrayRez(1, countArrayRez) = valueZn
                                            countArrayRez += 1
                                            If nameF Like "Name" Then
                                                nameField = valueZn
                                            End If
                                        End While
                                        If IsArray(arrayRez) = True And nameField.Trim.Length > 0 Then
                                            dictionaryFields.Add(nameField, arrayRez)
                                        End If
                                    End If
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "PropertyTables" Then
                                reader.Close()
                                If dictionaryFields.Count > 0 Then
                                    Return True
                                Else
                                    Return False
                                End If
                            ElseIf NameBlock Like "PropertyTable" Then
                                If dictionaryFields.Count > 0 Then
                                    reader.Close()
                                    Return True
                                End If

                            End If
                    End Select
                End While
                reader.Close()
            End If
        End If
    End Function

    'запись значений xData (массив имя поля - значение)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncCreateXData(ByVal acEnt As DwgObject, ByVal nameDict As String, ByVal ArrayData As String(,), Optional ByVal putchFileXml As String = "", Optional ByVal dictionaryFields As Dictionary(Of String, String(,)) = Nothing) As Boolean
        FuncCreateXData = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            Dim deskTable As String = ""
            If IsNothing(dictionaryFields) = True Then
                dictionaryFields = New Dictionary(Of String, String(,))
                Dim boolreadFieldsDict As Boolean = FuncReadDataFieldsByTablePSToXML(putchFileXml, nameDict, dictionaryFields, deskTable)
            End If
            'словарей нет вообще
            '=====================================================================================================================================
            If IsNothing(userDict) = True Then
                If dictionaryFields.Count > 0 Then
                    acEnt.CreateExtensionDictionary()
                    'добавляем новый словарь
                    userDict = acEnt.GetExtensionDictionary
                    If IsNothing(userDict) = False Then
                        Dim newDictionary As DwgDictionary = userDict.AddDictionary(nameDict)
                        If IsNothing(newDictionary) = False Then
                            newDictionary.SetString("TableName", nameDict)
                            newDictionary.SetString("TableDescription", deskTable)
                            'добавляем новому словарь поля
                            For i As Integer = 0 To dictionaryFields.Count - 1
                                'читаем поля
                                Dim dataFields As KeyValuePair(Of String, String(,)) = dictionaryFields.ElementAt(i)
                                'имя поля
                                Dim nameField As String = dataFields.Key
                                Dim newnewDictionary As DwgDictionary = newDictionary.AddDictionary(nameField)
                                Dim valField As String = ""
                                'массив с характеристиками
                                Dim arrayField As String(,) = dataFields.Value
                                If IsArray(arrayField) = True Then
                                    'ищем имя поля в массиве с характеристикаим
                                    Dim dataWrite As String = ""
                                    Dim boolWrite As Boolean = False
                                    For j As Integer = 0 To arrayField.GetUpperBound(1)
                                        nameField = arrayField(0, j)
                                        valField = arrayField(1, j)
                                        If IsArray(ArrayData) = True Then
                                            If IsNothing(nameField) = False Then
                                                If nameField Like "Name" Then
                                                    For k As Integer = 0 To ArrayData.GetUpperBound(1)
                                                        If ArrayData(0, k) Like valField Then
                                                            dataWrite = ArrayData(1, k)
                                                            boolWrite = True
                                                        End If
                                                    Next k
                                                ElseIf nameField Like "Value" Then
                                                    If boolWrite = True Then
                                                        valField = dataWrite
                                                        boolWrite = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                        newnewDictionary.SetString(nameField, valField)
                                        FuncCreateXData = True
                                    Next j
                                End If
                                'вставляем значения по умолчанию
                                Dim defDataStr As String = newnewDictionary.GetString("DefaultData", "")
                                If IsNothing(defDataStr) = True Then defDataStr = ""
                                Dim valDataStr As String = newnewDictionary.GetString("Value", "")
                                If IsNothing(valDataStr) = True Then valDataStr = ""
                                If valDataStr.Trim.Length = 0 And defDataStr.Trim.Length > 0 Then
                                    newnewDictionary.SetString("Value", defDataStr)
                                End If
                            Next i
                        End If
                    End If
                End If
            Else
                'у объекта уже есть подключенные словари
                '==================================================================================================================================
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    Dim boolFindDictionary As Boolean = False
                    For i As Integer = 0 To countDict - 1
                        Dim tempDictionary1 As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyNameDict As String = tempDictionary1.Key
                        'словарь найден
                        If keyNameDict Like nameDict Then
                            'если данных новых нет то далее ни чего не делаем
                            If IsArray(ArrayData) = True Then
                                Dim tempDictionary2 As DwgDictionary = tempDictionary1.Value
                                If tempDictionary2.Count > 1 Then
                                    Dim countArrayData As Integer = 0
                                    'читаем имена полей
                                    For j As Integer = 2 To tempDictionary2.Count - 1
                                        Dim tempDictionary3 As KeyValuePair(Of String, Object) = tempDictionary2.ElementAt(j)
                                        Dim keyNameFields As String = tempDictionary3.Key
                                        For j1 As Integer = 0 To ArrayData.GetUpperBound(1)
                                            Dim nameTempF As String = ArrayData(0, j1)
                                            Dim valTempF As String = ArrayData(1, j1)
                                            If nameTempF Like keyNameFields Then
                                                'читаем строки таблмцы
                                                Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                                                If tempDictionary4.Count > 0 Then
                                                    For k As Integer = 0 To tempDictionary4.Count - 1
                                                        Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                                                        Dim nameF As String = rez.Key
                                                        Dim valF As String = rez.Value
                                                        If nameF Like "Value" Then
                                                            tempDictionary4.SetString(nameF, valTempF)
                                                            FuncCreateXData = True
                                                            Exit For
                                                        End If
                                                    Next k
                                                    countArrayData += 1
                                                End If
                                            End If
                                        Next j1
                                    Next j
                                End If
                            End If
                            boolFindDictionary = True
                        End If
                        'Else
                        '    'читаем строки таблмцы
                        '    Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                        '    If tempDictionary4.Count > 0 Then
                        '        For k As Integer = 0 To tempDictionary4.Count - 1
                        '            Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                        '            Dim nameF As String = rez.Key
                        '            Dim valF As String = rez.Value
                        '            tempDictionary4.SetString(nameF, valF)
                        '            FuncCreateXData = True
                        '        Next k
                        '        'вставляем значения по умолчанию
                        '        Dim defDataStr As String = newnewDictionary.GetString("DefaultData", "")
                        '        Dim valDataStr As String = newnewDictionary.GetString("Value", "")
                        '        If valDataStr.Trim.Length = 0 And defDataStr.Trim.Length > 0 Then
                        '            newnewDictionary.SetString("Value", defDataStr)
                        '        End If
                        '        countArrayData += 1
                    Next i
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'словарь не найден, подключаем новый
                    If boolFindDictionary = False Then
                        'добаляем новый словарь
                        If dictionaryFields.Count > 0 Then
                            'добавляем новый словарь
                            userDict = acEnt.GetExtensionDictionary
                            If IsNothing(userDict) = False Then
                                Dim newDictionary As DwgDictionary = userDict.AddDictionary(nameDict)
                                If IsNothing(newDictionary) = False Then
                                    newDictionary.SetString("TableName", nameDict)
                                    newDictionary.SetString("TableDescription", deskTable)
                                    'добавляем новому словарь поля
                                    For i As Integer = 0 To dictionaryFields.Count - 1
                                        'читаем поля
                                        Dim dataFields As KeyValuePair(Of String, String(,)) = dictionaryFields.ElementAt(i)
                                        'имя поля
                                        Dim nameField As String = dataFields.Key
                                        Dim newnewDictionary As DwgDictionary = newDictionary.AddDictionary(nameField)
                                        Dim valField As String = ""
                                        'массив с характеристиками
                                        Dim arrayField As String(,) = dataFields.Value
                                        If IsArray(arrayField) = True Then
                                            'ищем имя поля в массиве с характеристикаим
                                            Dim dataWrite As String = ""
                                            Dim boolWrite As Boolean = False
                                            For j As Integer = 0 To arrayField.GetUpperBound(1)
                                                nameField = arrayField(0, j)
                                                valField = arrayField(1, j)
                                                If IsArray(ArrayData) = True Then
                                                    If IsNothing(nameField) = False Then
                                                        If nameField Like "Name" Then
                                                            For k As Integer = 0 To ArrayData.GetUpperBound(1)
                                                                If ArrayData(0, k) Like valField Then
                                                                    dataWrite = ArrayData(1, k)
                                                                    boolWrite = True
                                                                End If
                                                            Next k
                                                        ElseIf nameField Like "Value" Then
                                                            If boolWrite = True Then
                                                                valField = dataWrite
                                                                boolWrite = False
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                newnewDictionary.SetString(nameField, valField)
                                                FuncCreateXData = True
                                            Next j
                                        End If
                                        'вставляем значения по умолчанию
                                        Dim defDataStr As String = newnewDictionary.GetString("DefaultData", "")
                                        Dim valDataStr As String = newnewDictionary.GetString("Value", "")
                                        If valDataStr.Trim.Length = 0 And defDataStr.Trim.Length > 0 Then
                                            newnewDictionary.SetString("Value", defDataStr)
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                Else
                    'добаляем новый словарь
                    If dictionaryFields.Count > 0 Then
                        acEnt.CreateExtensionDictionary()
                        'добавляем новый словарь
                        userDict = acEnt.GetExtensionDictionary
                        If IsNothing(userDict) = False Then
                            Dim newDictionary As DwgDictionary = userDict.AddDictionary(nameDict)
                            If IsNothing(newDictionary) = False Then
                                newDictionary.SetString("TableName", nameDict)
                                newDictionary.SetString("TableDescription", deskTable)
                                'добавляем новому словарь поля
                                For i As Integer = 0 To dictionaryFields.Count - 1
                                    'читаем поля
                                    Dim dataFields As KeyValuePair(Of String, String(,)) = dictionaryFields.ElementAt(i)
                                    'имя поля
                                    Dim nameField As String = dataFields.Key
                                    Dim newnewDictionary As DwgDictionary = newDictionary.AddDictionary(nameField)
                                    Dim valField As String = ""
                                    'массив с характеристиками
                                    Dim arrayField As String(,) = dataFields.Value
                                    If IsArray(arrayField) = True Then
                                        'ищем имя поля в массиве с характеристикаим
                                        Dim dataWrite As String = ""
                                        Dim boolWrite As Boolean = False
                                        For j As Integer = 0 To arrayField.GetUpperBound(1)
                                            nameField = arrayField(0, j)
                                            valField = arrayField(1, j)
                                            If IsArray(ArrayData) = True Then
                                                If IsNothing(nameField) = False Then
                                                    If nameField Like "Name" Then
                                                        For k As Integer = 0 To ArrayData.GetUpperBound(1)
                                                            If ArrayData(0, k) Like valField Then
                                                                dataWrite = ArrayData(1, k)
                                                                boolWrite = True
                                                            End If
                                                        Next k
                                                    ElseIf nameField Like "Value" Then
                                                        If boolWrite = True Then
                                                            valField = dataWrite
                                                            boolWrite = False
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            newnewDictionary.SetString(nameField, valField)
                                            FuncCreateXData = True
                                        Next j
                                    End If
                                    'вставляем значения по умолчанию
                                    Dim defDataStr As String = newnewDictionary.GetString("DefaultData", "")
                                    Dim valDataStr As String = newnewDictionary.GetString("Value", "")
                                    If valDataStr.Trim.Length = 0 And defDataStr.Trim.Length > 0 Then
                                        newnewDictionary.SetString("Value", defDataStr)
                                    End If
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'чтение значений xData с определенного словаря
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadXData(ByVal acEnt As DwgObject, ByVal nameDictionary As String, ByRef arrayData As String(,)) As Boolean
        FuncReadXData = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return False
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    For i As Integer = 0 To countDict - 1
                        'ищем словари
                        Dim tempDictionary1 As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyDictName As String = tempDictionary1.Key
                        If keyDictName Like nameDictionary Then
                            'словарь найден
                            Dim tempDictionary2 As DwgDictionary = tempDictionary1.Value
                            'читаем имена полей
                            If tempDictionary2.Count > 1 Then
                                Dim countArrayData As Integer = 0
                                For j As Integer = 2 To tempDictionary2.Count - 1
                                    Dim tempDictionary3 As KeyValuePair(Of String, Object) = tempDictionary2.ElementAt(j)
                                    'читаем строки таблмцы
                                    Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                                    If tempDictionary4.Count > 0 Then
                                        ReDim Preserve arrayData(5, countArrayData)
                                        For k As Integer = 0 To tempDictionary4.Count - 1
                                            Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                                            Dim nameF As String = rez.Key
                                            Dim valF As String = rez.Value
                                            If nameF Like "Name" Then
                                                arrayData(0, countArrayData) = valF
                                            ElseIf nameF Like "Value" Then
                                                arrayData(1, countArrayData) = valF
                                            ElseIf nameF Like "Description" Then
                                                arrayData(2, countArrayData) = valF
                                            ElseIf nameF Like "DefaultData" Then
                                                arrayData(3, countArrayData) = valF
                                            ElseIf nameF Like "TypeField" Then
                                                arrayData(4, countArrayData) = valF
                                            ElseIf nameF Like "DisplayOrder" Then
                                                arrayData(5, countArrayData) = valF
                                            End If
                                        Next k
                                        countArrayData += 1
                                    End If
                                Next j
                            End If
                            If IsArray(arrayData) = True Then
                                Return True
                            End If
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'чтение определенного значения xData с определенного словаря
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadValueXDataByDictionary(ByVal acEnt As DwgObject, ByVal nameField As String, ByVal nameDictionary As String) As String
        FuncReadValueXDataByDictionary = ""
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return ""
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    Dim boolDictionary As Boolean = False
                    For i As Integer = 0 To countDict - 1
                        'ищем словари
                        Dim tempDictionary1 As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyDictName As String = tempDictionary1.Key
                        If nameDictionary.Trim.Length = 0 Then
                            boolDictionary = True
                        Else
                            If keyDictName Like nameDictionary Then
                                boolDictionary = True
                            End If
                        End If
                        If boolDictionary = True Then
                            'словарь найден
                            Dim tempDictionary2 As DwgDictionary = tempDictionary1.Value
                            'читаем имена полей
                            If tempDictionary2.Count > 1 Then
                                Dim countArrayData As Integer = 0
                                For j As Integer = 2 To tempDictionary2.Count - 1
                                    Dim tempDictionary3 As KeyValuePair(Of String, Object) = tempDictionary2.ElementAt(j)
                                    'читаем строки таблмцы
                                    Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                                    If tempDictionary4.Count > 0 Then
                                        Dim boolField As Boolean = False
                                        For k As Integer = 0 To tempDictionary4.Count - 1
                                            Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                                            Dim nameF As String = rez.Key
                                            Dim valF As String = rez.Value
                                            If nameF Like "Name" Then
                                                If nameF Like nameField Then
                                                    boolField = True
                                                End If
                                            ElseIf nameF Like "Value" Then
                                                If boolField = True Then
                                                    Return valF
                                                End If
                                            End If
                                        Next k
                                        countArrayData += 1
                                    End If
                                Next j
                            End If
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'чтение значений xData с определенного поля любого словаря (имя словаря возвращается)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadValueXData(ByVal acEnt As DwgObject, ByVal nameField As String, Optional ByRef nameTablePS As String = "") As String
        FuncReadValueXData = ""
        If IsNothing(acEnt) = True Then Return ""
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Exit Function
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    For i As Integer = 0 To countDict - 1
                        'ищем словари
                        Dim tempDictionary1 As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyDictName As String = tempDictionary1.Key
                        Dim tempDictionary2 As DwgDictionary = tempDictionary1.Value
                        'читаем имена полей
                        If tempDictionary2.Count > 1 Then
                            Dim countArrayData As Integer = 0
                            For j As Integer = 2 To tempDictionary2.Count - 1
                                Dim tempDictionary3 As KeyValuePair(Of String, Object) = tempDictionary2.ElementAt(j)
                                Dim namefieldDict As String = tempDictionary3.Key
                                If namefieldDict Like nameField Then
                                    Dim tempDictionary4 As DwgDictionary = tempDictionary3.Value
                                    If tempDictionary4.Count > 0 Then
                                        For k As Integer = 0 To tempDictionary4.Count - 1
                                            Dim rez As KeyValuePair(Of String, Object) = tempDictionary4.ElementAt(k)
                                            Dim nameF As String = rez.Key
                                            Dim valF As String = rez.Value
                                            If nameF Like "Value" Then
                                                If valF.Trim.Length > 0 Then
                                                    nameTablePS = keyDictName
                                                    Return valF
                                                End If
                                            ElseIf nameF Like "DefaultData" Then
                                                If valF.Trim.Length > 0 Then
                                                    nameTablePS = keyDictName
                                                    Return valF
                                                End If
                                            End If
                                        Next k
                                    End If
                                End If
                            Next j
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'удаление словаря xData
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncdeleteDictionaryXData(ByVal acEnt As DwgObject, ByVal arrayDictionary As String()) As Boolean
        FuncdeleteDictionaryXData = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return False
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    For i As Integer = 0 To countDict - 1
                        Dim tempDictionary As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyName As String = tempDictionary.Key
                        For j As Integer = 0 To arrayDictionary.Length - 1
                            Dim nameDict As String = arrayDictionary(j)
                            If nameDict Like keyName Then
                                userDict.Remove(keyName)
                                FuncdeleteDictionaryXData = True
                            End If
                        Next j
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'чтение словарей TablePS из объекта
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadTablesPS(ByVal acEnt As DwgObject, ByRef arrayTablesPS As String()) As Boolean
        FuncReadTablesPS = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return False
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    Dim countArrayTablePS As Integer = 0
                    For i As Integer = 0 To countDict - 1
                        Dim tempDictionary As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyName As String = tempDictionary.Key
                        ReDim Preserve arrayTablesPS(countArrayTablePS)
                        arrayTablesPS(countArrayTablePS) = keyName
                        countArrayTablePS += 1
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
        Return True
    End Function

    'проверка наличия определенного словаря
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncFindTablePS(ByVal acEnt As DwgObject, ByVal nameTablePS As String) As Boolean
        FuncFindTablePS = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return False
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    Dim countArrayTablePS As Integer = 0
                    For i As Integer = 0 To countDict - 1
                        Dim tempDictionary As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyName As String = tempDictionary.Key
                        If keyName Like nameTablePS.Trim Then
                            Return True
                        End If
                    Next i
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'запись значений xData в системный словарь
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncCreateXDataSystem(ByVal acEnt As DwgObject, ByVal nameSystemDictonary As String, ByVal ArrayData As String(,)) As Boolean
        FuncCreateXDataSystem = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                acEnt.CreateExtensionDictionary()
            End If
            userDict = acEnt.GetExtensionDictionary()
            If userDict.Count > 0 Then
                For i As Integer = 0 To userDict.Count - 1
                    'ищем словари
                    Dim tempDictionary1 As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                    Dim keyDictName As String = tempDictionary1.Key
                    If keyDictName Like nameSystemDictonary Then
                        userDict.Remove(nameSystemDictonary)
                    End If
                Next
            End If

            'добавляем новый словарь
            userDict = acEnt.GetExtensionDictionary
            Dim newDictionary As DwgDictionary = userDict.AddDictionary(nameSystemDictonary)
            If IsNothing(newDictionary) = False Then
                If IsNothing(newDictionary) = False Then
                    newDictionary.SetString("TableName", nameSystemDictonary)
                    newDictionary.SetString("TableDescription", "")
                    'добавляем новому словарь поля
                    For i As Integer = 0 To ArrayData.GetUpperBound(1)
                        Dim nameField As String = ArrayData(0, i)
                        Dim valField As String = ArrayData(1, i)
                        Dim newnewDictionary As DwgDictionary = newDictionary.AddDictionary(nameField)
                        newnewDictionary.SetString("Name", nameField)
                        newnewDictionary.SetString("Description", "")
                        newnewDictionary.SetString("Value", valField)
                        newnewDictionary.SetString("TypeField", "2")
                        newnewDictionary.SetString("DefaultData", "")
                        newnewDictionary.SetString("DisplayOrder", i + 1)
                    Next i
                End If
                Return True
            End If
        Catch ex As System.Exception
        End Try
    End Function

    'чтение значений xData
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Shared Function FuncReadXData1(ByVal acEnt As DwgObject, ByRef userDictRezult As Dictionary(Of String, String(,))) As Boolean
        FuncReadXData1 = False
        Try
            Dim userDict As DwgDictionary = acEnt.GetExtensionDictionary()
            If IsNothing(userDict) = True Then
                Return False
            Else
                Dim countDict As Integer = userDict.Count
                If countDict > 0 Then
                    For i As Integer = 0 To countDict - 1
                        Dim tempDictionary As KeyValuePair(Of String, Object) = userDict.ElementAt(i)
                        Dim keyName As String = tempDictionary.Key
                        Dim tempdwgDict As DwgDictionary = tempDictionary.Value
                        Dim arrayRez As String(,) = Nothing
                        Dim countRez As Integer = 0
                        For j As Integer = 0 To tempdwgDict.Count - 1
                            Dim rez As KeyValuePair(Of String, Object) = tempdwgDict.ElementAt(j)
                            ReDim Preserve arrayRez(4, countRez)
                            arrayRez(0, countRez) = rez.Key
                            arrayRez(1, countRez) = rez.Value
                            countRez += 1
                        Next j
                        userDictRezult.Add(keyName, arrayRez)
                    Next i
                End If
            End If
            If userDictRezult.Count > 0 Then
                Return True
            End If
        Catch ex As System.Exception
        End Try
    End Function




    'функция записывает данные в структурные линии
    Public Shared Function FuncWriteSemanticDataToStructureLine(ByVal structureLine As StructureLine, ByVal arrayData As String(,)) As Boolean
        FuncWriteSemanticDataToStructureLine = False
        If IsNothing(structureLine) = True Then
            Return Nothing
        End If
        Dim semanticData As SemanticDataSet = structureLine.LinearSemantic
        If semanticData.Count > 0 Then
            For i As Integer = 0 To semanticData.Count - 1
                If TypeOf semanticData.Root(i) Is SemanticStringNode Then
                    Dim nameData As SemanticStringNode = semanticData.Root(i)
                    Dim tagData As String = nameData.Tag
                    For j As Integer = 0 To arrayData.GetUpperBound(1)
                        Dim fieldData As String = arrayData(0, j)
                        If fieldData Like tagData Then
                            If IsNothing(arrayData(1, j)) = False Then
                                semanticData(tagData) = arrayData(1, j)
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    End Function

    'функция читает данные из структурной линии
    Public Shared Function FuncReadSemanticDataToStructureLine(ByVal structureLine As StructureLine, ByRef arrayData As String(,)) As Boolean
        FuncReadSemanticDataToStructureLine = False
        If IsNothing(structureLine) = True Then
            Return Nothing
        End If
        Dim semanticData As SemanticDataSet = structureLine.LinearSemantic
        If semanticData.Count > 0 Then
            For i As Integer = 0 To semanticData.Count - 1
                If TypeOf semanticData.Root(i) Is SemanticStringNode Then
                    Dim nameData As SemanticStringNode = semanticData.Root(i)
                    ReDim Preserve arrayData(1, i)
                    arrayData(0, i) = nameData.Tag
                    arrayData(1, i) = semanticData(nameData.Tag)
                End If
            Next i
        End If
    End Function
End Class
