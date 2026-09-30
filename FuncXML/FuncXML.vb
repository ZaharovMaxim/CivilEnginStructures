Imports System.Environment
Imports System.IO
Imports System.Windows.Documents
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Topomatic.FoundationClasses

Public Class FuncXML
    Public Shared Function FuncGetPatchTemplates() As String
        FuncGetPatchTemplates = ""
        Try
            Dim TempDir As String = GetFolderPath(SpecialFolder.ApplicationData)
            If IO.Directory.Exists(TempDir & "\Civil3DToolsUtility\TOPOTemplates") = True Then
                TempDir = TempDir & "\Civil3DToolsUtility\TOPOTemplates"
            Else
                TempDir = ""
            End If
            Return TempDir
        Catch ex As System.Security.SecurityException
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает имена категории условных знаков с xml файла
    Public Shared Function FuncReadCategoryToXML(ByVal putchFileXml As String, ByRef arrayCategory As String()) As Boolean
        FuncReadCategoryToXML = False
        Dim countArrayCategory As Integer = 0
        If IsArray(arrayCategory) = True Then
            countArrayCategory = arrayCategory.Count
        End If
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    Dim value As String = reader.Value
                                    If IsNothing(value) = False Then
                                        If nameN.Trim Like "CategoryName" Then
                                            ReDim Preserve arrayCategory(countArrayCategory)
                                            arrayCategory(countArrayCategory) = reader.Value
                                            countArrayCategory += 1
                                        End If
                                    End If
                                End While
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            reader.Close()
                            Return True
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает имена таблиц условных знаков по известной категории
    Public Shared Function FuncReadTablesXMLByCategory(ByVal putchFileXml As String, ByVal nameCategory As String, ByRef arrayTables As String()) As Boolean
        FuncReadTablesXMLByCategory = False
        Dim countarrayTables As Integer = 0
        If IsArray(arrayTables) = True Then
            countarrayTables = arrayTables.Count
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameCategory Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim value As String = reader.Value
                                        If IsNothing(value) = False Then
                                            If nameN Like "TableName" Then
                                                ReDim Preserve arrayTables(countarrayTables)
                                                arrayTables(countarrayTables) = reader.Value
                                                countarrayTables += 1
                                            End If
                                        End If
                                    End While
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            reader.Close()
                            Return True
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает описания услоных знаков из определенной категории и таблицы условных знаков
    Public Shared Function FuncReadDeskObjectByCategoryAndTablesToXML(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByRef arrayObject As String()) As Boolean
        FuncReadDeskObjectByCategoryAndTablesToXML = False
        Dim countArrayObject As Integer = 0
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If table Like reader.Value Then
                                            boolTable = True
                                        Else
                                            boolTable = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If boolCategory = True And boolTable = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If typeObject Like reader.Value Then
                                            boolType = True
                                        Else
                                            boolType = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            If boolCategory = True And boolTable And boolType = True Then
                                Dim deskObject As String = reader.GetAttribute("Description")
                                ReDim Preserve arrayObject(countArrayObject)
                                arrayObject(countArrayObject) = deskObject
                                countArrayObject += 1
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все характеристики из определенной строки по известному дескриптору объекта наименованию таблицы и категории из файла xml (Имя категории, Имя таблицы, Тип объекта, Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataObjectToXMLbyDesk(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByVal deskObject As String, ByRef dataObjectProperties As String(,)) As Boolean
        FuncReadDataObjectToXMLbyDesk = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = nameN
                                        dataProperties(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                        If nameN Like "Description" Then
                                            If Not (NameObject.Trim Like deskObject.Trim) Then
                                                boolDataFind = False
                                                Exit While
                                            Else
                                                boolDataFind = True
                                            End If
                                        End If
                                    End While
                                    If boolDataFind = True Then
                                        dataObjectProperties = dataProperties
                                        FuncReadDataObjectToXMLbyDesk = True
                                        reader.Close()
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все характеристики из определенной строки по известному коду объекта, наименованию таблицы и категории из файла xml (Имя категории, Имя таблицы, Тип объекта, Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataObjectToXMLbyRoburCode(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByVal roburCode As String, ByRef dataObjectProperties As String(,)) As Boolean
        FuncReadDataObjectToXMLbyRoburCode = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = nameN
                                        dataProperties(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                        If nameN Like "Object_code" Then
                                            Dim arrayCode As String() = NameObject.Split(",")
                                            If IsArray(arrayCode) = True Then
                                                For k1 As Integer = 0 To arrayCode.Length - 1
                                                    Dim tempCode As String = arrayCode(k1)
                                                    If Not (tempCode Like roburCode) Then
                                                        boolDataFind = False
                                                    Else
                                                        boolDataFind = True
                                                        Exit For
                                                    End If
                                                Next k1
                                            End If
                                            If boolDataFind = False Then
                                                Exit While
                                            End If
                                        End If
                                    End While
                                    If boolDataFind = True Then
                                        dataObjectProperties = dataProperties
                                        FuncReadDataObjectToXMLbyRoburCode = True
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = "TypeObject"
                                        dataProperties(1, countDataProperties) = typeObject
                                        reader.Close()
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все характеристики слоя из файла xml
    Public Shared Function FuncReadPropertiesLayerToXML(ByVal putchFileXml As String, ByVal nameFindLayer As String, ByRef dataLayerProperties As String(,)) As Boolean
        FuncReadPropertiesLayerToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Layer" Then
                            Dim boolfindLayer As Boolean = False
                            Dim countDataLayerProperties As Integer = 0
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    Dim NameObject As String = reader.Value
                                    If nameN Like "Name" Then
                                        If Not (NameObject Like nameFindLayer) Then
                                            Exit While
                                        Else
                                            boolfindLayer = True
                                        End If
                                    Else
                                        ReDim Preserve dataLayerProperties(1, countDataLayerProperties)
                                        dataLayerProperties(0, countDataLayerProperties) = nameN
                                        dataLayerProperties(1, countDataLayerProperties) = NameObject
                                        countDataLayerProperties += 1
                                    End If
                                End While
                                If boolfindLayer = True Then
                                    reader.Close()
                                    Return True
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все характеристики типа линии из файла xml
    Public Shared Function FuncReadPropertiesLineTypeToXML(ByVal putchFileXml As String, ByVal nameFindLineType As String, ByRef dataLayerProperties As String(,)) As Boolean
        FuncReadPropertiesLineTypeToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            Dim boolfindLineType As Boolean = False
                            Dim countDataLineTypeProperties As Integer = 0
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameField As String = reader.Name
                                    Dim value As String = reader.Value
                                    If nameField Like "NameLineType" Then
                                        If Not (value Like nameFindLineType) Then
                                            Exit While
                                        Else
                                            boolfindLineType = True
                                        End If
                                    Else
                                        ReDim Preserve dataLayerProperties(1, countDataLineTypeProperties)
                                        dataLayerProperties(0, countDataLineTypeProperties) = nameField
                                        dataLayerProperties(1, countDataLineTypeProperties) = value
                                        countDataLineTypeProperties += 1
                                    End If
                                End While
                                If boolfindLineType = True Then
                                    reader.Close()
                                    Return True
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            reader.Close()
                            Exit Function
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает ищет имя таблицы объекта и возвращает все характеристики этого объекта
    Public Shared Function FuncReadNameTableAndPropertiesObjectToXML(ByVal putchFileXml As String, ByVal nameObject As String, ByRef dataObjectProperties As String(,), Optional ByVal nameCategory As String = "", Optional ByVal typeObject As String = "") As String
        FuncReadNameTableAndPropertiesObjectToXML = ""
        Dim nameTable As String = ""
        Dim boolCategory As Boolean = False
        Dim boolTypeObject As Boolean = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If nameCategory.Trim.Length = 0 Then
                                boolCategory = True
                            Else
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If nameCategory Like reader.Value Then
                                            boolCategory = True
                                        Else
                                            boolCategory = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If nameN Like "TableName" Then
                                            nameTable = reader.Value
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If typeObject.Trim.Length = 0 Then
                                boolTypeObject = True
                            Else
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If typeObject Like reader.Value Then
                                            boolTypeObject = True
                                        Else
                                            boolTypeObject = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolCategory = True And boolTypeObject = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameF As String = reader.Name
                                        Dim ValueData As String = reader.Value
                                        ReDim Preserve dataProperties(1, countDataProperties)
                                        dataProperties(0, countDataProperties) = nameF
                                        dataProperties(1, countDataProperties) = ValueData
                                        countDataProperties += 1
                                        If nameF Like "Name" Then
                                            If Not (nameObject Like ValueData) Then
                                                boolDataFind = False
                                                Exit While
                                            Else
                                                boolDataFind = True
                                            End If
                                        End If
                                    End While
                                    If boolDataFind = True Then
                                        dataObjectProperties = dataProperties
                                        reader.Close()
                                        Return nameTable
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает ищет имя таблицы объекта и возвращает все характеристики этого объекта
    Public Shared Function FuncReadPropertiesObjectByNoteToXML(ByVal putchFileXml As String, ByVal deskObject As String, ByRef dataObjectProperties As String(,), ByVal valueNote As String, Optional ByVal nameCategory As String = "", Optional ByVal nameTable As String = "", Optional typeObject As String = "") As Boolean
        FuncReadPropertiesObjectByNoteToXML = True
        Dim boolCategory As Boolean = False
        Dim boolTable As Boolean = False
        Dim boolTypeObject As Boolean = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If nameCategory.Trim.Length = 0 Then
                                boolCategory = True
                            Else
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If nameCategory Like reader.Value Then
                                            boolCategory = True
                                        Else
                                            boolCategory = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "Table" Then
                            If boolCategory = True Then
                                If nameTable.Trim.Length = 0 Then
                                    boolTable = True
                                Else
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim nameN As String = reader.Name
                                            If nameN Like "TableName" Then
                                                boolTable = True
                                            Else
                                                boolTable = False
                                            End If
                                        End While
                                    End If
                                End If
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If boolCategory = True And boolTable = True Then
                                If typeObject.Trim.Length = 0 Then
                                    boolTypeObject = True
                                Else
                                    If reader.HasAttributes = True Then
                                        While (reader.MoveToNextAttribute())
                                            Dim nameN As String = reader.Name
                                            If typeObject Like reader.Value Then
                                                boolTypeObject = True
                                            Else
                                                boolTypeObject = False
                                            End If
                                        End While
                                    End If
                                End If
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim dataProperties As String(,) = Nothing
                            Dim countDataProperties As Integer = 0
                            Dim boolDataFind As Boolean = False
                            If boolCategory = True And boolTable = True And boolTypeObject = True Then
                                If reader.HasAttributes = True Then
                                    Dim valNote As String = reader.GetAttribute("Note")
                                    Dim str As String = "ParentObject=" & deskObject
                                    Dim pos As Integer = valNote.IndexOf(str)
                                    If pos > -1 Then
                                        While (reader.MoveToNextAttribute())
                                            Dim nameF As String = reader.Name
                                            Dim ValueData As String = reader.Value
                                            ReDim Preserve dataProperties(1, countDataProperties)
                                            dataProperties(0, countDataProperties) = nameF
                                            dataProperties(1, countDataProperties) = ValueData
                                            countDataProperties += 1
                                        End While
                                        dataObjectProperties = dataProperties
                                        reader.Close()
                                        Return True
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function
    '/////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция ищет код типа линии по имени типа линии
    Public Shared Function FuncFindLineTypeCodeByNameTypeLineToXML(ByVal putchFileXml As String, ByVal nameLineType As String, Optional deskLine As String = "") As String
        FuncFindLineTypeCodeByNameTypeLineToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            Dim codeTypeLine As String = ""
                            Dim deskLineTemp As String = ""
                            Dim nameLineTemp As String = ""
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    Dim NameObject As String = reader.Value
                                    If nameN Like "NameLineType" Then
                                        nameLineTemp = NameObject
                                    ElseIf nameN Like "CodeLineType" Then
                                        codeTypeLine = NameObject
                                    ElseIf nameN Like "DeskLineType" Then
                                        deskLineTemp = NameObject
                                    End If
                                End While
                                If nameLineTemp.Trim.Length > 0 Then
                                    If nameLineTemp.Trim Like nameLineType Then
                                        deskLine = deskLineTemp
                                        reader.Close()
                                        Return codeTypeLine.Trim
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            reader.Close()
                            Exit Function
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет код по описанию объекта
    Public Shared Function FuncFindCodeObjectByDesk(ByVal putchFileXml As String, ByVal deck As String, Optional typeObject As String = "*") As String
        FuncFindCodeObjectByDesk = ""
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        Dim typeObjectUser As String = "*"
                        If NameBlock Like "TypeObject" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    typeObjectUser = reader.Value
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            If typeObjectUser Like typeObjectUser Then
                                Dim codeObject As String = ""
                                Dim deskObject As String = ""
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "Description" Then
                                            deskObject = NameObject
                                        ElseIf nameN Like "Object_code" Then
                                            codeObject = NameObject
                                        End If
                                    End While
                                End If
                                If codeObject.Trim.Length > 0 Then
                                    If deskObject.Trim Like deck.Trim Then
                                        reader.Close()
                                        Return codeObject
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция ищет код дорожной разметки
    Public Shared Function FuncFindCodeRoadsMarkToXML(ByVal putchFileXml As String, ByVal indexRoadsMark As String, ByRef dataRoadsMarkProperties As String(,), Optional speedRoad As String = "") As Boolean
        FuncFindCodeRoadsMarkToXML = False
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        Dim boolRoadsMark As Boolean = False
        If IsArray(dataRoadsMarkProperties) = True Then
            Erase dataRoadsMarkProperties
        End If
        Dim countdataRoadsMarkProperties As Integer = 0
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Table" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    If reader.Value Like "Дорожная разметка" Then
                                        boolRoadsMark = True
                                        Exit While
                                    Else
                                        boolRoadsMark = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            If boolRoadsMark = True Then
                                Dim strNote As String = ""
                                Dim boolFindElement As Boolean = False
                                If reader.HasAttributes = True Then
                                    ReDim Preserve dataRoadsMarkProperties(1, reader.AttributeCount + 1)
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        dataRoadsMarkProperties(0, countdataRoadsMarkProperties) = nameN
                                        dataRoadsMarkProperties(1, countdataRoadsMarkProperties) = NameObject
                                        countdataRoadsMarkProperties += 1
                                        If nameN Like "Note" Then
                                            Dim strIndex As String = ""
                                            Dim strSpeed As String = ""
                                            If IsNothing(NameObject) = False OrElse (NameObject.Trim.Length > 0) Then
                                                Dim tempArrayNote As String() = NameObject.Split(";")
                                                Dim indexFindMark As String = ""
                                                Dim speedFindMark As String = ""
                                                If IsArray(tempArrayNote) = True Then
                                                    For i As Integer = 0 To tempArrayNote.Length - 1
                                                        Dim noteStr As String = tempArrayNote(i)
                                                        Dim pos1 As Integer = noteStr.IndexOf("GOST_NUMBER")
                                                        Dim pos2 As Integer = noteStr.IndexOf("SPEED")
                                                        Dim pos3 As Integer = noteStr.IndexOf("=")
                                                        If pos1 > -1 Then
                                                            indexFindMark = Mid(noteStr, pos3 + 2)
                                                        ElseIf pos2 > -1 And pos3 > -1 Then
                                                            speedFindMark = Mid(noteStr, pos3 + 2)
                                                        End If
                                                    Next i
                                                End If
                                                If indexFindMark.Trim Like indexRoadsMark.Trim Then
                                                    If speedRoad.Trim.Length > 0 Then
                                                        If speedRoad.Trim Like speedFindMark.Trim Then
                                                            boolFindElement = True
                                                        End If
                                                    Else
                                                        boolFindElement = True
                                                    End If
                                                End If
                                                dataRoadsMarkProperties(0, countdataRoadsMarkProperties) = "GOST_NUMBER"
                                                dataRoadsMarkProperties(1, countdataRoadsMarkProperties) = indexFindMark
                                                countdataRoadsMarkProperties += 1
                                                dataRoadsMarkProperties(0, countdataRoadsMarkProperties) = "SPEED"
                                                dataRoadsMarkProperties(1, countdataRoadsMarkProperties) = speedFindMark
                                            End If
                                        End If
                                    End While
                                End If
                                If boolFindElement = True Then
                                    reader.Close()
                                    Return True
                                Else
                                    countdataRoadsMarkProperties = 0
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LibrarySymbols" Then
                            reader.Close()
                            Exit Function
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все имена типов линии из xml
    Public Shared Function FuncReadLineTypesToXML(ByVal putchFileXml As String, ByRef arrayLineTypes As String()) As Boolean
        FuncReadLineTypesToXML = False
        Dim countArrayLineTypes As Integer = 0
        If IsArray(arrayLineTypes) = True Then
            countArrayLineTypes = arrayLineTypes.Length
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            Dim boolfindLineType As Boolean = False
                            Dim countDataLineTypeProperties As Integer = 0
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameField As String = reader.Name
                                    Dim value As String = reader.Value
                                    If nameField Like "NameLineType" Then
                                        ReDim Preserve arrayLineTypes(countArrayLineTypes)
                                        arrayLineTypes(countArrayLineTypes) = value
                                        countArrayLineTypes += 1
                                    End If
                                End While
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecords" Then
                            reader.Close()
                            Exit Function
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читает имена всех стилей текста из файла XML
    Public Shared Function FuncReadTextStylesToXML(ByVal putchFileXml As String, ByRef nameTextStyle As String()) As Boolean
        FuncReadTextStylesToXML = False
        Dim countArrayTextStyles As Integer = 0
        If IsArray(nameTextStyle) = True Then
            countArrayTextStyles = nameTextStyle.Length
        End If
        If File.Exists(putchFileXml) = True Then
            'читаем xml и загружаем ветки с категориями
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "TextStyle" Then
                                Dim boolfindTextStyle As Boolean = False
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "NameTextStyle" Then
                                            If NameObject.Trim.Length > 0 Then
                                                ReDim Preserve nameTextStyle(countArrayTextStyles)
                                                nameTextStyle(countArrayTextStyles) = NameObject
                                                countArrayTextStyles += 1
                                            End If
                                        End If
                                    End While
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "TextStyles" Then
                                reader.Close()
                                Exit Function
                            End If
                    End Select
                End While
                reader.Close()
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читает характеристики стиля текста из файла XML
    Public Shared Function FuncReadTextStylePropertiesToXML(ByVal putchFileXml As String, ByVal nameTextStyle As String, ByRef arrayTextStyle As String(,)) As Boolean
        FuncReadTextStylePropertiesToXML = False
        Dim countArrayLineTypes As Integer = 0
        If File.Exists(putchFileXml) = True Then
            'читаем xml и загружаем ветки с категориями
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "TextStyle" Then
                                Dim boolfindTextStyle As Boolean = False
                                Dim countDataProperties As Integer = 0
                                Erase arrayTextStyle
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "NameTextStyle" Then
                                            If Not (NameObject Like nameTextStyle) Then
                                                Exit While
                                            Else
                                                boolfindTextStyle = True
                                            End If
                                        End If
                                        ReDim Preserve arrayTextStyle(1, countDataProperties)
                                        arrayTextStyle(0, countDataProperties) = nameN
                                        arrayTextStyle(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                    End While
                                    If boolfindTextStyle = True Then
                                        reader.Close()
                                        Return True
                                    End If
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "TextStyles" Then
                                reader.Close()
                                Exit Function
                            End If
                    End Select
                End While
                reader.Close()
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'читает характеристики стиля мультивыноски из файда XML
    Public Shared Function FuncReadMLederStylePropertiesToXML(ByVal putchFileXml As String, ByVal nameStyleMLeader As String, ByRef arrayMLeaderStyle As String(,)) As Boolean
        FuncReadMLederStylePropertiesToXML = False
        If File.Exists(putchFileXml) = True Then
            'читаем xml и загружаем ветки с категориями
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "MLeaderStyle" Then
                                Dim boolfindStyle As Boolean = False
                                Dim countDataProperties As Integer = 0
                                Erase arrayMLeaderStyle
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "NameMLeaderStyle" Then
                                            If Not (NameObject Like nameStyleMLeader) Then
                                                Exit While
                                            Else
                                                boolfindStyle = True
                                            End If
                                        End If
                                        ReDim Preserve arrayMLeaderStyle(1, countDataProperties)
                                        arrayMLeaderStyle(0, countDataProperties) = nameN
                                        arrayMLeaderStyle(1, countDataProperties) = NameObject
                                        countDataProperties += 1
                                    End While
                                    If boolfindStyle = True Then
                                        reader.Close()
                                        Return True
                                    End If
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "MLeaderStyle" Then
                                reader.Close()
                                Exit Function
                            End If
                    End Select
                End While
                reader.Close()
            End If
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает все имена таблиц pS из xml
    Public Shared Function FuncReadPSTablesToXML(ByVal putchFileXml As String, ByRef arrayPSTables As String()) As Boolean
        FuncReadPSTablesToXML = False
        Dim countarrayPSTables As Integer = 0
        If IsArray(arrayPSTables) = True Then
            countarrayPSTables = arrayPSTables.Length
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTable" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameField As String = reader.Name
                                    Dim value As String = reader.Value
                                    If nameField Like "TableName" Then
                                        ReDim Preserve arrayPSTables(countarrayPSTables)
                                        arrayPSTables(countarrayPSTables) = value
                                        countarrayPSTables += 1
                                    End If
                                End While
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTables" Then
                            reader.Close()
                            If IsArray(arrayPSTables) = True Then
                                Return True
                            Else
                                Return False
                            End If
                        End If
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает значение определенного поля таблицы PS из файла xml
    Public Shared Function FuncReadFieldValueByPSTablesToXML(ByVal putchFileXml As String, ByVal nameTablePS As String, ByVal nameFieldPS As String, ByRef arrayData As String()) As Boolean
        FuncReadFieldValueByPSTablesToXML = False
        Erase arrayData
        'читаем xml и загружаем ветки с категориями
        Dim boolTablePS As Boolean = False
        If IO.File.Exists(putchFileXml) = True Then
            Using reader As New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "PropertyTable" Then
                                Dim tableName As String = reader.GetAttribute("TableName")
                                boolTablePS = tableName IsNot Nothing AndAlso tableName.Trim Like nameTablePS
                            ElseIf NameBlock Like "Field" AndAlso boolTablePS Then
                                Dim fieldName As String = reader.GetAttribute("Name")
                                Dim fieldValue As String = reader.GetAttribute("Value")
                                If fieldName IsNot Nothing AndAlso fieldName Like nameFieldPS AndAlso
                                   fieldValue IsNot Nothing AndAlso fieldValue.Trim.Length > 0 Then
                                    arrayData = fieldValue.Split(";")
                                    Return True
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "PropertyTable" Then
                                boolTablePS = False
                            ElseIf NameBlock Like "PropertyTables" Then
                                Return False
                            End If
                    End Select
                End While
            End Using
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'конвертирование альбома балок из формата xls в формат xml
    Public Shared Function createXMLFileBeamsByExcel(ByVal putchXLSFiles As String(), ByVal putchFileXML As String) As Boolean
        createXMLFileBeamsByExcel = False
        Dim xws As XmlWriterSettings = New XmlWriterSettings()
        xws.Indent = True
        xws.NewLineOnAttributes = True
        Using xw As XmlWriter = XmlWriter.Create(putchFileXML, xws)
            xw.WriteStartDocument()
            xw.WriteStartElement("Beams")
            If IsArray(putchXLSFiles) = True Then
                Dim xlApp As Excel.Application = New Excel.Application()
                For i As Integer = 0 To putchXLSFiles.Length - 1
                    Dim fileXLS As String = putchXLSFiles(i)
                    If IO.File.Exists(fileXLS) = True Then
                        Dim xlWorkBook As Excel.Workbook = Nothing
                        Dim xlWorkSheet As Excel.Worksheet = Nothing
                        Try
                            xlWorkBook = xlApp.Workbooks.Open(fileXLS,, True)
                            xlWorkSheet = xlWorkBook.Worksheets.Item(1)
                            If xlWorkSheet IsNot Nothing Then
                                For indRows As Integer = 2 To xlWorkSheet.Rows.Count - 1
                                    For indColl As Integer = 1 To 100
                                        Dim nameField As String = xlWorkSheet.Cells(1, indColl).value
                                        Dim value As String = xlWorkSheet.Cells(indRows, indColl).value
                                        If IsNothing(nameField) = True Then
                                            Exit For
                                        End If
                                        If nameField.Trim.Length = 0 Then
                                            Exit For
                                        End If
                                        nameField = nameField.Replace(" ", "_")
                                        If IsNothing(value) = True Then
                                            value = ""
                                        End If
                                        If indColl = 1 And value.Trim.Length = 0 Then
                                            Exit For
                                        End If
                                        If indColl = 1 Then
                                            xw.WriteStartElement("Beam")
                                        End If
                                        xw.WriteAttributeString(nameField, value)
                                    Next indColl
                                    xw.WriteEndElement()
                                Next indRows
                            End If
                        Catch ex As Exception
                            If IsNothing(xlWorkBook) = False Then
                                xlWorkBook.Close(True)
                                xlWorkBook = Nothing
                            End If
                            If IsNothing(xlApp) = False Then
                                xlApp.Quit()
                            End If
                            Return False
                        Finally
                            If IsNothing(xlWorkBook) = False Then
                                xlWorkBook.Close(True)
                            End If
                        End Try
                    End If
                Next i
                If IsNothing(xlApp) = False Then
                    xlApp.Quit()
                End If
                xw.WriteEndElement()
                xw.WriteEndDocument()
                xw.Flush()
                xw.Close()
            End If
        End Using
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение альбома балок из файла xml
    Public Shared Function readAlbumBeamsFromXMLFiles(ByVal fullPatchXML As String, ByRef dictionaryBearm As Dictionary(Of String, BeamI()), ByRef arrayNameBeams As String(), Optional nameAlbum As String = "") As Boolean
        readAlbumBeamsFromXMLFiles = False
        Dim countArrayBeams As Integer = 0
        If IsArray(arrayNameBeams) = True Then
            If arrayNameBeams.Length > 0 Then
                countArrayBeams = arrayNameBeams.Length
            End If
        End If
        Dim countBeams As Integer = 0
        Dim arrayBeams As BeamI() = Nothing
        If dictionaryBearm.Count > 0 Then
            If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                arrayBeams = dictionaryBearm.Item(nameAlbum)
                countBeams = arrayBeams.Length
            End If
        End If
        If IO.File.Exists(fullPatchXML) = True Then
            Dim xDoc As XmlDocument = New XmlDocument()
            Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Beam" Then
                            If reader.HasAttributes = True Then
                                Dim userBeam As BeamI = New BeamI
                                While (reader.MoveToNextAttribute())
                                    Dim name As String = reader.Name
                                    Dim valN As String = reader.Value
                                    If name Like "model" Then
                                        userBeam.model = valN
                                        ReDim Preserve arrayNameBeams(countArrayBeams)
                                        arrayNameBeams(countArrayBeams) = userBeam.model
                                        countArrayBeams += 1
                                    ElseIf name Like "fullLenght" Then
                                        userBeam.lenght = Val(valN) / 1000
                                    ElseIf name Like "widthTopPlateLeft" Then
                                        userBeam.widthTopPlateLeft = Val(valN) / 1000
                                    ElseIf name Like "widthTopPlateRight" Then
                                        userBeam.widthTopPlateRight = Val(valN) / 1000
                                    ElseIf name Like "heightTopPlate" Then
                                        userBeam.heightTopPlate = Val(valN) / 1000
                                    ElseIf name Like "WidthTop" Then
                                        userBeam.WidthTop = Val(valN) / 1000
                                    ElseIf name Like "widthBottom" Then
                                        userBeam.widthBottom = Val(valN) / 1000
                                    ElseIf name Like "height" Then
                                        userBeam.height = Val(valN) / 1000
                                    ElseIf name Like "a" Then
                                        userBeam.a = Val(valN) / 1000
                                    ElseIf name Like "b" Then
                                        userBeam.b = Val(valN) / 1000
                                    ElseIf name Like "verticalOffsetRibZone" Then
                                        userBeam.VerticalOffsetRibZone = Val(valN) / 1000
                                    ElseIf name Like "horizontalOffsetRibZone" Then
                                        userBeam.HorizontalOffsetRibZone = Val(valN) / 1000
                                    ElseIf name Like "deltaRib" Then
                                        userBeam.DeltaRib = Val(valN)
                                    ElseIf name Like "radiusTop" Then
                                        userBeam.radiusTop = Val(valN) / 1000
                                    ElseIf name Like "radiusBottom" Then
                                        userBeam.radiusBottom = Val(valN) / 1000
                                    ElseIf name Like "gWidth" Then
                                        userBeam.gWidth = Val(valN) / 1000
                                    ElseIf name Like "mass" Then
                                        userBeam.mass = Val(valN)
                                    ElseIf name Like "volume" Then
                                        userBeam.volume = Val(valN)
                                    ElseIf name Like "modelTLS" Then
                                        userBeam.modelTLS = valN
                                    End If
                                    'userBeam.nameAlbum = nameAlbum
                                    'ReDim Preserve arrayBeams(countBeams)
                                    'arrayBeams(countBeams) = userBeam
                                    'countBeams += 1
                                End While
                                userBeam.nameAlbum = nameAlbum
                                ReDim Preserve arrayBeams(countBeams)
                                arrayBeams(countBeams) = userBeam
                                countBeams += 1
                            End If
                        End If
                End Select
            End While
            reader.Close()
            If IsArray(arrayBeams) = True Then
                If dictionaryBearm.ContainsKey(nameAlbum) = True Then
                    dictionaryBearm.Remove(nameAlbum)
                End If
                dictionaryBearm.Add(nameAlbum, arrayBeams)
            End If
            Return True
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение имен балок из альбома балок
    Public Shared Function readNamesBeamsFromXML(ByVal fullPatchXML As String, ByRef arrayNameBeams As String()) As Boolean
        readNamesBeamsFromXML = False
        Dim countArrayNameBeams As Integer = 0
        If IsArray(arrayNameBeams) = True Then
            If arrayNameBeams.Length > 0 Then
                countArrayNameBeams = arrayNameBeams.Length
            End If
        End If
        If IO.File.Exists(fullPatchXML) = True Then
            Dim xDoc As XmlDocument = New XmlDocument()
            Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Beam" Then
                            If reader.HasAttributes = True Then
                                Dim userBeam As BeamI = New BeamI
                                While (reader.MoveToNextAttribute())
                                    Dim name As String = reader.Name
                                    Dim valN As String = reader.Value
                                    If name Like "model" Then
                                        userBeam.model = valN
                                        ReDim Preserve arrayNameBeams(countArrayNameBeams)
                                        arrayNameBeams(countArrayNameBeams) = userBeam.model
                                        countArrayNameBeams += 1
                                        Exit While
                                    End If
                                End While
                            End If
                        End If
                End Select
            End While
            reader.Close()
            Return True
        End If
    End Function




    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает имена объектов из определенной категории условных знаков
    Public Shared Function FuncReadNameObjectCategoryToXML(ByVal putchFileXml As String, ByVal category As String, ByVal typeObject As String, ByRef arrayObject As String()) As Boolean
        FuncReadNameObjectCategoryToXML = False
        Dim countArrayObject As Integer = 0
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If boolCategory = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        If typeObject Like reader.Value Then
                                            boolType = True
                                        Else
                                            boolType = False
                                        End If
                                    End While
                                End If
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            If boolCategory = True And boolType = True Then
                                Dim nameBlk As String = reader.GetAttribute("Name")
                                ReDim Preserve arrayObject(countArrayObject)
                                arrayObject(countArrayObject) = nameBlk
                                countArrayObject += 1
                            End If
                        End If
                    Case XmlNodeType.EndElement
                End Select
            End While
            reader.Close()
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает определенное значение из массива
    Public Shared Function FuncReturnValueByArrayProperties(ByVal nameF As String, ByRef arrayPropertiesObject As String(,), Optional valueIndex As Integer = 1) As String
        FuncReturnValueByArrayProperties = ""
        If IsArray(arrayPropertiesObject) = True Then
            For i As Integer = 0 To arrayPropertiesObject.GetUpperBound(1)
                Dim nameField As String = arrayPropertiesObject(0, i)
                If IsNothing(nameField) = False Then
                    If nameField.Trim Like nameF Then
                        Return arrayPropertiesObject(valueIndex, i)
                    End If
                End If
            Next i
        End If
    End Function


End Class
