Imports System.Windows.Forms
Imports Topomatic.Cad.View
Imports Topomatic.Dwg
Imports Topomatic.Cad.Foundation
Imports System.IO
Imports System.Drawing
Imports System.Xml

Partial Public Class FormEditXMLSymbol
    Public templateXML As String = ""
    Public boolIzmDataGrid As Boolean = True
    Public arrayTypeLine As String() 'массив с типами линий
    Public countarrayTypeLine As Integer = 0
    Public arrayLayers As String() 'массив со слоями
    Public countarrayLayers As Integer = 0
    Public arrayPS As String() 'массив с таблицами
    Public countarrayPS As Integer = 0
    Public userCadView As CadView = Nothing
    Public boolSelectObject As Boolean = False

    Private nodeUserClone As TreeNode = Nothing
    Private nameParentUserClone As String = ""
    Private handleID As IntPtr = Nothing

    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ' Добавить код инициализации после вызова InitializeComponent().
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If TreeView1.Nodes.Count = 0 Then
            Dim parentDirectory As String = FuncFiles.getFindDirectorySupport("InfrastradaToolsUtility")
            If parentDirectory.Trim.Length = 0 Then
                MsgBox("Каталог InfrastradaToolsUtility, во вспомогательных файлах поддержки не найден!!!")
                Exit Sub
            End If
            Dim pos As Integer = parentDirectory.IndexOf("InfrastradaToolsUtility")
            If pos > -1 Then
                parentDirectory = Mid(parentDirectory, 1, pos + 23)
            End If
            Dim directorySupport As String = parentDirectory & "\FileResources\Sample\"
            If Directory.Exists(directorySupport) = True Then
                Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
                If IsArray(xmlFiles) = True Then
                    For i As Integer = 0 To xmlFiles.Length - 1
                        '===============================================================================================
                        'условные знаки
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        Dim userNode As TreeNode = Nothing
                        Dim userNode0 As TreeNode = Nothing 'категория
                        Dim userNode1 As TreeNode = Nothing 'категория
                        Dim userNode2 As TreeNode = Nothing 'имя таблицы
                        Dim userNode3 As TreeNode = Nothing 'тип данных
                        Dim userNode4 As TreeNode = Nothing 'данные
                        '===============================================================================================
                        'слои
                        Dim userNode5 As TreeNode = Nothing 'идентификатор слоя
                        Dim userNode5_1 As TreeNode = Nothing 'имя фильтра
                        Dim userNode6 As TreeNode = Nothing 'данные
                        '===============================================================================================
                        'PS
                        Dim userNode7 As TreeNode = Nothing 'идентификатор таблицы
                        Dim userNode7_1 As TreeNode = Nothing 'имя категории
                        Dim userNode8 As TreeNode = Nothing 'имена таблиц
                        Dim userNode9 As TreeNode = Nothing 'поля таблиц
                        '===============================================================================================
                        'типы линии
                        Dim userNode10 As TreeNode = Nothing 'идентификатор таблицы
                        Dim userNode10_1 As TreeNode = Nothing 'имя категории
                        Dim userNode11 As TreeNode = Nothing 'имена таблиц
                        Dim userNode12 As TreeNode = Nothing 'поля таблиц
                        '===============================================================================================
                        'стили текста
                        Dim userNode13 As TreeNode = Nothing 'идентификатор стилей текста
                        Dim userNode14 As TreeNode = Nothing 'имена стилей текста
                        Dim userNode15 As TreeNode = Nothing 'поля с характерисиками
                        '===============================================================================================
                        'стили мультивыносок
                        Dim userNode16 As TreeNode = Nothing 'идентификатор стилей мультивыноски
                        Dim userNode17 As TreeNode = Nothing 'имена группы мультивыноски
                        Dim userNode18 As TreeNode = Nothing 'поля с характерисиками

                        'читаем xml
                        templateXML = xmlFiles(i)
                        Dim numberVersion As Integer = 0
                        Dim xDoc As XmlDocument = New XmlDocument()
                        If IO.File.Exists(templateXML) = True Then
                            Dim reader As XmlTextReader = New XmlTextReader(templateXML)
                            Dim arrayTag As String(,) = Nothing
                            Dim countArrayTag As Integer = 0
                            Dim count As Integer = 0
                            While reader.Read()
                                Select Case reader.NodeType
                                    Case XmlNodeType.Element
                                        count += 1
                                        Dim NameBlock As String = reader.Name 'читаем ветку
                                        If NameBlock Like "Classifier_ADSK" Then
                                            Dim arrayData As String(,) = Nothing
                                            Dim countArrayData As Integer = 0
                                            ReDim Preserve arrayData(1, countArrayData)
                                            arrayData(0, 0) = "Classifier_ADSK"
                                            arrayData(1, 0) = templateXML
                                            countArrayData += 1
                                            userNode = New TreeNode(NameBlock)
                                            userNode.Name = NameBlock
                                            If reader.HasAttributes = True Then
                                                While (reader.MoveToNextAttribute())
                                                    Dim name As String = reader.Name
                                                    Dim valN As String = reader.Value
                                                    ReDim Preserve arrayData(1, countArrayData)
                                                    arrayData(0, countArrayData) = name
                                                    arrayData(1, countArrayData) = valN
                                                    countArrayData += 1
                                                    If name Like "NameClassifier" Then
                                                        userNode.Text = valN
                                                    End If
                                                End While
                                            End If
                                            TreeView1.Nodes.Add(userNode)
                                            'TreeView2.Nodes.Add(userNode)
                                            userNode.Tag = arrayData
                                        ElseIf NameBlock Like "LibrarySymbols" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "LibrarySymbols"
                                                countArrayData += 1
                                                userNode0 = New TreeNode(NameBlock)
                                                userNode.Nodes.Add(userNode0)
                                                userNode0.Name = NameBlock
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameLibrarySymbols" Then
                                                            userNode0.Text = valN
                                                        ElseIf name Like "Version" Then
                                                            numberVersion = Val(valN)
                                                        End If
                                                    End While
                                                End If
                                                userNode0.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "Category" Then
                                            If IsNothing(userNode0) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Category"
                                                countArrayData += 1
                                                userNode1 = New TreeNode(NameBlock)
                                                userNode1.Name = NameBlock
                                                userNode0.Nodes.Add(userNode1)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "CategoryName" Then
                                                            userNode1.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode1.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "Table" Then
                                            If IsNothing(userNode1) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Table"
                                                countArrayData += 1
                                                userNode2 = New TreeNode(NameBlock)
                                                userNode2.Name = NameBlock
                                                userNode1.Nodes.Add(userNode2)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "TableName" Then
                                                            userNode2.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode2.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "TypeObject" Then
                                            If IsNothing(userNode2) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "TypeObject"
                                                countArrayData += 1
                                                userNode3 = New TreeNode(NameBlock)
                                                userNode3.Name = NameBlock
                                                userNode2.Nodes.Add(userNode3)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "Type" Then
                                                            userNode3.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode3.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "Rows" Then
                                            If IsNothing(userNode3) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Rows"
                                                countArrayData += 1
                                                userNode4 = New TreeNode(NameBlock)
                                                userNode4.Name = NameBlock
                                                userNode3.Nodes.Add(userNode4)
                                                If reader.HasAttributes = True Then
                                                    'переменные для заполнения массива с типами линий
                                                    Dim strNameTypeLine As String = ""
                                                    Dim strNameObject As String = ""
                                                    Dim strCodeTypeLine As String = ""
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "Description" Then
                                                            userNode4.Text = valN
                                                        ElseIf name Like "Name" Then
                                                            strNameObject = valN
                                                        ElseIf name Like "LineTypeCode" Then
                                                            strCodeTypeLine = valN
                                                        ElseIf name Like "LineType" Then
                                                            strNameTypeLine = valN
                                                        End If
                                                    End While
                                                    If IsArray(arrayData) = True Then
                                                        userNode4.Tag = arrayData
                                                    End If
                                                    If strCodeTypeLine.Trim.Length > 0 Then
                                                        If strNameTypeLine.Trim.Length > 0 Then
                                                            ReDim Preserve arrayTypeLine(countarrayTypeLine)
                                                            arrayTypeLine(countarrayTypeLine) = strNameTypeLine.Trim
                                                            countarrayTypeLine += 1
                                                        Else
                                                            ReDim Preserve arrayTypeLine(countarrayTypeLine)
                                                            arrayTypeLine(countarrayTypeLine) = strNameObject.Trim
                                                            countarrayTypeLine += 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'слои
                                        ElseIf NameBlock Like "Layers" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Layers"
                                                countArrayData += 1
                                                userNode5 = New TreeNode(NameBlock)
                                                userNode5.Name = NameBlock
                                                userNode.Nodes.Add(userNode5)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameLayers" Then
                                                            userNode5.Text = valN
                                                        End If
                                                    End While

                                                End If
                                                userNode5.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "LayerGroup" Then
                                            If IsNothing(userNode5) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "LayerGroup"
                                                countArrayData += 1
                                                userNode5_1 = New TreeNode(NameBlock)
                                                userNode5_1.Name = NameBlock
                                                userNode5.Nodes.Add(userNode5_1)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameGroup" Then
                                                            userNode5_1.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode5_1.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "Layer" Then
                                            If IsNothing(userNode5) = False Or IsNothing(userNode5_1) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Layer"
                                                countArrayData += 1
                                                userNode6 = New TreeNode(NameBlock)
                                                userNode6.Name = NameBlock
                                                If IsNothing(userNode5_1) = False Then
                                                    userNode5_1.Nodes.Add(userNode6)
                                                Else
                                                    userNode5.Nodes.Add(userNode6)
                                                End If
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "Name" Then
                                                            userNode6.Text = valN
                                                            ReDim Preserve arrayLayers(countarrayLayers)
                                                            arrayLayers(countarrayLayers) = valN
                                                            countarrayLayers += 1
                                                        End If
                                                    End While
                                                End If
                                                userNode6.Tag = arrayData
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'типы линий
                                        ElseIf NameBlock Like "LinetypeRecords" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "LinetypeRecords"
                                                countArrayData += 1
                                                userNode10 = New TreeNode(NameBlock)
                                                userNode10.Name = NameBlock
                                                userNode.Nodes.Add(userNode10)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameLineTypeRecords" Then
                                                            userNode10.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode10.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "LineTypeGroup" Then
                                            If IsNothing(userNode10) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "LineTypeGroup"
                                                countArrayData += 1
                                                userNode10_1 = New TreeNode(NameBlock)
                                                userNode10_1.Name = NameBlock
                                                userNode10.Nodes.Add(userNode10_1)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameGroupLinetypeRecords" Then
                                                            userNode10_1.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode10_1.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "LinetypeRecord" Then
                                            If IsNothing(userNode10) = False Or IsNothing(userNode10_1) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "LinetypeRecord"
                                                countArrayData += 1
                                                userNode11 = New TreeNode(NameBlock)
                                                userNode11.Name = NameBlock
                                                If IsNothing(userNode10_1) = False Then
                                                    userNode10_1.Nodes.Add(userNode11)
                                                Else
                                                    userNode10.Nodes.Add(userNode11)
                                                End If
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameLineType" Then
                                                            userNode11.Text = valN
                                                            'записываем слои в массив
                                                            ReDim Preserve arrayTypeLine(countarrayTypeLine)
                                                            arrayTypeLine(countarrayTypeLine) = valN
                                                            countarrayTypeLine += 1
                                                        End If
                                                    End While
                                                End If
                                                userNode11.Tag = arrayData
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'текстовые слои
                                        ElseIf NameBlock Like "TextStyles" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "TextStyles"
                                                countArrayData += 1
                                                userNode13 = New TreeNode(NameBlock)
                                                userNode13.Name = NameBlock
                                                userNode.Nodes.Add(userNode13)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameTextStyles" Then
                                                            userNode13.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode13.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "TextStyle" Then
                                            If IsNothing(userNode13) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "TextStyle"
                                                countArrayData += 1
                                                userNode14 = New TreeNode(NameBlock)
                                                userNode14.Name = NameBlock
                                                userNode13.Nodes.Add(userNode14)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameTextStyle" Then
                                                            userNode14.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode14.Tag = arrayData
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'мультивноски
                                        ElseIf NameBlock Like "MLeaderStyles" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "MLeaderStyles"
                                                countArrayData += 1
                                                userNode16 = New TreeNode(NameBlock)
                                                userNode16.Name = NameBlock
                                                userNode.Nodes.Add(userNode16)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameMLeaderStyles" Then
                                                            userNode16.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode16.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "MLeaderStylesGroup" Then
                                            If IsNothing(userNode16) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "MLeaderStylesGroup"
                                                countArrayData += 1
                                                userNode17 = New TreeNode(NameBlock)
                                                userNode17.Name = NameBlock
                                                userNode16.Nodes.Add(userNode17)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameMLeaderStylesGroup" Then
                                                            userNode17.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode17.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "MLeaderStyle" Then
                                            If IsNothing(userNode16) = False Or IsNothing(userNode17) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "MLeaderStyle"
                                                countArrayData += 1
                                                userNode18 = New TreeNode(NameBlock)
                                                userNode18.Name = NameBlock
                                                If IsNothing(userNode17) = False Then
                                                    userNode17.Nodes.Add(userNode18)
                                                Else
                                                    userNode16.Nodes.Add(userNode18)
                                                End If
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "Name" Then
                                                            userNode18.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode18.Tag = arrayData
                                            End If
                                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                            'таблицы PS
                                        ElseIf NameBlock Like "PropertyTables" Then
                                            If IsNothing(userNode) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "PropertyTables"
                                                countArrayData += 1
                                                userNode7 = New TreeNode(NameBlock)
                                                userNode7.Name = NameBlock
                                                userNode.Nodes.Add(userNode7)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NamePropertyTables" Then
                                                            userNode7.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode7.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "PSGroup" Then
                                            If IsNothing(userNode7) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "PSGroup"
                                                countArrayData += 1
                                                userNode7_1 = New TreeNode(NameBlock)
                                                userNode7_1.Name = NameBlock
                                                userNode7.Nodes.Add(userNode7_1)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "NameGroupPS" Then
                                                            userNode7_1.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode7_1.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "PropertyTable" Then
                                            If IsNothing(userNode7) = False Or IsNothing(userNode7_1) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "PropertyTable"
                                                countArrayData += 1
                                                userNode8 = New TreeNode(NameBlock)
                                                userNode8.Name = NameBlock
                                                If IsNothing(userNode7_1) = False Then
                                                    userNode7_1.Nodes.Add(userNode8)
                                                Else
                                                    userNode7.Nodes.Add(userNode8)
                                                End If
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "TableName" Then
                                                            userNode8.Text = valN
                                                            ReDim Preserve arrayPS(countarrayPS)
                                                            arrayPS(countarrayPS) = valN
                                                            countarrayPS += 1
                                                        End If
                                                    End While
                                                End If
                                                userNode8.Tag = arrayData
                                            End If
                                        ElseIf NameBlock Like "Field" Then
                                            If IsNothing(userNode8) = False Then
                                                Dim arrayData As String(,) = Nothing
                                                Dim countArrayData As Integer = 0
                                                ReDim Preserve arrayData(1, countArrayData)
                                                arrayData(0, countArrayData) = "Field"
                                                countArrayData += 1
                                                userNode9 = New TreeNode(NameBlock)
                                                userNode9.Name = NameBlock
                                                userNode8.Nodes.Add(userNode9)
                                                If reader.HasAttributes = True Then
                                                    While (reader.MoveToNextAttribute())
                                                        Dim name As String = reader.Name
                                                        Dim valN As String = reader.Value
                                                        ReDim Preserve arrayData(1, countArrayData)
                                                        arrayData(0, countArrayData) = name
                                                        arrayData(1, countArrayData) = valN
                                                        countArrayData += 1
                                                        If name Like "Name" Then
                                                            userNode9.Text = valN
                                                        End If
                                                    End While
                                                End If
                                                userNode9.Tag = arrayData
                                            End If
                                        End If
                                End Select
                            End While
                            reader.Close()
                        End If
                    Next
                End If
            End If
        End If


    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор узла дерева
    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Try
            Label1.Text = e.Node.Index
            Dim selectNode As TreeNode = e.Node
            boolIzmDataGrid = False
            If IsArray(selectNode.Tag) = True Then
                Dim tagSelectNode As String(,) = selectNode.Tag
                If tagSelectNode.GetUpperBound(1) > 0 Then
                    'чистим список
                    DataGridView1.Rows.Clear()
                    For i As Integer = 0 To tagSelectNode.GetUpperBound(1)
                        'добавляем данные
                        boolIzmDataGrid = False
                        DataGridView1.Rows.Add()
                        DataGridView1.Rows(i).Cells(0).Value = tagSelectNode(0, i)
                        DataGridView1.Rows(i).Cells(1).Value = tagSelectNode(1, i)
                    Next i
                    If IsNothing(selectNode.Parent) = False Then
                        DataGridView1.Tag = selectNode.Parent.Text
                    End If
                End If
            End If
            boolIzmDataGrid = True
        Catch ex As System.NullReferenceException
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'редактирование таблицы DataGrid
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If boolIzmDataGrid = True Then
            Try
                Dim userNode As TreeNode = TreeView1.SelectedNode
                Dim indRows As Integer = e.RowIndex
                Dim nameField As String = DataGridView1.Rows(indRows).Cells(0).Value
                Dim valueData As String = DataGridView1.Rows(indRows).Cells(1).Value
                If IsArray(userNode.Tag) Then
                    Dim arrayData As String(,) = userNode.Tag
                    If arrayData.GetUpperBound(1) > 0 Then
                        For i As Integer = 1 To arrayData.GetUpperBound(1)
                            If arrayData(0, i).Trim Like nameField Then
                                arrayData(1, i) = valueData
                                DataGridView1.Rows(indRows).Cells(1).Style.ForeColor = Color.Green
                                'If arrayData(0, i).Trim Like arrayData(1, 0).Trim Then
                                'userNode.Text = valueData
                                'End If
                                Exit For
                            End If
                        Next i
                    End If
                    userNode.Tag = arrayData
                End If
            Catch ex As System.ArgumentOutOfRangeException
            End Try
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'добавить новую ветку
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim userNode As TreeNode = TreeView1.SelectedNode
        If IsArray(userNode.Tag) = True Then
            Dim arrayData As String(,) = userNode.Tag
            Dim ScreenField As String = arrayData(1, 0)
            Dim nameNewNode As String = "Копия " & userNode.Text
            Dim copyArrayTag As String(,) = Nothing
            ReDim Preserve copyArrayTag(1, arrayData.GetUpperBound(1))
            For i As Integer = 0 To arrayData.GetUpperBound(1)
                copyArrayTag(0, i) = arrayData(0, i)
                copyArrayTag(1, i) = arrayData(1, i)
                If arrayData(0, i).Trim Like ScreenField Then
                    copyArrayTag(1, i) = nameNewNode
                End If
            Next i
            Dim clonedNode As TreeNode = New TreeNode(nameNewNode)
            clonedNode.Tag = copyArrayTag
            userNode.Parent.Nodes.Insert(0, clonedNode)
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'удалить ветку
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim userNode As TreeNode = TreeView1.SelectedNode
        Dim nameNode As String = userNode.Text
        Dim rez As MsgBoxResult = MsgBox("Удалить ветку " & nameNode & "?", MsgBoxStyle.OkCancel)
        If rez.Ok Then
            userNode.Remove()
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'записать новый XML
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim parentDirectory As String = FuncFiles.getFindDirectorySupport("InfrastradaToolsUtility")
        If parentDirectory.Trim.Length = 0 Then
            MsgBox("Каталог InfrastradaToolsUtility, во вспомогательных файлах поддержки не найден!!!")
            Exit Sub
        End If
        Dim pos As Integer = parentDirectory.IndexOf("InfrastradaToolsUtility")
        If pos > -1 Then
            parentDirectory = Mid(parentDirectory, 1, pos + 23)
        End If
        Dim directorySupport As String = parentDirectory & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            If Directory.Exists(directorySupport & "OldVersion") = False Then
                directorySupport = directorySupport & "OldVersion"
                IO.Directory.CreateDirectory(directorySupport)
            Else
                directorySupport = directorySupport & "OldVersion"
            End If
            Dim nameDir As String = Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second
            directorySupport = directorySupport & "\" & nameDir
            IO.Directory.CreateDirectory(directorySupport)
        End If
        'читаем дочерние элементы
        Dim startNodeADSK As TreeNodeCollection = TreeView1.Nodes
        Dim xws As XmlWriterSettings = New XmlWriterSettings()
        xws.Indent = True
        xws.NewLineOnAttributes = True
        Dim xw As XmlWriter = Nothing
        If startNodeADSK.Count > 0 Then
            For i As Integer = 0 To startNodeADSK.Count - 1
                Dim userNode0 As TreeNode = startNodeADSK.Item(i)
                Dim arrayUserTag As String(,) = userNode0.Tag
                If IsArray(arrayUserTag) = True Then
                    Dim nameStartNode As String = userNode0.Name
                    If nameStartNode Like "Classifier_ADSK" Then
                        Dim valStartNode As String = arrayUserTag(1, 0) 'читаем путь к файлу
                        If File.Exists(valStartNode) = True Then
                            Dim tempfile As String = Path.GetFileName(valStartNode)
                            IO.File.Copy(valStartNode, directorySupport & "\" & tempfile, True)
                            xw = XmlWriter.Create(valStartNode, xws)
                            xw.WriteStartDocument()
                            xw.WriteStartElement(nameStartNode)
                            For n As Integer = 1 To arrayUserTag.GetUpperBound(1)
                                xw.WriteAttributeString(arrayUserTag(0, n), arrayUserTag(1, n))
                            Next n
                        End If
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'условнее знаки, Типы линий и т.п.
                        Dim node1 As TreeNodeCollection = userNode0.Nodes
                        If node1.Count > 0 Then
                            For i1 As Integer = 0 To node1.Count - 1
                                Dim userNode1 As TreeNode = node1.Item(i1)
                                Dim arrayTag1 As String(,) = userNode1.Tag
                                xw.WriteStartElement(arrayTag1(0, 0))
                                For n As Integer = 1 To arrayTag1.GetUpperBound(1)
                                    xw.WriteAttributeString(arrayTag1(0, n), arrayTag1(1, n))
                                Next n
                                'третий уровень с категориями знаков
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                Dim node2 As TreeNodeCollection = userNode1.Nodes
                                If node2.Count > 0 Then
                                    For i2 As Integer = 0 To node2.Count - 1
                                        Dim userNode2 As TreeNode = node2.Item(i2)
                                        Dim arrayTag2 As String(,) = userNode2.Tag
                                        xw.WriteStartElement(arrayTag2(0, 0))
                                        For n As Integer = 1 To arrayTag2.GetUpperBound(1)
                                            xw.WriteAttributeString(arrayTag2(0, n), arrayTag2(1, n))
                                        Next n
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'ищем таблицы
                                        Dim node3 As TreeNodeCollection = userNode2.Nodes
                                        If node3.Count > 0 Then
                                            For i3 As Integer = 0 To node3.Count - 1
                                                Dim userNode3 As TreeNode = node3.Item(i3)
                                                Dim arrayTag3 As String(,) = userNode3.Tag
                                                xw.WriteStartElement(arrayTag3(0, 0))
                                                For n As Integer = 1 To arrayTag3.GetUpperBound(1)
                                                    xw.WriteAttributeString(arrayTag3(0, n), arrayTag3(1, n))
                                                Next n
                                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                'тип данных
                                                Dim node4 As TreeNodeCollection = userNode3.Nodes
                                                If node4.Count > 0 Then
                                                    For i4 As Integer = 0 To node4.Count - 1
                                                        Dim userNode4 As TreeNode = node4.Item(i4)
                                                        Dim arrayTag4 As String(,) = userNode4.Tag
                                                        xw.WriteStartElement(arrayTag4(0, 0))
                                                        For n As Integer = 1 To arrayTag4.GetUpperBound(1)
                                                            xw.WriteAttributeString(arrayTag4(0, n), arrayTag4(1, n))
                                                        Next n
                                                        xw.WriteEndElement()
                                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                                        'данные
                                                        Dim node5 As TreeNodeCollection = userNode4.Nodes
                                                        If node5.Count > 0 Then
                                                            For i5 As Integer = 0 To node5.Count - 1
                                                                Dim userNode5 As TreeNode = node5.Item(i5)
                                                                Dim arrayTag5 As String(,) = userNode5.Tag
                                                                xw.WriteStartElement(arrayTag5(0, 0))
                                                                For n As Integer = 1 To arrayTag5.GetUpperBound(1)
                                                                    xw.WriteAttributeString(arrayTag5(0, n), arrayTag5(1, n))
                                                                Next n
                                                                xw.WriteEndElement()
                                                            Next i5
                                                        End If
                                                    Next i4
                                                End If
                                                xw.WriteEndElement()
                                            Next i3
                                        End If
                                        xw.WriteEndElement()
                                    Next i2
                                End If
                                xw.WriteEndElement()
                            Next i1
                        End If
                    End If
                    xw.WriteEndElement()
                    xw.WriteEndDocument()
                    xw.Flush()
                    xw.Close()
                    xw = Nothing
                End If
            Next i
        End If
        '=========================================================================================
        '    Dim workDirectory As String = Path.GetDirectoryName(templateXML)
        '    If IO.Directory.Exists(workDirectory & "ГруппаБ") = True Then
        '        Dim tempfile As String = Path.GetFileNameWithoutExtension(templateXML)
        '        tempfile = tempfile & rngFolder.Next & ".xml"
        '        IO.File.Copy(templateXML, workDirectory & "\ГруппаБ\" & tempfile, True)
        '    Else
        '        IO.Directory.CreateDirectory(workDirectory & "\ГруппаБ")
        '        Dim tempfile As String = Path.GetFileNameWithoutExtension(templateXML)
        '        tempfile = tempfile & rngFolder.Next & ".xml"
        '        IO.File.Copy(templateXML, workDirectory & "\ГруппаБ\" & tempfile, True)
        '    End If
        '    IO.File.Copy(putchFileXml, templateXML, True)
        'End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'указать слой
    Private Sub ЗадатьСлойToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьСлойToolStripMenuItem.Click
        If IsArray(arrayLayers) = True Then
            Dim formSelect As FormSelectObject = New FormSelectObject
            formSelect.ListBox1.DataSource = arrayLayers
            formSelect.ShowDialog()
            If formSelect.boolShow = True Then
                Dim nameLayer As String = formSelect.ListBox1.SelectedValue
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(0).Value Like "Layer" Then
                        row.Cells(1).Value = nameLayer
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'задать таблицу PS
    Private Sub ЗадатьТаблицуPSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьТаблицуPSToolStripMenuItem.Click
        If IsArray(arrayPS) = True Then
            Dim formSelect As FormSelectObject = New FormSelectObject
            formSelect.ListBox1.DataSource = arrayPS
            formSelect.ShowDialog()
            If formSelect.boolShow = True Then
                Dim namePSTable As String = formSelect.ListBox1.SelectedValue
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(0).Value Like "PS" Then
                        row.Cells(1).Value = namePSTable
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'задать тип линии
    Private Sub ЗадатьТипЛинииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьТипЛинииToolStripMenuItem.Click
        If IsArray(arrayTypeLine) = True Then
            Dim formSelect As FormSelectObject = New FormSelectObject

            formSelect.ListBox1.DataSource = arrayTypeLine
            formSelect.ShowDialog()
            If formSelect.boolShow = True Then
                Dim nameTypeLine As String = formSelect.ListBox1.SelectedValue
                Dim codeTypeLine As String = FuncXML.FuncFindLineTypeCodeByNameTypeLineToXML(templateXML, nameTypeLine)
                'ищем код типа линии и записываем его в таблицу
                If codeTypeLine.Trim.Length > 0 Then
                    For Each row As DataGridViewRow In DataGridView1.Rows
                        If row.Cells(0).Value Like "LineTypeCode" Then
                            row.Cells(1).Value = codeTypeLine
                        End If
                    Next
                End If
                'записываем имя типа линии
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(0).Value Like "LineType" Then
                        row.Cells(1).Value = nameTypeLine
                        Exit Sub
                    End If
                Next
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(0).Value Like "Name" Then
                        row.Cells(1).Value = nameTypeLine
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'задать толщину линию
    Private Sub ЗадатьТолщинуЛинииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьТолщинуЛинииToolStripMenuItem.Click
        'задать толщину линии
        Dim lineWight As Lineweight = Lineweight.lw020
        Dim boolLineWight As Boolean = Topomatic.Dwg.UI.LineweightsBrowserDlg.Execute(lineWight, True)
        If boolLineWight = True Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells(0).Value Like "LineWidth" Then
                    Dim rezStr As String = lineWight.ToString()
                    row.Cells(1).Value = Val(Mid(rezStr, 3))
                    Exit For
                End If
            Next
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор объекта
    Private Sub ВыбратьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыбратьToolStripMenuItem.Click
        boolSelectObject = True
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'закрыть редактор
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim msgboxrez As MsgBoxResult = MsgBox("Вы действительно хотите закрыть редактор? ", MsgBoxStyle.OkCancel)
        If msgboxrez.Ok Then
            Me.Hide()
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'подсказки
    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If (e.ColumnIndex = 0 AndAlso e.Value IsNot Nothing) Then
            Dim cell As DataGridViewCell = Me.DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
            If e.Value.Equals("Description") Then
                cell.ToolTipText = "Описание объекта"
            ElseIf e.Value.Equals("Name") Then
                cell.ToolTipText = "Тип линии, имя блока, штриховки стиля и т.п."
            ElseIf e.Value.Equals("Layer") Then
                cell.ToolTipText = "Слой"
            ElseIf e.Value.Equals("Color") Then
                cell.ToolTipText = "Цвет"
            ElseIf e.Value.Equals("LineTypeCode") Then
                cell.ToolTipText = "Код типа линии"
            ElseIf e.Value.Equals("LineWidth") Then
                cell.ToolTipText = "Толщина линии"
            ElseIf e.Value.Equals("LineScale") Then
                cell.ToolTipText = "Масштаб типа линии"
            ElseIf e.Value.Equals("LineGWidth") Then
                cell.ToolTipText = "Глобальная ширина"
            ElseIf e.Value.Equals("PS") Then
                cell.ToolTipText = "Кодификатор"
            ElseIf e.Value.Equals("Note") Then
                cell.ToolTipText = "Примечание"
            ElseIf e.Value.Equals("Scale") Then
                cell.ToolTipText = "Масштаб объекта"
            ElseIf e.Value.Equals("TextHeight") Then
                cell.ToolTipText = "Высота текста"
            ElseIf e.Value.Equals("TextColor") Then
                cell.ToolTipText = "Цвет текста"
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'задать цвет
    Private Sub ЗадатьЦветToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьЦветToolStripMenuItem.Click
        Dim userColor As CadColor = New CadColor(7)
        Dim boolColorDlg As Boolean = Topomatic.Cad.View.ColorsBrowserDlg.Execute(userColor)
        If boolColorDlg = True Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Cells(0).Value Like "Color" Then
                    row.Cells(1).Value = userColor.ColorIndex
                    Exit For
                End If
            Next
        End If
    End Sub


    Private Sub ЗадатьИмяШтриховкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЗадатьИмяШтриховкиToolStripMenuItem.Click
        Dim pat As IEnumerable(Of HatchPattern) = HatchPatternManager.Current.GetDefinedPatterns()
        Dim arrayHatch As String() = Nothing
        Dim countArrayHatch As Integer = 0
        If pat.Count > 0 Then
            For i As Integer = 0 To pat.Count - 1
                Dim name As String = pat.ElementAt(i).Name
                If IsNothing(name) = False Then
                    If name.Trim.Length > 0 Then
                        ReDim Preserve arrayHatch(countArrayHatch)
                        arrayHatch(countArrayHatch) = name
                        countArrayHatch += 1
                    End If
                End If
            Next i
        End If
        If IsArray(arrayHatch) = True Then
            Dim formSelect As FormSelectObject = New FormSelectObject
            formSelect.ListBox1.DataSource = arrayHatch
            formSelect.ShowDialog()
            If formSelect.boolShow = True Then
                Dim nameHatch As String = formSelect.ListBox1.SelectedValue
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.Cells(0).Value Like "Name" Then
                        row.Cells(1).Value = nameHatch
                        Exit For
                    End If
                Next
            End If
        End If

    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'копироват элемент
    Private Sub КопироватьЭлементToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КопироватьЭлементToolStripMenuItem.Click
        Dim nodeTempClone As TreeNode = TreeView1.SelectedNode
        If IsNothing(nodeTempClone) = False Then
            If IsNothing(nodeTempClone.Parent) = True Then
                MsgBox("Нельзя копировать корневой элемент!!!")
                Exit Sub
            Else
                nodeUserClone = nodeTempClone.Clone
                handleID = nodeTempClone.Parent.Handle
                nameParentUserClone = nodeTempClone.Parent.Name
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вставить элемент
    Private Sub ВставитьЭлементToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВставитьЭлементToolStripMenuItem.Click
        If IsNothing(nodeUserClone) = False Then
            Dim nameNode As String = nodeUserClone.Name
            'Dim parentNameNode As String = nodeUserClone.Parent.Name
            Dim startNodeADSK As TreeNodeCollection = TreeView1.Nodes
            Dim destNodeSelect As TreeNode = TreeView1.SelectedNode
            If nameParentUserClone Like destNodeSelect.Name Then
                Dim countNodes As Integer = destNodeSelect.Nodes.Count
                destNodeSelect.Nodes.Insert(countNodes, nodeUserClone)
            End If
            'Dim firstNode As TreeNode = Nothing
            'While IsNothing(destNodeSelect.Parent)
            '    destNodeSelect = destNodeSelect.Parent
            'End While
            'If IsNothing(destNodeSelect) = False Then
            '    If destNodeSelect.Name Like nameNode Then
            '    End If
            'End If
            'If IsNothing(destNodeSelect.Parent) = False Then
            '    Dim parentDestUserNode As TreeNode = destNodeSelect.Parent
            'End If
        End If
    End Sub
    'удалить элемент
    Private Sub УдалитьЭлементToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьЭлементToolStripMenuItem.Click
        Dim nodeTempClone As TreeNode = TreeView1.SelectedNode
        If IsNothing(nodeTempClone) = False Then
            Dim rez As MsgBoxResult = MsgBox("Вы действительно хотите удалить выбранный элемент? ", MsgBoxStyle.OkCancel, "Удалить элемент")
            If rez.Ok Then
                nodeTempClone.Remove()
            End If
        End If
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub FormEditXMLSymbol_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub
End Class

