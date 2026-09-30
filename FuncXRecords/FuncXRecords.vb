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
                                    ElseIf nameF Like "ClassBridgeObject" Then
                                        xData.ClassBridgeObject = CType([Enum].Parse(GetType(StructureElement.classBridge), valF), StructureElement.classBridge)
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
            If nameTable = StructureElement.tableXRecords.PROJECT_STRUCTURES Then
                BridgeDrawingGroupManager.TryGroupEntity(entity)
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
        SetStringIfNotEmpty(fieldDictionary, "ClassBridgeObject", xData.ClassBridgeObject)
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
End Class
