Imports System.IO
Imports System.Windows.Forms
Public Class FuncFiles
    'поиск файла во вспомогательных папках поддержки
    Public Shared Function getFileToDirectorySupport(ByVal shortNameFile As String) As String
        Dim result As String = ""
        Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
        Dim fileShpUser As String() = pr.SearchFiles("LookInFolders", shortNameFile)
        If IsNothing(fileShpUser) = False Then
            If fileShpUser.Length > 0 Then
                For i As Integer = 0 To fileShpUser.Length - 1
                    Dim tempFile As String = fileShpUser(i)
                    If File.Exists(tempFile) = True Then
                        result = tempFile
                    End If
                Next i
            End If
        End If
        Return result
    End Function
    'поиск каталога во вспомогательных папках поддержки
    Public Shared Function getFindDirectorySupport(ByVal directory As String) As String
        Dim result As String = ""
        Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
        Dim fileShpUser As String() = pr.GetDirectories("LookInFolders")
        If IsNothing(fileShpUser) = False Then
            If fileShpUser.Length > 0 Then
                For i As Integer = 0 To fileShpUser.Length - 1
                    Dim tempDirectory As String = fileShpUser(i)
                    Dim pos As Integer = tempDirectory.LastIndexOf(directory)
                    If pos > -1 Then
                        result = tempDirectory
                    End If
                Next i
            End If
        End If
        Return result
    End Function
    'записать каталог в файлы поддержки
    Public Shared Function writeDirectoryToDirectorySupport(ByVal putchLookDirectory As String) As Boolean
        Dim result As Boolean = False
        Try
            If Directory.Exists(putchLookDirectory) = True Then
                Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
                Dim nameDirectories As String() = pr.GetDirectories("LookInFolders")
                Dim boolWrite As Boolean = False
                If IsArray(nameDirectories) = True Then
                    For i As Integer = 0 To nameDirectories.Length - 1
                        Dim tempDir As String = nameDirectories(i)
                        If tempDir Like putchLookDirectory Then
                            boolWrite = True
                            Exit For
                        End If
                    Next
                End If
                If boolWrite = False Then
                    Dim userModif As Topomatic.ApplicationEnvironment.EnvironmentModifier = Topomatic.ApplicationEnvironment.EnvironmentModifier.Process
                    pr.Insert(0, userModif, "LookInFolders", putchLookDirectory)
                End If
                result = True
            End If
        Catch ex As Exception
        End Try
        Return result
    End Function
    'поиск определенной директории в файлах поддержки
    Public Shared Function readDirectoriesSupport(ByRef nameDirectories As String()) As Boolean
        Dim result As Boolean = False
        Try
            Dim pr As Topomatic.ApplicationEnvironment.ProcessEnvironment = Topomatic.ApplicationEnvironment.ProcessEnvironment.Current
            nameDirectories = pr.GetDirectories("LookInFolders")
            If nameDirectories.Length > 0 Then
                result = True
            End If
        Catch ex As Exception
        End Try
        Return result
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'вывод диалогового окна выбор каталога
    Public Shared Function FuncGetAbsoluteFolderPathName(ByVal strTitle As String, Optional lngRegim As Long = 1) As String
        ' Определяем переменные
        Dim objShell As Object
        Dim objFolder As Object
        Dim objFolderItem As Object
        ' Создаём переменную Shell
        objShell = CreateObject("Shell.Application")
        ' Выводим диалоговое окно выбора папки с нужными параметрами
        objFolder = objShell.BrowseForFolder(0, strTitle, lngRegim)
        ' Если объект создан не удачно (НЕ выбрали какую-то папку), то
        ' возвращаем пустую строку в качестве результата работы функции...
        If objFolder Is Nothing Then
            FuncGetAbsoluteFolderPathName = ""
            Exit Function
        End If
        ' Получаем объект, у которого "можно спросить" его path
        objFolderItem = objFolder.Self
        ' Получаем значение Path
        FuncGetAbsoluteFolderPathName = objFolderItem.Path
        ' Удаляем все использованные объекты
        objFolderItem = Nothing
        objFolder = Nothing
        objShell = Nothing
    End Function

    Public Shared Function FuncGetAbsoluteFilePathName(ByVal strTitle As String, Optional ByVal strFilter As String = "All files (*.*)|*.*") As String
        '"txt files (*.txt)|*.txt|All files (*.*)|*.*";
        FuncGetAbsoluteFilePathName = ""
        Dim openFileDlg As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        openFileDlg.InitialDirectory = "c:\\"
        openFileDlg.Filter = strFilter
        openFileDlg.RestoreDirectory = True
        If openFileDlg.ShowDialog() = DialogResult.OK Then
            Return openFileDlg.FileName
        End If
    End Function

    Public Shared Function FuncSaveFilePathName(ByVal strTitle As String, Optional ByVal strFilter As String = "All files (*.*)|*.*") As String
        '"txt files (*.txt)|*.txt|All files (*.*)|*.*";
        FuncSaveFilePathName = ""
        Dim saveFileDlg As System.Windows.Forms.SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        saveFileDlg.InitialDirectory = "c:\\"
        saveFileDlg.Filter = strFilter
        saveFileDlg.RestoreDirectory = True
        If saveFileDlg.ShowDialog() = DialogResult.OK Then
            Return saveFileDlg.FileName
        End If
    End Function

    Public Shared Function IsFolderInPath(filePath As String, folderName As String) As Boolean
        If String.IsNullOrEmpty(filePath) OrElse String.IsNullOrEmpty(folderName) Then
            Return False
        End If

        ' Приводим к одному регистру для регистронезависимого поиска
        Dim normalizedPath As String = filePath.ToLower()
        Dim normalizedFolder As String = folderName.ToLower()

        Return normalizedPath.Contains("\" & normalizedFolder & "\") OrElse
               normalizedPath.StartsWith(normalizedFolder & "\") OrElse
               normalizedPath.EndsWith("\" & normalizedFolder)
    End Function
    Public Shared Function TrimPathAfter(ByVal fullPath As String, Optional ByVal nameFolderTrim As String = "InfrastradaToolsUtility") As String
        ' Проверка на пустую строку
        If String.IsNullOrWhiteSpace(fullPath) Then
            Return String.Empty
        End If


        ' Ищем позицию "TopomaticRobur" в пути
        Dim index As Integer = fullPath.IndexOf(nameFolderTrim, StringComparison.OrdinalIgnoreCase)

        ' Если подстрока найдена
        If index >= 0 Then
            ' Находим позицию конца слова "TopomaticRobur"
            Dim endIndex As Integer = index + nameFolderTrim.Length

            ' Обрезаем строку до конца слова TopomaticRobur
            Dim result As String = fullPath.Substring(0, endIndex)

            Return result
        Else
            ' Если подстрока не найдена, возвращаем исходный путь
            Return fullPath
        End If
    End Function
    Public Shared Function TrimPath(fullPath As String, Optional ByVal nameFolderTrim As String = "InfrastradaToolsUtility") As String
        Dim index As Integer = fullPath.IndexOf(nameFolderTrim, StringComparison.OrdinalIgnoreCase)

        If index >= 0 Then
            ' Возвращаем путь от TopomaticRobur до конца
            Return fullPath.Substring(index)
        Else
            ' Если не найдено, возвращаем пустую строку или исходный путь
            Return String.Empty
        End If
    End Function


    Public Shared Function IsPathContainsFolder(ByVal filePath As String, ByVal targetFolder As String) As Boolean
        If String.IsNullOrEmpty(filePath) OrElse String.IsNullOrEmpty(targetFolder) Then
            Return False
        End If
        Try
            ' Нормализация путей
            Dim normalizedFilePath As String = filePath.Trim().ToLower()
            Dim normalizedTargetFolder As String = targetFolder.Trim().ToLower()

            ' Замена обратных слешей на прямые для единообразия
            normalizedFilePath = normalizedFilePath.Replace("\", "/")
            normalizedTargetFolder = normalizedTargetFolder.Replace("\", "/")

            ' Удаление лишних слешей в начале и конце
            normalizedTargetFolder = normalizedTargetFolder.Trim("/"c)

            ' Если целевая папка пуста после нормализации
            If normalizedTargetFolder.Length = 0 Then
                Return True
            End If

            ' Получение директории файла
            Dim directoryPath As String = ""
            If IO.File.Exists(filePath) Then
                directoryPath = IO.Path.GetDirectoryName(filePath)
            Else
                directoryPath = filePath
            End If

            ' Нормализация директории
            directoryPath = directoryPath.Replace("\", "/").ToLower()
            directoryPath = directoryPath.Trim("/"c)

            ' Разбиваем путь на компоненты
            Dim pathComponents As String() = directoryPath.Split("/"c)
            Dim targetComponents As String() = normalizedTargetFolder.Split("/"c)

            ' Поиск последовательности компонентов целевой папки в пути
            For i As Integer = 0 To pathComponents.Length - targetComponents.Length
                Dim match As Boolean = True

                For j As Integer = 0 To targetComponents.Length - 1
                    If pathComponents(i + j) <> targetComponents(j) Then
                        match = False
                        Exit For
                    End If
                Next

                If match Then
                    Return True
                End If
            Next

            ' Альтернативная простая проверка (менее точная, но быстрая)
            If directoryPath.Contains(normalizedTargetFolder) Then
                ' Дополнительная проверка, чтобы избежать частичных совпадений
                Dim index As Integer = directoryPath.IndexOf(normalizedTargetFolder)
                If index >= 0 Then
                    ' Проверяем, что это действительно полное имя папки, а не часть другого слова
                    If (index = 0 OrElse directoryPath(index - 1) = "/"c) AndAlso
                       (index + normalizedTargetFolder.Length = directoryPath.Length OrElse
                        directoryPath(index + normalizedTargetFolder.Length) = "/"c) Then
                        Return True
                    End If
                End If
            End If

        Catch ex As Exception
            ' Debug.WriteLine($"Ошибка в IsPathContainsFolder: {ex.Message}")
        End Try

        Return False
    End Function

End Class
