Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Environment
Imports System.IO
Imports System.IO.Packaging
Imports System.Media
Imports System.Net
Imports System.Net.Mime.MediaTypeNames
Imports System.Net.Security
Imports System.Runtime.InteropServices.ComTypes
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Forms
Imports System.Windows.Forms.Control
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Media3D
Imports System.Windows.Shapes
Imports System.Xml
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports MS.Internal
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Mathematics
Imports Newtonsoft.Json
Imports Topomatic
Imports Topomatic.Acax.Export
Imports Topomatic.Alg
Imports Topomatic.Alg.Bridges
Imports Topomatic.Alg.Crs
Imports Topomatic.Alg.Layers.Wrappers
Imports Topomatic.Alg.Plan
Imports Topomatic.Alg.Prf
Imports Topomatic.Alg.Road
Imports Topomatic.Alg.Road.Core
Imports Topomatic.Alg.Road.Urb.Border
Imports Topomatic.Alg.Road.Urb.Descent
Imports Topomatic.Alg.Road.Urb.Descent.DescentStrip
Imports Topomatic.Alg.Road.Urb.UserStrips
Imports Topomatic.Alg.Runtime.ServiceClasses
Imports Topomatic.Alg.Runtime.Tools.ProfileVisibleCalculator
Imports Topomatic.Alg.Runtime.Tools.StraighteningSimplePlanSolver
Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Core.ProjectStateConfiguration
Imports Topomatic.ApplicationPlatform.Plugins
Imports Topomatic.Arrangements
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Design
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Cad.View.Tools
Imports Topomatic.Controls.Dialogs
Imports Topomatic.Crs.Templates
Imports Topomatic.Dtm
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Dwg.Layer
Imports Topomatic.Dwg.Layer.DrawingLayer
Imports Topomatic.FoundationClasses
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.FoundationClasses.Undo
Imports Topomatic.FoundationClasses.Vcs
Imports Topomatic.Pipes.Layers.Caches.Plan.Segments.SegmentPlanCache.DrawingData
Imports Topomatic.Pipes.Layers.Profile.StaticNodeDrawers.NodeExtendedHeaderDrawer
Imports Topomatic.Pipes.Runtime.Plt.PlanCross.Fields
Imports Topomatic.Pipes.Runtime.Plt.Profile.DescrEnumConverters
Imports Topomatic.Planchet
Imports Topomatic.Planchet.Entities
Imports Topomatic.Sfc
Imports Topomatic.Sfc.Layer
Imports Topomatic.Sfc.Layer.Wrappers
Imports Topomatic.Sfc.Style
Imports Topomatic.Sites.Core
Imports Topomatic.Smt
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Constructions
Imports Topomatic.Visualization.Geometry
Imports Topomatic.Visualization.LinearSolidBuilder
Imports Topomatic.Visualization.Runtime
Imports Path = System.IO.Path
Imports Vector2D = Topomatic.Cad.Foundation.Vector2D
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D
Namespace RopExample1
    Partial Public Class RopExample1Module
        Inherits Topomatic.ApplicationPlatform.Plugins.PluginInitializator
        Public Overrides Sub Initialize(ByVal factory As PluginFactory)
            MyBase.Initialize(factory)
            'Регистрируем нашу модель в проекте
            factory.RegisterModelEditor("bridge", New ModelEditorInfo("Мостовое сооружение|*.bridgex", ".bridgex", "Мостовое сооружение", "bridge", "CreateModelBridge"))
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'пользовательский элемент управления
        Public userControlPanelPropertiesObject As PanelPropertiesObject = Nothing
        Public userControlPanelProjectBridge As PanelProjectBridge = Nothing
        Public timeModule As Date = New Date
        Public boolIns As String = "Error"

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'активировать лицензию
        <cmd("ActivateModule")>
        Public Sub ActivateModule()
            'находим папку с лицензией
            Dim LicDir As String = GetFolderPath(SpecialFolder.ApplicationData)
            Dim key As String = ""
            Dim nameUser As String = ""
            Dim rngUser As String = ""
            Dim licFile As String = ""
            Dim rng As Random = New Random
            If IO.Directory.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic") = True Then
                If IO.File.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt") = True Then
                    licFile = LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt"
                    Dim input As StreamReader = New StreamReader(licFile, Encoding.GetEncoding(1251), True)
                    Dim countLine As Integer = 0
                    Do Until input.EndOfStream
                        Dim line1 As String = input.ReadLine().Trim 'считываем строку
                        If countLine = 0 Then
                            nameUser = line1
                        ElseIf countLine = 1 Then
                            key = line1
                        ElseIf countLine = 2 Then
                            rngUser = line1
                            Exit Do
                        End If
                        countLine += 1
                    Loop
                    input.Close()
                End If
            End If
            If rngUser.Trim.Length = 0 Then
                rngUser = rng.Next
            End If
            Dim formReg As FormRegistration = New FormRegistration
            formReg.TextBox1.Text = nameUser
            formReg.MaskedTextBox1.Text = key
            formReg.MaskedTextBox1.Tag = rngUser
            formReg.ShowDialog()
            If formReg.boolShow = False Then
                boolIns = False
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем запрос на сайт
            key = formReg.MaskedTextBox1.Text
            nameUser = formReg.TextBox1.Text
            Dim hasp As String = formReg.hasp
            boolIns = PostNEt.PostRequest(key, hasp)
            If boolIns = "{""status"":""ok""}" Then
                timeModule.AddYears(2000)
                MsgBox("Программа успешно активирована!!!")
            Else
                timeModule.AddYears(2000)
                MsgBox("Не удалось активировать программу. Проверьте интернет-соединение и ключ доступа!!!")
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            If IO.Directory.Exists(LicDir & "\Civil3DToolsUtility") = False Then
                IO.Directory.CreateDirectory(LicDir & "\Civil3DToolsUtility")
            End If
            If IO.Directory.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic") = False Then
                IO.Directory.CreateDirectory(LicDir & "\Civil3DToolsUtility\InfrastradaLic")
            End If
            If IO.File.Exists(LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt") = False Then
                Dim filePatchLic As String = LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt"
                Dim fileLic As StreamWriter = New StreamWriter(filePatchLic, False, System.Text.Encoding.Unicode)
                fileLic.WriteLine(nameUser)
                fileLic.WriteLine(key)
                fileLic.WriteLine(rngUser)
                fileLic.Close()
            Else
                Dim filePatchLic As String = LicDir & "\Civil3DToolsUtility\InfrastradaLic\License.txt"
                Dim fileLic As StreamWriter = New StreamWriter(filePatchLic, False, System.Text.Encoding.Unicode)
                fileLic.WriteLine(nameUser)
                fileLic.WriteLine(key)
                fileLic.WriteLine(rngUser)
                fileLic.Close()
            End If
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'установить программу
        <cmd("LoadInfrastrada")>
        Public Sub LoadInfrastrada()
            '========================================================================================
            'загружаем пути к сборкам
            Dim myPath As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выбор папки InfrastradaToolsUtility")
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'прописываем пути к системным каталогами
            Dim patch1 As String = myPath & "\FileResources\Sample"
            Dim patch2 As String = myPath & "\FileResources\SHPForms"
            'Dim patch3 As String = myPath & "\TopomaticRobur\MenuTools"
            Dim bool1 As Boolean = FuncFiles.writeDirectoryToDirectorySupport(patch1)
            Dim bool2 As Boolean = FuncFiles.writeDirectoryToDirectorySupport(patch2)
            If bool2 And bool1 Then
                MsgBox("Пути к вспомогательным файлам поддержки успешно установлены!" & vbLf & patch1 & vbLf & patch2)
            End If
            'Dim bool3 As Boolean = FuncFiles.funcWriteDirectoryToDirectorySupport(patch3)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем шрифты
            Dim errFiles As String() = Nothing
            Dim countErrFiles As Integer = 0
            Dim userFontManager As FontManager = Topomatic.Cad.Foundation.FontManager.Current
            Dim defFileFont As String = userFontManager.DefaultFont.FilePath
            Dim folderSustemDir As String = Path.GetDirectoryName(defFileFont)
            Dim userDirFont As String = myPath & "\Fonts\"
            If IO.Directory.Exists(userDirFont) = True Then
                Dim fileEntries As String() = Directory.GetFiles(userDirFont)
                For Each fileName As String In fileEntries
                    Dim namefilefont As String = Path.GetFileName(fileName)
                    Dim extension As String = Path.GetExtension(fileName)
                    If extension Like ".ttf" Or extension Like ".shx" Or extension Like ".SHX" Then
                        Try
                            File.Copy(fileName, folderSustemDir & "\" & namefilefont, True)
                        Catch ex As System.UnauthorizedAccessException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As System.ArgumentException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As PathTooLongException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As DirectoryNotFoundException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As FileNotFoundException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As IOException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        Catch ex As NotSupportedException
                            ReDim Preserve errFiles(countErrFiles)
                            errFiles(countErrFiles) = fileName
                            countErrFiles += 1
                        End Try
                    End If

                Next
                If IsArray(errFiles) = True Then
                    Dim errMess As String = "Данные шрифты не были установлены в папку " & folderSustemDir
                    For i As Integer = 0 To errFiles.Length - 1
                        errMess = errMess & vbLf & errFiles(i)
                    Next i
                    MsgBox(errMess)
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'копируем формы SHP и SHX
            Dim arrayLookDirectories As String() = Nothing
            Dim boolFolders As Boolean = FuncFiles.readDirectoriesSupport(arrayLookDirectories)
            If IsArray(arrayLookDirectories) = True Then
                For i As Integer = 0 To arrayLookDirectories.Length - 1
                    Erase errFiles
                    countErrFiles = 0
                    Dim tempdir As String = arrayLookDirectories(i)
                    If tempdir Like "C:\ProgramData\Topomatic\Robur road\16.0\Support" Or tempdir Like "C:\ProgramData\Topomatic\Robur road\16.0\Fonts" Then
                        If IO.Directory.Exists(tempdir) = True Then
                            Dim fileEntries As String() = Directory.GetFiles(patch2)
                            For Each fileName As String In fileEntries
                                Dim namefilefont As String = Path.GetFileName(fileName)
                                Dim extension As String = Path.GetExtension(fileName)
                                If extension Like ".shp" Or extension Like ".SHP" Or extension Like ".shx" Or extension Like ".SHX" Then
                                    Try
                                        File.Copy(fileName, folderSustemDir & "\" & namefilefont, True)
                                    Catch ex As System.UnauthorizedAccessException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As System.ArgumentException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As PathTooLongException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As DirectoryNotFoundException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As FileNotFoundException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As IOException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    Catch ex As NotSupportedException
                                        ReDim Preserve errFiles(countErrFiles)
                                        errFiles(countErrFiles) = fileName
                                        countErrFiles += 1
                                    End Try
                                End If
                            Next
                        End If
                    End If
                    If IsArray(errFiles) = True Then
                        Dim errMess As String = "Данные Формы не были установлены в папку " & tempdir
                        For j As Integer = 0 To errFiles.Length - 1
                            errMess = errMess & vbLf & errFiles(j)
                        Next j
                        MsgBox(errMess)
                    End If
                Next i
            End If
        End Sub

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'панель пользовательских свойств
        <cmd("CreatePropertiesObject")>
        Private Function CreatePropertiesObject() As System.Windows.Forms.Control
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!===")
                Exit Function
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                userControlPanelPropertiesObject = New PanelPropertiesObject()
                userControlPanelPropertiesObject.userCadView = cadWiever
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Return userControlPanelPropertiesObject
            End If
        End Function
        'обновление панели пользовательских свойств
        <cmd("RefreshPropertiesObject")>
        Private Function RefreshPropertiesObject() As System.Windows.Forms.Control
            Dim cadWiever As CadView = Me.CadView
            If IsNothing(cadWiever) Then
                Exit Function
            End If
            If Not (ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow) Then
                Exit Function
            End If
            userControlPanelPropertiesObject.userCadView = cadWiever
            If cadWiever.SelectionSet.Count > 0 Then
                For Each acEnt As Object In cadWiever.SelectionSet
                    If TypeOf acEnt Is DwgEntity Then
                        Dim acObj As DwgEntity = acEnt
                        If Not (acObj.ObjectID Like userControlPanelPropertiesObject.ComboBox_DataTable.Tag) Then
                            userControlPanelPropertiesObject.BeginPropertiesGridUpdate()
                            Try
                                userControlPanelPropertiesObject.selectDwgObject = Nothing
                                userControlPanelPropertiesObject.dataStructureElement = Nothing
                                userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Clear()
                                userControlPanelPropertiesObject.ComboBox_DataTable.Tag = ""
                                userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = Nothing
                                userControlPanelPropertiesObject.SetGlobalPropertiesMode()

                                Dim arrayTablesPS As String() = {"PROJECT_STRUCTURES"}
                                If IsArray(arrayTablesPS) = True Then
                                    userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = arrayTablesPS
                                    userControlPanelPropertiesObject.ComboBox_DataTable.Tag = acObj.ObjectID

                                    Dim projectStructuresIndex As Integer = -1
                                    For comboIndex As Integer = 0 To userControlPanelPropertiesObject.ComboBox_DataTable.Items.Count - 1
                                        Dim comboItem As Object = userControlPanelPropertiesObject.ComboBox_DataTable.Items(comboIndex)
                                        If comboItem IsNot Nothing AndAlso String.Equals(comboItem.ToString(), StructureElement.tableXRecords.PROJECT_STRUCTURES.ToString(), StringComparison.OrdinalIgnoreCase) Then
                                            projectStructuresIndex = comboIndex
                                            Exit For
                                        End If
                                    Next comboIndex

                                    If projectStructuresIndex > -1 Then
                                        userControlPanelPropertiesObject.ComboBox_DataTable.SelectedIndex = projectStructuresIndex
                                        Dim dataElement As StructureElement = New StructureElement
                                        Dim boolFindProperties As Boolean = FuncXRecords.getXRecords(acObj, dataElement)
                                        If boolFindProperties = True Then
                                            userControlPanelPropertiesObject.selectDwgObject = acObj
                                            userControlPanelPropertiesObject.dataStructureElement = dataElement
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Add(10)
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(0).Cells(0).Value = "LABEL"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(0).Cells(1).Value = dataElement.Label
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(1).Cells(0).Value = "ClassBridgeObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(1).Cells(1).Value = dataElement.ClassBridgeObject.ToString
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(2).Cells(0).Value = "ClassObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(2).Cells(1).Value = dataElement.ClassObject.ToString
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(3).Cells(0).Value = "Name"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(3).Cells(1).Value = dataElement.Name
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(4).Cells(0).Value = "Description"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(4).Cells(1).Value = dataElement.Description
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(0).Value = "KeyParameter"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(5).Cells(1).Value = dataElement.KeyParameter
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(6).Cells(0).Value = "IdStructure"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(6).Cells(1).Value = dataElement.IdStructure
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(7).Cells(0).Value = "IdElement"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(7).Cells(1).Value = dataElement.IdElement
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(8).Cells(0).Value = "Note"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(8).Cells(1).Value = dataElement.Note
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(9).Cells(0).Value = "IdObject"
                                            userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows(9).Cells(1).Value = acObj.ObjectID.ToString
                                        End If
                                    End If
                                End If
                            Finally
                                userControlPanelPropertiesObject.EndPropertiesGridUpdate()
                            End Try
                        End If
                    End If
                    Exit For
                Next
            Else
                userControlPanelPropertiesObject.BeginPropertiesGridUpdate()
                Try
                    userControlPanelPropertiesObject.selectDwgObject = Nothing
                    userControlPanelPropertiesObject.dataStructureElement = Nothing
                    userControlPanelPropertiesObject.DataGrid_PropertiesEnt.Rows.Clear()
                    userControlPanelPropertiesObject.ComboBox_DataTable.Tag = ""
                    userControlPanelPropertiesObject.ComboBox_DataTable.DataSource = Nothing
                    userControlPanelPropertiesObject.SetGlobalPropertiesMode()
                Finally
                    userControlPanelPropertiesObject.EndPropertiesGridUpdate()
                End Try
            End If
            Return userControlPanelPropertiesObject
        End Function
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'панель редактирования искусственных сооружений
        <cmd("EditProjectBridge")>
        Private Function EditProjectBridge() As System.Windows.Forms.Control
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Function
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                Exit Function
            End If
            If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                userControlPanelProjectBridge = New PanelProjectBridge()
                userControlPanelProjectBridge.CadViewPanel = userCadView
                '===============================================================================================================
                Dim civilEngenProj As ProjectCivilStructures = New ProjectCivilStructures
                Dim listProjectStructure As List(Of ArrangementModel) = civilEngenProj.ListModelStructures
                If listProjectStructure.Count > 0 Then
                    Dim firstProject As ArrangementModel = listProjectStructure.Item(0)
                    'userControlPanelProjectBridge.dictModelProject = firstProject
                    userControlPanelProjectBridge.CBox_ListNamesArrProject.DataSource = civilEngenProj.listNameArrangementModels
                    'userControlPanelProjectBridge.ComboBox1.Tag = firstProject.IProjectModelStructures.Project.TargetProjectUri.AsFilePath
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Return userControlPanelProjectBridge
            End If
        End Function
        'обновление панели проектирования мостов
        <cmd("RefreshPanelProjectBridge")>
        Private Function RefreshPanelProjectBridge() As System.Windows.Forms.Control
            Try
                Dim userCadView As CadView = CadViewDesignUtils.OnCadViewSelect(CadViewDesignUtils.PlanCadViewAlias)
                'Dim userCadView As CadView = Me.CadView
                If IsNothing(userCadView) = True Then Exit Function
                If Not (userCadView Is userControlPanelProjectBridge.CadViewPanel) Then
                    userControlPanelProjectBridge.CadViewPanel = userCadView
                    Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
                    If IsNothing(DrawLayer) = False Then
                        'получаем активный документ
                        Dim ActivDocument As Dwg.Drawing = DrawLayer.Drawing
                        userControlPanelProjectBridge.ActivDocumentPanel = ActivDocument
                    Else
                        Exit Function
                    End If
                End If
                If IsNothing(ApplicationHost.Current.ActiveProject.ActiveDocument) = True Then
                    Exit Function
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ищем узел
                If IsNothing(userCadView) = False Then
                    If userCadView.SelectionSet.Count = 1 Then
                        For Each acEnt As Object In userCadView.SelectionSet
                            If TypeOf acEnt Is DwgEntity Then
                                Dim hgElement As UInteger = acEnt.ObjectID
                                If hgElement Like userControlPanelProjectBridge.selectIDObject Then
                                    Continue For
                                End If
                                If IsNothing(hgElement) = False Then
                                    If hgElement > 0 Then
                                        Dim userTreeView As System.Windows.Forms.TreeView = userControlPanelProjectBridge.TreeView1
                                        Dim foundNode As TreeNode = userControlPanelProjectBridge.FindNodeByTag_ObjectID(userTreeView, hgElement)
                                        If IsNothing(foundNode) = False Then
                                            If IsArray(foundNode.Tag) = True Then
                                                Dim arrayTag As String() = foundNode.Tag
                                                If arrayTag.Length > 3 Then
                                                    Dim nodeObjectID As String = arrayTag(4)
                                                    If Not (nodeObjectID Like acEnt.ObjectID) Then
                                                        Dim nodeColl As TreeNodeCollection = foundNode.Nodes
                                                        If nodeColl.Count > 0 Then
                                                            For Each node As TreeNode In nodeColl
                                                                If IsArray(node.Tag) Then
                                                                    Dim arrayTag1 As String() = node.Tag
                                                                    If arrayTag1.Length > 3 Then
                                                                        Dim nodeObjectID1 As String = arrayTag1(4)
                                                                        If (nodeObjectID1 Like acEnt.ObjectID) Then
                                                                            foundNode = node
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            userTreeView.SelectedNode = foundNode
                                            foundNode.EnsureVisible()
                                            userTreeView.Focus()
                                            Dim args As New TreeNodeMouseClickEventArgs(foundNode, System.Windows.Forms.MouseButtons.Left, 1, 0, 0)
                                            'userControlPanelProjectBridge.TreeView1_NodeMouseClick(userTreeView, args)
                                            userControlPanelProjectBridge.selectIDObject = acEnt.ObjectID
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                If ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.PlanWindow Then

                ElseIf ApplicationHost.Current.ActiveProject.ActiveDocument.UID = Consts.VolumeWindow Then
                    Dim a = 0
                End If
            Catch ex As System.NullReferenceException
            End Try
        End Function

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'разложить балки для всех типов путепроводов
        <cmd("PlacementFixedBeams")>
        Public Sub PlacementFixedBeams()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                MsgBox("Для запуска программы, пожалуйста активируйте окно ПЛАН!")
                Return
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'таблицы кодификатора
            Dim categoryTables As String = "Искусственные сооружения"
            Dim nameTable As String = "Мостовое сооружение"
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'инициализируем новый проект
            Dim projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
            'получаем для данного проекта список подъобъектов
            Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures
            Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
            'штшциализируем форму
            Dim FormBridge As FormPlacementBeams = New FormPlacementBeams()
            FormBridge.civilStructuresProject = projectCivil
            FormBridge.CBox_ListModelStructures.DataSource = projectCivil.listNameArrangementModels() 'проект для раскладки балок
            FormBridge.CBox_ListAxisRoads.DataSource = projectCivil.listNameRoadModels() 'доступные трассы
            FormBridge.CBox_ListProjectSurfaces.DataSource = projectCivil.listNameTerrainModels()  'доступные поверхности
            'получаем шаблон оформления
            Dim dictionaryFilesTemlateXML As Dictionary(Of String, String) = ProjectCivilStructuresStyle.getTemplateXml()
            FormBridge.CBox_ListTemplateXML.DataSource = dictionaryFilesTemlateXML.Keys.ToList()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем мост в первой модели
            Dim projectBridge As ProjectBridge = New ProjectBridge()
            If listProjectStructure.Count > 0 Then
                Dim arrProject As ArrangementModel = projectCivil.getArrangementModelByIndex(0)
                projectBridge.BridgeModel = arrProject
                projectBridge.getBridges()
                Dim listNameBridge As List(Of String) = projectBridge.getNamesBridges()
                FormBridge.CBox_ListNameStructures.DataSource = listNameBridge
                FormBridge.dictionaryBridge = projectBridge.ListBridges
            End If
            '============================================================================================================
            'определяем стиль линии для выбранных траекторий
            Dim styleAxisRowBeams As ProjectCivilStructuresStyle = Nothing
            Dim putchTemlateXML As String = ""
            Dim indexProject As Integer = 0
            Dim projectArrangement As ArrangementModel = Nothing
            Dim userAlign As Alignment = Nothing
            Dim surfaceProject As Surface = Nothing
            '============================================================================================================
            'пеердаем в форму переменные
            FormBridge.civilBridgeProject = projectBridge
LineErr:
            'запускаем форму
            FormBridge.ShowDialog()
            If FormBridge.boolButtonRows = True Then
                userCadView.SelectionSet.Clear()
                CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgPolyline, "Выберите траекторию раскладки балок: ")
                For Each acEnt As DwgEntity In userCadView.SelectionSet
                    If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgPolyline Then
                        Dim acAxisLineElement As DwgPolyline = acEnt.Clone()
                        FormBridge.DG_RowProperties.Rows(FormBridge.numberSelectRows).Cells(2).Value = Math.Round(acAxisLineElement.Length, 3) & " м"
                        FormBridge.DG_RowProperties.Rows(FormBridge.numberSelectRows).Cells(2).Tag = acAxisLineElement
                        FormBridge.boolButtonRows = False
                        GoTo LineErr
                    Else
                        FormBridge.DG_RowProperties.Rows(FormBridge.numberSelectRows).Cells(2).Value = "Ось трассы"
                        FormBridge.DG_RowProperties.Rows(FormBridge.numberSelectRows).Cells(2).Tag = Nothing
                    End If
                Next
            ElseIf FormBridge.boolButtonSelectPillar = True Then
                'выбор оси опоры
                userCadView.SelectionSet.Clear()
                CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите ось опоры: ")
                Dim listCoord As List(Of Vector3D) = New List(Of Vector3D)
                If userCadView.SelectionSet.Count = 0 Then
                    FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Value = "0"
                    FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Tag = Nothing
                End If
                For Each acEnt As DwgEntity In userCadView.SelectionSet
                    If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgLine Then
                        Dim tempLine As DwgLine = acEnt
                        listCoord.Add(tempLine.StartPoint)
                        listCoord.Add(tempLine.EndPoint)
                        FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Value = "Назначена"
                        FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Tag = listCoord
                    Else
                        FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Value = "0"
                        FormBridge.DG_PillarsProperties.Rows(FormBridge.numberSelectRows).Cells(4).Tag = Nothing
                    End If
                Next
                FormBridge.boolButtonSelectPillar = False
                GoTo LineErr
            End If
            If FormBridge.boolShowDlg = False Then
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'основное тело прогаммы
            Dim ScaleK As Double = userCadView.AnnotationScale
            'заново переопределяем путь к шаблону оформления
            If dictionaryFilesTemlateXML.ContainsKey(FormBridge.CBox_ListTemplateXML.Text) = True Then
                putchTemlateXML = dictionaryFilesTemlateXML.Item(FormBridge.CBox_ListTemplateXML.Text)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим документ куда будем раскладывать балки
            indexProject = FormBridge.CBox_ListModelStructures.SelectedIndex
            Dim civilBridgeProject As ProjectBridge = FormBridge.civilBridgeProject
            projectArrangement = civilBridgeProject.getArrangementModelByIndex(indexProject)
            Dim ActivDocument As Dwg.Drawing = projectArrangement.Drawing
            If IsNothing(ActivDocument) = True Then
                MsgBox("Проект для раскладки мостовых балок задан не корректно.")
                Exit Sub
            Else
                projectBridge.ActivDocument = ActivDocument
                projectBridge.BridgeModel = projectArrangement
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim indexSurface As Integer = FormBridge.CBox_ListProjectSurfaces.SelectedIndex
            Dim projectTerrainModel As TerrainModel = Nothing
            If FormBridge.CB_SurfaceFromAlign.Checked = False Then
                projectTerrainModel = projectCivil.getTerrainMidelByIndex(indexSurface)
                Dim listSurf As TransactableList(Of String) = projectArrangement.ProjectSurfacesRelativePaths
                Dim mf As IModelFinder = projectArrangement.ModelFinder
                If IsNothing(mf) = False Then
                    Dim str As String = mf.FindRelativePath(projectTerrainModel)
                    listSurf.Add(str)
                End If
                surfaceProject = projectTerrainModel.Surface
                projectBridge.ProjectSurface = surfaceProject
            End If
            If IsNothing(projectTerrainModel) = True And FormBridge.CB_SurfaceFromAlign.Checked = False Then
                MsgBox("Проектная поверхность не задана.")
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим поверхность и ось
            Dim indexRoad As Integer = FormBridge.CBox_ListAxisRoads.SelectedIndex
            Dim projectRoadModel As RoadModel = projectCivil.getRoadModelByIndex(indexRoad)
            If IsNothing(projectRoadModel) = False Then
                userAlign = projectRoadModel.Alignment
                Dim listSurf As TransactableList(Of String) = projectArrangement.ProjectSurfacesRelativePaths
                Dim mf As IModelFinder = projectArrangement.ModelFinder
                If IsNothing(mf) = False Then
                    Dim str As String = mf.FindRelativePath(projectRoadModel)
                    listSurf.Add(str)
                End If
                Dim l As DwgLayer = projectRoadModel.Drawing.Layers.Item("Треугольники")
                If IsNothing(l) = False Then
                    If l.Visible = False Then
                        l.Visible = True
                    End If
                End If
                If FormBridge.CB_SurfaceFromAlign.Checked = True Then
                    surfaceProject = projectRoadModel.Surface
                    projectBridge.ProjectSurface = surfaceProject
                End If
                projectBridge.ProjectAlignment = userAlign
            Else
                Dim userRoadModel As Topomatic.Alg.Survey.Core.SurveyModel = projectCivil.getAlignSurveyDrawingByIndex(indexSurface)
                If IsNothing(userRoadModel) = False Then
                    userAlign = userRoadModel.Alignment
                    Dim listSurf As TransactableList(Of String) = projectArrangement.ProjectSurfacesRelativePaths
                    Dim mf As IModelFinder = projectArrangement.ModelFinder
                    If IsNothing(mf) = False Then
                        Dim str As String = mf.FindRelativePath(userRoadModel)
                        listSurf.Add(str)
                    End If
                    Dim l As DwgLayer = userRoadModel.Drawing.Layers.Item("Треугольники")
                    If IsNothing(l) = False Then
                        If l.Visible = False Then
                            l.Visible = True
                        End If
                    End If
                    If FormBridge.CB_SurfaceFromAlign.Checked = True Then
                        surfaceProject = userRoadModel.Surface
                        projectBridge.ProjectSurface = surfaceProject
                    End If
                    projectBridge.ProjectAlignment = userAlign
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'значение по умолчанию
            Dim nameBridge As String = FormBridge.CBox_ListNameStructures.Text
            If nameBridge.Trim.Length = 0 Then nameBridge = "Искусственное сооружение"
            Dim putchAlbumBeams As String = FormBridge.CBox_AlbumsBeams.Tag
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'размер сооружения
            Dim dimBridgeLeft As Double = FormBridge.NUpD_dimLeftBridge.Value / 1000
            Dim dimBrigeRight As Double = FormBridge.NUpD_dimRightBridge.Value / 1000
            'смешение осей (лево)
            Dim countRowsLeftBearm As Double = FormBridge.NUpD_CountLeftRows.Value 'количество
            Dim offsetLeftFirstBearm As Double = FormBridge.NumericUpDown3.Value / 1000 'смещение первой оси
            Dim offsetLeftLastBearm As Double = FormBridge.NumericUpDown5.Value / 1000 'смещение последней оси
            If (offsetLeftLastBearm - offsetLeftFirstBearm) <= 0 And FormBridge.NUpD_CountLeftRows.Value > 0 Then
                MsgBox("Расстояние от оси до левой крайней балки должно быть больше чем до оси первой балки!")
                GoTo LineErr
            End If
            'смешение осей (лево)
            Dim countRowsRightBearm As Double = FormBridge.NUpD_CountRightRows.Value 'количество
            Dim offsetRightFirstBearm As Double = FormBridge.NumericUpDown8.Value / 1000 'смещение первой оси
            Dim offsetRightLastBearm As Double = FormBridge.NumericUpDown9.Value / 1000 'смещение последней оси
            If (offsetRightLastBearm - offsetRightFirstBearm) <= 0 And FormBridge.NUpD_CountRightRows.Value > 0 Then
                MsgBox("Расстояние от оси до крайней  правой балки должно быть больше чем до оси первой балки!")
                GoTo LineErr
            End If
            Dim countProlet As Integer = FormBridge.NUpD_CountProlet.Value
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим мост
            Dim indexBridge As Integer = FormBridge.CBox_ListNameStructures.SelectedIndex
            If indexBridge = -1 Then indexBridge = 0
            Dim userBridge As Bridges = projectBridge.getBridgeByIndex(indexBridge)
            Dim elementAxisBridge As StructureElement = projectBridge.getStructureElementBridgeByIndex(indexBridge)
            Dim idBridge As String = projectBridge.getIdStructureByIndex(indexBridge)
            If IsNothing(idBridge) = True Then
                idBridge = Guid.NewGuid.ToString = Guid.NewGuid.ToString
            End If
            If idBridge.Trim.Length = 0 Then
                idBridge = Guid.NewGuid.ToString
            End If
            If IsNothing(userBridge) = True Then
                userBridge = New Bridges
            End If
            Dim dataBridje As StructureElement = projectBridge.getDataBridgeByID(idBridge)
            userBridge.NameBridge = nameBridge
            If FormBridge.CBox_ListPlacementBeams.SelectedIndex = 1 Then
                userBridge.TypeBridge = Bridges.typePlacementBeam.float
            ElseIf FormBridge.CBox_ListPlacementBeams.SelectedIndex = 2 Then
                userBridge.TypeBridge = Bridges.typePlacementBeam.maxClearence
            Else
                userBridge.TypeBridge = Bridges.typePlacementBeam.fixed
            End If
            userBridge.ProletCount = FormBridge.NUpD_CountProlet.Value
            userBridge.centerAxis = FormBridge.ChB_CenterBeam.Checked
            userBridge.LeftRowsCount = FormBridge.NUpD_CountLeftRows.Value
            userBridge.RightRowsCount = FormBridge.NUpD_CountRightRows.Value
            userBridge.LeftStructureWidth = FormBridge.NUpD_dimLeftBridge.Value / 1000
            userBridge.RightStructureWidth = FormBridge.NUpD_dimRightBridge.Value / 1000
            userBridge.TransverseOffset = FormBridge.NUpD_TraverseOffset.Value / 1000
            userBridge.VerticalOffset = FormBridge.NUpD_VerticalOffset.Value / 1000
            userBridge.AlignmentName = FormBridge.CBox_ListAxisRoads.Text
            If FormBridge.CB_SurfaceFromAlign.Checked = False Then
                userBridge.projectSurfaceName = FormBridge.CBox_ListProjectSurfaces.Text
            End If
            'начальный пикет раскладки
            If FormBridge.CheckBox3.Checked = True Then
                Dim startPK As Double
                If FuncFormatZn.TryParsePKText(FormBridge.MaskTB_PK.Text, startPK) Then
                    userBridge.startPlacementPosition = startPK
                End If
            Else
                userBridge.startPlacementPosition = 0
            End If
            'ищем уже существующую трассу автодороги 
            Dim acPlineAlign As DwgPolyline = Nothing
            If IsNothing(elementAxisBridge) = True Then
                elementAxisBridge = Bridges.createBridge()
                idBridge = elementAxisBridge.IdStructure
            Else
                acPlineAlign = elementAxisBridge.DWGEntity
            End If
            If IsNothing(acPlineAlign) = True Then
                acPlineAlign = New DwgPolyline
                elementAxisBridge.DWGEntity = acPlineAlign
            End If
            Dim strGSon As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBridge)
            elementAxisBridge.KeyParameter = strGSon
            Dim boolInsDataPS1 As Boolean = FuncXRecords.setXRecords(acPlineAlign, StructureElement.tableXRecords.PROJECT_STRUCTURES, elementAxisBridge)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'получаем все элементы мостового сооружения
            Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(elementAxisBridge.DWGEntity)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'заполняем словарь с опорами
            'заполняем словарь с осями опор, вписываем туда левый, правый зазор и участок омоноличивания балки
            Dim axisPillarsDictionary As Dictionary(Of Integer, List(Of StructureElement)) = userBridge.getPillars(dictionaryBridgeElements, idBridge)
            '=========================================================================================
            'заполняем таблицу опор
            Dim definitAxisPillar As DwgLine = Nothing
            If axisPillarsDictionary.Count > 0 Then
                For i As Integer = 0 To axisPillarsDictionary.Count - 1
                    Dim numberPillar As Integer = axisPillarsDictionary.ElementAt(i).Key
                    Dim userListPillar As List(Of StructureElement) = axisPillarsDictionary.ElementAt(i).Value
                    Dim dataPillar As StructureElement = userListPillar.Item(1)
                    If IsNothing(dataPillar) = True Then
                        dataPillar = New StructureElement()
                    End If
                    Dim userPillar As Pillar = dataPillar.getPillar()
                    If IsNothing(userPillar) = True Then
                        userPillar = New Pillar
                    End If
                    For j As Integer = 0 To FormBridge.DG_PillarsProperties.RowCount - 1
                        If FormBridge.DG_PillarsProperties.Rows(j).Tag = numberPillar Then
                            userPillar.Defining = FormBridge.DG_PillarsProperties.Rows(j).Cells(0).Value
                            If userPillar.Defining = True Then
                                definitAxisPillar = dataPillar.DWGEntity
                                If ActivDocument.ActiveSpace.Entities.Contains(definitAxisPillar) = False Then
                                    userCadView.SelectionSet.Clear()
                                    CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите предполагаемую ось опоры №" & numberPillar & ". (ОТРЕЗОК):")
                                    Dim boolSelectLine As Boolean = False
                                    For Each acEnt As Object In userCadView.SelectionSet
                                        If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgLine Then
                                            If dataPillar.ClassObject = StructureElement.classStructure.LastPillar Then
                                                definitAxisPillar = acEnt
                                            Else
                                                dataPillar.DWGEntity = definitAxisPillar.Clone()
                                            End If
                                            boolSelectLine = True
                                            Exit For
                                        End If
                                    Next
                                    If boolSelectLine = False Then
                                        MsgBox("Прервано пользователем. Сооружение не построено.")
                                        Exit Sub
                                    End If
                                End If
                            End If
                            userPillar.Clearence = Val(FormBridge.DG_PillarsProperties.Rows(j).Cells(1).Value) / 1000
                            userPillar.RightClearence = Val(FormBridge.DG_PillarsProperties.Rows(j).Cells(2).Value) / 1000
                            If userPillar.RightClearence = 0 Then
                                userPillar.RightClearence = userPillar.Clearence
                            End If
                            If userPillar.Clearence < userPillar.RightClearence Then
                                userPillar.MinClearence = userPillar.Clearence
                            Else
                                userPillar.MinClearence = userPillar.RightClearence
                            End If
                            userPillar.SiteMonolit = Val(FormBridge.DG_PillarsProperties.Rows(j).Cells(3).Value) / 1000
                            userPillar.Number = axisPillarsDictionary.ElementAt(i).Key
                            Dim listCoordAxisPillar As List(Of Vector3D) = FormBridge.DG_PillarsProperties.Rows(j).Cells(4).Tag
                            If IsNothing(listCoordAxisPillar) = False Then
                                If listCoordAxisPillar.Count > 1 Then
                                    Dim acAxisLinePillar As DwgLine = New DwgLine()
                                    acAxisLinePillar.StartPoint = listCoordAxisPillar.Item(0)
                                    acAxisLinePillar.EndPoint = listCoordAxisPillar.Item(1)
                                    dataPillar.DWGEntity = acAxisLinePillar
                                End If
                            End If
                            Dim jsonStr As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
                            dataPillar.KeyParameter = jsonStr
                            userListPillar.Item(1) = dataPillar
                            axisPillarsDictionary.Item(numberPillar) = userListPillar
                            Exit For
                        End If
                    Next j
                Next i
            End If
            'заполняем словарь с траекториями раскладки балок
            Dim globalHOffset As Double = FormBridge.NUpD_TraverseOffset.Value / 1000
            Dim globalVOffset As Double = FormBridge.NUpD_VerticalOffset.Value / 1000
            Dim dictTraectoryPlacementBeams As Dictionary(Of Integer, StructureElement) = New Dictionary(Of Integer, StructureElement)
            If FormBridge.DG_RowProperties.RowCount > 0 Then
                For i As Integer = 0 To FormBridge.DG_RowProperties.RowCount - 1
                    Dim numberRow As Integer = Val(FormBridge.DG_RowProperties.Rows(i).Tag) 'номер ряда
                    Dim offsetHorizontalRow As Double = Val(FormBridge.DG_RowProperties.Rows(i).Cells(0).Value) / 1000 'отступ от оси
                    If numberRow < 0 Then
                        offsetHorizontalRow = -1 * offsetHorizontalRow + globalHOffset
                    Else
                        offsetHorizontalRow = offsetHorizontalRow + globalHOffset
                    End If
                    Dim offsetVerticalRow As Double = Val(FormBridge.DG_RowProperties.Rows(i).Cells(1).Value) / 1000 + globalVOffset 'отступ от поверхности
                    Dim dataTraectory As StructureElement = TrajectoryPlacementBeams.createTrajectoryPlacementBeams(idBridge)
                    Dim userTraectory As TrajectoryPlacementBeams = New TrajectoryPlacementBeams(numberRow, offsetHorizontalRow, offsetVerticalRow, TrajectoryPlacementBeams.TypeTrajectoryPlacementBeams.None)
                    'траектория
                    Dim traectoryObject As Object = FormBridge.DG_RowProperties.Rows(i).Cells(2).Tag
                    Dim poly3d As Polyline3D = New Polyline3D
                    If Not (TypeOf traectoryObject Is DwgPolyline) Then 'ось трассы
                        Dim tempPlineAxisElement As DwgPolyline = FuncAlignment.getPolylineOffsetByAlignment(userAlign, offsetHorizontalRow)
                        userTraectory.TypeTrajectory = TrajectoryPlacementBeams.TypeTrajectoryPlacementBeams.ProjectAlignment
                        dataTraectory.DWGEntity = tempPlineAxisElement
                    Else
                        Dim tempPlineAxisElement As DwgPolyline = traectoryObject
                        userTraectory.TypeTrajectory = TrajectoryPlacementBeams.TypeTrajectoryPlacementBeams.UserPolyline
                        dataTraectory.DWGEntity = tempPlineAxisElement
                    End If
                    Dim keyJson As String = Newtonsoft.Json.JsonConvert.SerializeObject(userTraectory)
                    dataTraectory.KeyParameter = keyJson
                    dictTraectoryPlacementBeams.Add(numberRow, dataTraectory)
                Next i
            End If
            'заполняем словарь с балками
            'If ActivDocument.ActiveSpace.Entities.Contains(acPlineAlign) = False Then
            'ActivDocument.ActiveSpace.Entities.Add(acPlineAlign)
            'End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Dim metodPlastmentBeams As Integer = FormBridge.metodPlacmentBeams
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''если раскладка не фиксированная, то заранее создаем оси опор
            'If Not (userBridge.TypeBridge = Bridges.typePlacementBeam.fixed) Then
            '    Dim styleAxisPillar As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(ActivDocument)
            '    styleAxisPillar.setObjectStyle(putchTemlateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опоры")
            '    For i As Integer = 0 To FormBridge.DG_PillarsProperties.RowCount - 1
            '        Dim listCoordAxisPillar As List(Of Vector3D) = FormBridge.DG_PillarsProperties.Rows(i).Cells(4).Tag
            '        If IsNothing(listCoordAxisPillar) = False Then
            '            If listCoordAxisPillar.Count > 1 Then
            '                Dim numberPillar As Integer = FormBridge.DG_PillarsProperties.Rows(i).Tag
            '                Dim userPillar As Pillar = New Pillar()
            '                userPillar.Number = numberPillar
            '                Dim dataPillar As StructureElement = Nothing
            '                If numberPillar = 1 Or numberPillar = FormBridge.NUpD_CountProlet.Value + 1 Then
            '                    dataPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.LastPillar)
            '                Else
            '                    dataPillar = Pillar.createAxis(idBridge, StructureElement.classStructure.MiddlePillar)
            '                End If
            '                Dim defined As Boolean = FormBridge.DG_PillarsProperties.Rows(i).Cells(0).Value
            '                userPillar.Defining = True
            '                Dim acAxisLinePillar As DwgLine = New DwgLine()
            '                acAxisLinePillar.StartPoint = listCoordAxisPillar.Item(0)
            '                acAxisLinePillar.EndPoint = listCoordAxisPillar.Item(1)
            '                dataPillar.DWGEntity = acAxisLinePillar
            '                Dim strGSONAxis As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
            '                dataPillar.KeyParameter = strGSONAxis
            '                Dim boolInsDataPS2 As Boolean = FuncXRecords.setXRecords(acAxisLinePillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
            '                ActivDocument.ActiveSpace.Entities.Add(acAxisLinePillar)
            '                If IsNothing(styleAxisPillar) = False Then
            '                    styleAxisPillar.setObjectStyle(acAxisLinePillar)
            '                End If
            '            End If
            '        End If
            '    Next
            'End If
            ''получаем все элементы мостового сооружения
            'Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(elementAxisBridge.DWGEntity)
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''заполняем словарь с осями опор, вписываем туда левый, правый зазор и участок омоноличивания балки
            'Dim axisPillarsDictionary As Dictionary(Of Integer, List(Of StructureElement)) = userBridge.getPillars(dictionaryBridgeElements)
            ''=========================================================================================
            ''заполняем таблицу опор
            'If axisPillarsDictionary.Count > 0 Then
            '    For i As Integer = 0 To axisPillarsDictionary.Count - 1
            '        Dim numberPillar As Integer = axisPillarsDictionary.ElementAt(i).Key
            '        Dim userListPillar As List(Of StructureElement) = axisPillarsDictionary.ElementAt(i).Value
            '        Dim dataElement As StructureElement = Pillar.getAxisPillar(dictionaryBridgeElements, numberPillar)
            '        If IsNothing(dataElement) = True Then
            '            dataElement = userListPillar(1)
            '        End If
            '        Dim userPillar As Pillar = dataElement.getPillar()
            '        If IsNothing(userPillar) = True Then
            '            userPillar = New Pillar
            '        End If
            '        userPillar.Clearence = Val(FormBridge.DG_PillarsProperties.Rows(i).Cells(1).Value) / 1000
            '        userPillar.RightClearence = Val(FormBridge.DG_PillarsProperties.Rows(i).Cells(2).Value) / 1000
            '        userPillar.SiteMonolit = Val(FormBridge.DG_PillarsProperties.Rows(i).Cells(3).Value) / 1000
            '        userPillar.Number = axisPillarsDictionary.ElementAt(i).Key
            '        Dim jsonStr As String = Newtonsoft.Json.JsonConvert.SerializeObject(userPillar)
            '        dataElement.KeyParameter = jsonStr
            '        userListPillar.Item(1) = dataElement
            '        axisPillarsDictionary.Item(numberPillar) = userListPillar
            '    Next i
            'End If
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''считываем датагрид с характеристиками рядов (0-номер ряда, 1-отступ от оси ,2-отступ от поверхности,3-handle объекта,4-LocalID траектории)
            ''массив с крайними рядами
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''левый ряд
            'Dim arrayLastAxisBeams As String(,) = Nothing
            'Dim globalHOffset As Double = FormBridge.NUpD_TraverseOffset.Value / 1000
            'Dim globalVOffset As Double = FormBridge.NUpD_VerticalOffset.Value / 1000

            'If FormBridge.DG_RowProperties.RowCount > 0 Then
            '    ReDim Preserve arrayLastAxisBeams(4, 0)
            '    Dim numberRow As Integer = Val(FormBridge.DG_RowProperties.Rows(0).Tag) 'номер ряда
            '    arrayLastAxisBeams(0, 0) = numberRow
            '    Dim offsetRow As Double = Val(FormBridge.DG_RowProperties.Rows(0).Cells(0).Value) / 1000 'отступ от оси
            '    If numberRow < 0 Then
            '        arrayLastAxisBeams(1, 0) = -1 * offsetRow + globalHOffset
            '    Else
            '        arrayLastAxisBeams(1, 0) = offsetRow + globalHOffset
            '    End If
            '    arrayLastAxisBeams(2, 0) = Val(FormBridge.DG_RowProperties.Rows(0).Cells(1).Value) / 1000 + globalVOffset 'отступ от поверхности
            '    'траектория
            '    Dim traectoryObject As Object = FormBridge.DG_RowProperties.Rows(0).Cells(2).Tag
            '    If Not (TypeOf traectoryObject Is DwgPolyline) Then 'ось трассы
            '        arrayLastAxisBeams(3, 0) = ""
            '        arrayLastAxisBeams(4, 0) = ""
            '    Else 'траектория задана
            '        Dim tempPlineAxisElement As DwgPolyline = traectoryObject
            '        If ActivDocument.ActiveSpace.Entities.Contains(tempPlineAxisElement) = False Then
            '            ActivDocument.ActiveSpace.Entities.Add(tempPlineAxisElement)
            '        End If
            '        'добавляем в xrecord информацию о том, что данная полилиния является осью раскладки балок
            '        Dim structElement As StructureElement = New StructureElement()
            '        Dim findStruct As Boolean = FuncXRecords.getXRecords(tempPlineAxisElement, structElement)
            '        If findStruct = False Then
            '            structElement = TrajectoryPlacementBeams.createTrajectoryPlacementBeams(idBridge)
            '            Dim userTraectory As TrajectoryPlacementBeams = New TrajectoryPlacementBeams
            '            userTraectory.numberRows = numberRow
            '            Dim strGSONAxis As String = Newtonsoft.Json.JsonConvert.SerializeObject(userTraectory)
            '            structElement.KeyParameter = strGSONAxis
            '            Dim boolInsDataPS2 As Boolean = FuncXRecords.setXRecords(tempPlineAxisElement, StructureElement.tableXRecords.PROJECT_STRUCTURES, structElement)
            '            Dim styleTrajectoryPlacementBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(ActivDocument)
            '            styleTrajectoryPlacementBeams.setObjectStyle(putchTemlateXML, categoryTables, "Мостовое сооружение", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Траектория раскладки балок")
            '            styleTrajectoryPlacementBeams.setObjectStyle(tempPlineAxisElement)
            '        Else
            '            Dim userTraectory As TrajectoryPlacementBeams = structElement.getTrajectoryPlacementBeams()
            '            userTraectory.numberRows = numberRow
            '        End If
            '        arrayLastAxisBeams(3, 0) = tempPlineAxisElement.ObjectID.ToString
            '        arrayLastAxisBeams(4, 0) = structElement.IdElement
            '    End If
            'End If
            ''правый ряд
            'If FormBridge.DG_RowProperties.RowCount > 1 Then
            '    Dim indLastRow As Integer = FormBridge.DG_RowProperties.RowCount - 1
            '    ReDim Preserve arrayLastAxisBeams(4, 1)
            '    Dim numberRow As Integer = Val(FormBridge.DG_RowProperties.Rows(indLastRow).Tag) 'номер ряда
            '    arrayLastAxisBeams(0, 1) = numberRow
            '    Dim offsetRow As Double = Val(FormBridge.DG_RowProperties.Rows(indLastRow).Cells(0).Value) / 1000 'отступ от оси
            '    If numberRow < 0 Then
            '        arrayLastAxisBeams(1, 1) = -1 * offsetRow + globalHOffset
            '    Else
            '        arrayLastAxisBeams(1, 1) = offsetRow + globalHOffset
            '    End If
            '    arrayLastAxisBeams(2, 1) = Val(FormBridge.DG_RowProperties.Rows(indLastRow).Cells(1).Value) / 1000 + globalVOffset 'отступ от поверхности
            '    'траектория
            '    Dim traectoryObject As Object = FormBridge.DG_RowProperties.Rows(indLastRow).Cells(2).Tag
            '    If Not (TypeOf traectoryObject Is DwgPolyline) Then 'ось трассы
            '        arrayLastAxisBeams(3, 1) = ""
            '        arrayLastAxisBeams(4, 1) = ""
            '    Else
            '        Dim tempPlineAxisElement As DwgPolyline = traectoryObject
            '        If ActivDocument.ActiveSpace.Entities.Contains(tempPlineAxisElement) = False Then
            '            ActivDocument.ActiveSpace.Entities.Add(tempPlineAxisElement)
            '        End If
            '        'добавляем в xrecord информацию о том, что данная полилиния является осью раскладки балок
            '        Dim structElement As StructureElement = New StructureElement()
            '        Dim findStruct As Boolean = FuncXRecords.getXRecords(tempPlineAxisElement, structElement)
            '        If findStruct = False Then
            '            structElement = TrajectoryPlacementBeams.createTrajectoryPlacementBeams(idBridge)
            '            Dim userTraectory As TrajectoryPlacementBeams = New TrajectoryPlacementBeams
            '            userTraectory.numberRows = numberRow
            '            Dim strGSONAxis As String = Newtonsoft.Json.JsonConvert.SerializeObject(userTraectory)
            '            structElement.KeyParameter = strGSONAxis
            '            Dim boolInsDataPS2 As Boolean = FuncXRecords.setXRecords(tempPlineAxisElement, StructureElement.tableXRecords.PROJECT_STRUCTURES, structElement)
            '            Dim styleTrajectoryPlacementBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(ActivDocument)
            '            styleTrajectoryPlacementBeams.setObjectStyle(putchTemlateXML, categoryTables, "Мостовое сооружение", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Траектория раскладки балок")
            '            styleTrajectoryPlacementBeams.setObjectStyle(tempPlineAxisElement)
            '        Else
            '            Dim userTraectory As TrajectoryPlacementBeams = structElement.getTrajectoryPlacementBeams()
            '            userTraectory.numberRows = numberRow
            '        End If
            '        arrayLastAxisBeams(3, 1) = tempPlineAxisElement.ObjectID.ToString
            '        arrayLastAxisBeams(4, 1) = structElement.IdElement
            '    End If
            'End If
            ''массив с промежуточными рядами
            'Dim arrayMiddleAxisBeams As String(,) = Nothing
            'Dim countMiddleAxisBeams As Integer = 0
            'If FormBridge.DG_RowProperties.RowCount > 2 Then
            '    For i As Integer = 1 To FormBridge.DG_RowProperties.RowCount - 2
            '        ReDim Preserve arrayMiddleAxisBeams(4, countMiddleAxisBeams)
            '        Dim numberRow As Integer = Val(FormBridge.DG_RowProperties.Rows(i).Tag) 'номер ряда
            '        arrayMiddleAxisBeams(0, countMiddleAxisBeams) = numberRow
            '        Dim offsetRow As Double = Val(FormBridge.DG_RowProperties.Rows(i).Cells(0).Value) / 1000 'отступ от оси
            '        If numberRow < 0 Then
            '            arrayMiddleAxisBeams(1, countMiddleAxisBeams) = -1 * offsetRow + globalHOffset
            '        Else
            '            arrayMiddleAxisBeams(1, countMiddleAxisBeams) = offsetRow + globalHOffset
            '        End If
            '        arrayMiddleAxisBeams(2, countMiddleAxisBeams) = Val(FormBridge.DG_RowProperties.Rows(i).Cells(1).Value) / 1000 + globalVOffset 'отступ от поверхности
            '        'траектория
            '        Dim traectoryObject As Object = FormBridge.DG_RowProperties.Rows(i).Cells(2).Tag
            '        If Not (TypeOf traectoryObject Is DwgPolyline) Then 'ось трассы
            '            arrayMiddleAxisBeams(3, countMiddleAxisBeams) = ""
            '            arrayMiddleAxisBeams(4, countMiddleAxisBeams) = ""
            '        Else
            '            Dim tempPlineAxisElement As DwgPolyline = traectoryObject
            '            If ActivDocument.ActiveSpace.Entities.Contains(tempPlineAxisElement) = False Then
            '                ActivDocument.ActiveSpace.Entities.Add(tempPlineAxisElement)
            '            End If
            '            'добавляем в xrecord информацию о том, что данная полилиния является осью раскладки балок
            '            Dim structElement As StructureElement = New StructureElement()
            '            Dim findStruct As Boolean = FuncXRecords.getXRecords(tempPlineAxisElement, structElement)
            '            If findStruct = False Then
            '                structElement = TrajectoryPlacementBeams.createTrajectoryPlacementBeams(idBridge)
            '                Dim userTraectory As TrajectoryPlacementBeams = New TrajectoryPlacementBeams
            '                userTraectory.numberRows = numberRow
            '                Dim strGSONAxis As String = Newtonsoft.Json.JsonConvert.SerializeObject(userTraectory)
            '                structElement.KeyParameter = strGSONAxis
            '                Dim boolInsDataPS2 As Boolean = FuncXRecords.setXRecords(tempPlineAxisElement, StructureElement.tableXRecords.PROJECT_STRUCTURES, structElement)
            '                Dim styleTrajectoryPlacementBeams As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(ActivDocument)
            '                styleTrajectoryPlacementBeams.setObjectStyle(putchTemlateXML, categoryTables, "Мостовое сооружение", ProjectCivilStructuresStyle.typeEntity.Полилиния, "Траектория раскладки балок")
            '                styleTrajectoryPlacementBeams.setObjectStyle(tempPlineAxisElement)
            '            Else
            '                Dim userTraectory As TrajectoryPlacementBeams = structElement.getTrajectoryPlacementBeams()
            '                userTraectory.numberRows = numberRow
            '            End If
            '            arrayMiddleAxisBeams(3, countMiddleAxisBeams) = tempPlineAxisElement.ObjectID.ToString
            '            arrayMiddleAxisBeams(4, countMiddleAxisBeams) = structElement.IdElement
            '        End If
            '        countMiddleAxisBeams += 1
            '    Next i
            'End If
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            ''массив с балками
            Dim dictionaryBridgeBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
            For i As Integer = 0 To FormBridge.DG_ProletListBeams.ColumnCount - 1
                Dim indProlet As Integer = Val(FormBridge.DG_ProletListBeams.Columns(i).Tag)
                If indProlet = 0 Then Continue For
                If dictionaryBridgeBeams.ContainsKey(indProlet) = True Then
                    Dim listBeamProlet As Dictionary(Of Integer, StructureElement) = dictionaryBridgeBeams.Item(indProlet)
                    If IsNothing(listBeamProlet) = False Then
                        If listBeamProlet.Count > 0 Then
                            For j As Integer = 0 To FormBridge.DG_ProletListBeams.RowCount - 1
                                Dim indRow As Integer = Val(FormBridge.DG_ProletListBeams.Rows(j).Tag)
                                Dim modelBeam As String = FormBridge.DG_ProletListBeams.Rows(j).Cells(i).Value
                                Dim arrayDataBeam As String() = FormBridge.DG_ProletListBeams.Rows(j).Cells(i).Tag
                                If IsNothing(arrayDataBeam) = False Then
                                    Dim nameAlbumBeam As String = arrayDataBeam(0)
                                    If IsNothing(modelBeam) = False Then
                                        'получаем характеристики балки
                                        Dim fileXML As String = putchAlbumBeams & nameAlbumBeam & "\" & nameAlbumBeam & ".xml"
                                        If File.Exists(fileXML) = True Then
                                            Dim userBeam As BeamI = New BeamI()
                                            userBeam = userBeam.setPropertiesFromXML(fileXML, modelBeam)
                                            userBeam.numberProlet = indProlet
                                            userBeam.numberRow = indRow
                                            userBeam.nameAlbum = nameAlbumBeam
                                            Dim keyParamBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBeam)
                                            Dim dataBeam As StructureElement = BeamI.createAxis(idBridge)
                                            dataBeam.KeyParameter = keyParamBeam
                                            listBeamProlet.Item(indRow) = dataBeam
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    End If
                    dictionaryBridgeBeams.Item(indProlet) = listBeamProlet
                End If
            Next i
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            '3. Выбираем начальное сечение
            'Dim numberDefinedAxisPillars As Integer = 1
            'Dim defAxisPillar As DwgLine = Nothing
            'For i As Integer = 0 To FormBridge.DG_PillarsProperties.RowCount - 1
            '    Dim check As Boolean = FormBridge.DG_PillarsProperties.Rows(i).Cells(0).Value
            '    If check = True Then
            '        numberDefinedAxisPillars = FormBridge.DG_PillarsProperties.Rows(i).Tag
            '        Exit For
            '    End If
            'Next i
            'If axisPillarsDictionary.Count > 0 Then
            '    If axisPillarsDictionary.ContainsKey(numberDefinedAxisPillars) = True Then
            '        Dim listAxisPillars As List(Of StructureElement) = axisPillarsDictionary.Item(numberDefinedAxisPillars)
            '        If listAxisPillars.Count > 2 Then
            '            Dim dataPillar As StructureElement = listAxisPillars.Item(1)
            '            If ActivDocument.ActiveSpace.Entities.Contains(dataPillar.DWGEntity) = False Then
            '                userCadView.SelectionSet.Clear()
            '                CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите предполагаемую ось опоры №" & numberDefinedAxisPillars & ". (ОТРЕЗОК):")
            '                For Each acEnt As Object In userCadView.SelectionSet
            '                    If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgLine Then
            '                        defAxisPillar = acEnt
            '                        dataPillar.DWGEntity = defAxisPillar.Clone()
            '                        Exit For
            '                    End If
            '                Next
            '            End If
            '            Dim defPillar As Pillar = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Pillar)(dataPillar.KeyParameter)
            '            defPillar.Defining = True
            '            Dim strPillarGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(defPillar)
            '            dataPillar.KeyParameter = strPillarGSON
            '            Dim boolRecData As Boolean = FuncXRecords.setXRecords(dataPillar.DWGEntity, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataPillar)
            '        End If
            '    End If
            'End If
            Try
                If IsNothing(projectArrangement) = False Then
                    'удаляем лишние опоры
                    'Dim boolRemoveAxis As Boolean = Pillar.removeAxisPillarFromBridge(userBridge, dictionaryBridgeElements)
                    'Dim boolRemoveBeams As Boolean = BeamI.removeBeamsFromBridge(userBridge, dictionaryBridgeElements)
                    'civilBridgeProject.PlacementBeams(arrayLastAxisBeams, arrayMiddleAxisBeams, dictionaryBridgeBeams, axisPillarsDictionary, dictionaryBridgeElements, acPlineAlign, putchAlbumBeams, putchTemlateXML)

                    dataBridje.IdStructure = idBridge
                    Dim strGSONBridge As String = Newtonsoft.Json.JsonConvert.SerializeObject(userBridge)
                    dataBridje.KeyParameter = strGSONBridge
                    dataBridje.DWGEntity = acPlineAlign
                    civilBridgeProject.PlacementStructureBeams(dataBridje, dictionaryBridgeBeams, definitAxisPillar, axisPillarsDictionary, dictTraectoryPlacementBeams, dictionaryBridgeElements, putchAlbumBeams, putchTemlateXML)
                End If
            Finally
                'If IsNothing(userIProject) = False Then
                'userIProject.UnlockWrite()
                'End If
            End Try
            CadView.Unlock()
            CadView.Invalidate()
            Try
                ApplicationHost.Current.Plugins.Execute("redrawall")
            Catch ex As System.Exception
            End Try
        End Sub
        'удалить мост целиком
        <cmd("RomoveStructures")>
        Public Sub RomoveStructures()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim ent As DwgEntity = Nothing
            userCadView.SelectionSet.Clear()
            CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите один любой объект сооружения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgEntity Then
                    ent = acEnt
                    Exit For
                End If
            Next
            If IsNothing(ent) = True Then
                MsgBox("Объект не выбран. Макро прервано!!!")
                Exit Sub
            End If
            Dim dataObject As StructureElement = New StructureElement
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(ent, dataObject)
            If IsNothing(dataObject) = False Then
                Dim idBridge As String = dataObject.IdStructure
                If IsNothing(idBridge) = False Then
                    If idBridge.Trim.Length > 2 Then
                        ActivDocument = ent.Drawing
                        Dim arrayEnt As DwgEntity() = Nothing
                        Dim countArrayEnt As Integer = 0
                        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                            If TypeOf acEnt Is DwgModel3DElement Or TypeOf acEnt Is DwgLine Or TypeOf acEnt Is DwgPolyline Or TypeOf acEnt Is DwgPolyline3D Or TypeOf acEnt Is DwgHatch Then
                                Dim tempDataObject As StructureElement = Nothing
                                Dim boolTempDataObject As Boolean = FuncXRecords.getXRecords(acEnt, tempDataObject, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                If IsNothing(tempDataObject) = False Then
                                    If tempDataObject.IdStructure Like idBridge Then
                                        ReDim Preserve arrayEnt(countArrayEnt)
                                        arrayEnt(countArrayEnt) = acEnt
                                        countArrayEnt += 1
                                    End If
                                End If
                            End If
                        Next
                        If IsArray(arrayEnt) = True Then
                            ActivDocument.BeginUpdate()
                            Try
                                For i As Integer = 0 To arrayEnt.Length - 1
                                    Dim userEnt As DwgEntity = arrayEnt(i)
                                    ActivDocument.ActiveSpace.Entities.Remove(userEnt)
                                Next
                            Catch ex As Exception
                            Finally
                                ActivDocument.EndUpdate()
                                userCadView.SelectionSet.Clear()
                            End Try
                        End If
                    End If
                End If
            End If
        End Sub
        'удалить раскладку балок
        <cmd("ClearPlacementBeams")>
        Public Sub ClearPlacementBeams()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim ent As DwgEntity = Nothing
            userCadView.SelectionSet.Clear()
            CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите один любой объект сооружения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgEntity Then
                    ent = acEnt
                    Exit For
                End If
            Next
            If IsNothing(ent) = True Then
                MsgBox("Объект не выбран. Макро прервано!!!")
                Exit Sub
            End If
            Dim dataObject As StructureElement = New StructureElement
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(ent, dataObject)
            If IsNothing(dataObject) = False Then
                Dim idBridge As String = dataObject.IdStructure
                If IsNothing(idBridge) = False Then
                    If idBridge.Trim.Length > 2 Then
                        ActivDocument = ent.Drawing
                        Dim arrayEnt As DwgEntity() = Nothing
                        Dim countArrayEnt As Integer = 0
                        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                            Dim tempDataObject As StructureElement = Nothing
                            Dim boolTempDataObject As Boolean = FuncXRecords.getXRecords(acEnt, tempDataObject)
                            If IsNothing(tempDataObject) = False Then
                                If tempDataObject.IdStructure Like idBridge Then
                                    If tempDataObject.Name = StructureElement.typeObject.axisBeam Or tempDataObject.Name = StructureElement.typeObject.axisPillarBeams Or tempDataObject.Name = StructureElement.typeObject.counterBottomBeam Or tempDataObject.Name = StructureElement.typeObject.counterTopBeam Or tempDataObject.Name = StructureElement.typeObject.modelBeam Then
                                        ReDim Preserve arrayEnt(countArrayEnt)
                                        arrayEnt(countArrayEnt) = acEnt
                                        countArrayEnt += 1
                                    End If
                                End If
                            End If
                        Next
                        If IsArray(arrayEnt) = True Then
                            ActivDocument.BeginUpdate()
                            Try
                                For i As Integer = 0 To arrayEnt.Length - 1
                                    Dim userEnt As DwgEntity = arrayEnt(i)
                                    ActivDocument.ActiveSpace.Entities.Remove(userEnt)
                                Next
                            Catch ex As Exception
                            Finally
                                ActivDocument.EndUpdate()
                                userCadView.SelectionSet.Clear()
                            End Try
                        End If
                    End If
                End If
            End If
        End Sub
        'удалить участки омоличивания балок
        <cmd("ClearMonolitSites")>
        Public Sub ClearMonolitSites()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim ent As DwgEntity = Nothing
            userCadView.SelectionSet.Clear()
            CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите один любой объект сооружения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgEntity Then
                    ent = acEnt
                    Exit For
                End If
            Next
            If IsNothing(ent) = True Then
                MsgBox("Объект не выбран. Макро прервано!!!")
                Exit Sub
            End If
            Dim dataObject As StructureElement = New StructureElement
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(ent, dataObject)
            If IsNothing(dataObject) = False Then
                Dim idBridge As String = dataObject.IdStructure
                If IsNothing(idBridge) = False Then
                    If idBridge.Trim.Length > 2 Then
                        ActivDocument = ent.Drawing
                        Dim arrayEnt As DwgEntity() = Nothing
                        Dim countArrayEnt As Integer = 0
                        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                            Dim tempDataObject As StructureElement = Nothing
                            Dim boolTempDataObject As Boolean = FuncXRecords.getXRecords(acEnt, tempDataObject)
                            If IsNothing(tempDataObject) = False Then
                                If tempDataObject.IdStructure Like idBridge Then
                                    If tempDataObject.Name = StructureElement.typeObject.axisSiteMonolitBeams Or tempDataObject.Name = StructureElement.typeObject.counterSiteMonolitBeamsTop Or tempDataObject.Name = StructureElement.typeObject.counterSiteMonolitBeamsBottom Or tempDataObject.Name = StructureElement.typeObject.hatchSiteMonolitBeams Or tempDataObject.Name = StructureElement.typeObject.modelSiteMonolitBeams Then
                                        ReDim Preserve arrayEnt(countArrayEnt)
                                        arrayEnt(countArrayEnt) = acEnt
                                        countArrayEnt += 1
                                    End If
                                End If
                            End If
                        Next
                        If IsArray(arrayEnt) = True Then
                            ActivDocument.BeginUpdate()
                            Try
                                For i As Integer = 0 To arrayEnt.Length - 1
                                    Dim userEnt As DwgEntity = arrayEnt(i)
                                    ActivDocument.ActiveSpace.Entities.Remove(userEnt)
                                Next
                            Catch ex As Exception
                            Finally
                                ActivDocument.EndUpdate()
                                userCadView.SelectionSet.Clear()
                            End Try
                        End If
                    End If
                End If
            End If
        End Sub
        'удалить все опоры
        <cmd("ClearPillarsBridge")>
        Public Sub ClearPillarsBridge()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim ent As DwgEntity = Nothing
            userCadView.SelectionSet.Clear()
            CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите один любой объект сооружения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgEntity Then
                    ent = acEnt
                    Exit For
                End If
            Next
            If IsNothing(ent) = True Then
                MsgBox("Объект не выбран. Макро прервано!!!")
                Exit Sub
            End If
            Dim dataObject As StructureElement = New StructureElement
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(ent, dataObject)
            If IsNothing(dataObject) = False Then
                Dim idBridge As String = dataObject.IdStructure
                If IsNothing(idBridge) = False Then
                    If idBridge.Trim.Length > 2 Then
                        ActivDocument = ent.Drawing
                        Dim arrayEnt As DwgEntity() = {}
                        Dim countArrayEnt As Integer = 0
                        For Each acEnt As DwgEntity In ActivDocument.ActiveSpace.Entities
                            Dim tempDataObject As StructureElement = Nothing
                            Dim boolTempDataObject As Boolean = FuncXRecords.getXRecords(acEnt, tempDataObject, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                            If IsNothing(tempDataObject) = False Then
                                If tempDataObject.IdStructure Like idBridge Then
                                    If tempDataObject.ClassBridgeObject = StructureElement.classBridge.Pillars Then
                                        If Not (tempDataObject.Name = StructureElement.typeObject.axisPillar) Then
                                            ReDim Preserve arrayEnt(countArrayEnt)
                                            arrayEnt(countArrayEnt) = acEnt
                                            countArrayEnt += 1
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If IsArray(arrayEnt) = True Then
                            ActivDocument.BeginUpdate()
                            Try
                                For i As Integer = 0 To arrayEnt.Length - 1
                                    Dim userEnt As DwgEntity = arrayEnt(i)
                                    ActivDocument.ActiveSpace.Entities.Remove(userEnt)
                                Next
                            Catch ex As Exception
                            Finally
                                ActivDocument.EndUpdate()
                            End Try
                        End If
                    End If
                End If
            End If
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создать участок омоноличивания
        <cmd("CreateMonolitSetesBeams")>
        Public Sub CreateMonolitSetesBeams()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                MsgBox("Для запуска программы, пожалуйста активируйте окно ПЛАН!")
                Return
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'таблицы кодификатора
            Dim categoryTables As String = "Искусственные сооружения"
            Dim nameTable As String = "Мостовое сооружение"
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'инициализируем новый проект
            Dim projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
            'получаем для данного проекта список подъобъектов
            Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures
            Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
            'штшциализируем форму
            Dim FormMonolitSites As FormCreateMonolitSitesBeams = New FormCreateMonolitSitesBeams
            FormMonolitSites.civilStructuresProject = projectCivil
            FormMonolitSites.CBox_ListNamesArrProject.DataSource = projectCivil.listNameArrangementModels() 'проект для раскладки балок
            'получаем шаблон оформления
            Dim dictionaryFilesTemlateXML As Dictionary(Of String, String) = ProjectCivilStructuresStyle.getTemplateXml()
            FormMonolitSites.CBox_ListNamesTemplateXML.DataSource = dictionaryFilesTemlateXML.Keys.ToList()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем мост в первой модели
            Dim projectBridge As ProjectBridge = New ProjectBridge()
            If listProjectStructure.Count > 0 Then
                Dim arrProject As ArrangementModel = projectCivil.getArrangementModelByIndex(0)
                projectBridge.BridgeModel = arrProject
                projectBridge.getBridges()
                Dim listNameBridge As List(Of String) = projectBridge.getNamesBridges()
                FormMonolitSites.CBox_ListNamesBridge.DataSource = listNameBridge
                FormMonolitSites.dictionaryBridge = projectBridge.ListBridges
            End If
            '============================================================================================================
            'определяем стиль линии для выбранных траекторий
            Dim styleAxisRowBeams As ProjectCivilStructuresStyle = Nothing
            Dim putchTemlateXML As String = ""
            Dim indexProject As Integer = 0
            Dim projectArrangement As ArrangementModel = Nothing
            Dim userAlign As Alignment = Nothing
            Dim surfaceProject As Surface = Nothing
            '============================================================================================================
            'пеердаем в форму переменные
            FormMonolitSites.civilBridgeProject = projectBridge
            FormMonolitSites.ShowDialog()
            If FormMonolitSites.boolShowDlg = False Then
                Exit Sub
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'основное тело прогаммы
            Dim ScaleK As Double = userCadView.AnnotationScale
            'заново переопределяем путь к шаблону оформления
            If dictionaryFilesTemlateXML.ContainsKey(FormMonolitSites.CBox_ListNamesTemplateXML.Text) = True Then
                putchTemlateXML = dictionaryFilesTemlateXML.Item(FormMonolitSites.CBox_ListNamesTemplateXML.Text)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим документ куда будем раскладывать участки омоноличивания
            indexProject = FormMonolitSites.CBox_ListNamesArrProject.SelectedIndex
            Dim civilBridgeProject As ProjectBridge = FormMonolitSites.civilBridgeProject
            projectArrangement = civilBridgeProject.getArrangementModelByIndex(indexProject)
            Dim ActivDocument As Dwg.Drawing = projectArrangement.Drawing
            If IsNothing(ActivDocument) = True Then
                MsgBox("Проект для раскладки мостовых балок задан не корректно.")
                Exit Sub
            Else
                projectBridge.ActivDocument = ActivDocument
                projectBridge.BridgeModel = projectArrangement
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим мост
            Dim indexBridge As Integer = FormMonolitSites.CBox_ListNamesBridge.SelectedIndex
            Dim userBridge As Bridges = projectBridge.getBridgeByIndex(indexBridge)
            Dim elementAxisBridge As StructureElement = projectBridge.getStructureElementBridgeByIndex(indexBridge)
            Dim idBridge As String = projectBridge.getIdStructureByIndex(indexBridge)
            If IsNothing(idBridge) = True Then
                idBridge = Guid.NewGuid.ToString = Guid.NewGuid.ToString
            End If
            If idBridge.Trim.Length = 0 Then
                idBridge = Guid.NewGuid.ToString
            End If
            If IsNothing(userBridge) = True Then
                userBridge = New Bridges
            End If
            'ищем уже существующую трассу автодороги 
            Dim acPlineAlign As DwgPolyline = elementAxisBridge.DWGEntity
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'получаем все элементы мостового сооружения
            Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(elementAxisBridge.DWGEntity)
            'получаем все балки
            Dim dictionaryBridgeBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)
            If FormMonolitSites.RadioButton2.Checked = True Then
                Dim numberProlet As Integer = Val(FormMonolitSites.CBox_ListProlet.Text)
                If numberProlet > 0 Then
                    For i As Integer = 0 To dictionaryBridgeBeams.Count - 1
                        Dim tempNumberProlet As Integer = dictionaryBridgeBeams.ElementAt(i).Key
                        If tempNumberProlet <> numberProlet Then
                            dictionaryBridgeBeams.Remove(tempNumberProlet)
                        End If
                    Next i
                End If
            End If
            Dim typeMonolitPr As Boolean = True
            If FormMonolitSites.RadioButton1.Checked = True Then
                typeMonolitPr = False
            End If
            userBridge.CreateMonolithingBeams(ActivDocument, dictionaryBridgeBeams, idBridge, dictionaryBridgeElements, True, putchTemlateXML, typeMonolitPr)
            'If FormMonolitSites.CBox_TypeMonolitSites.Text Like "Балки" Then
            '    
            'Else
            'Dim typeMonolitPr As Boolean = True
            'If formMonolitSetes.RadioButton4.Checked = False Then
            '    typeMonolitPr = False
            'End If
            'Dim arrayPillarsNumber As Integer() = {}
            'If formMonolitSetes.RadioButton2.Checked = True Then
            '    ReDim arrayPillarsNumber(0)
            '    arrayPillarsNumber(0) = Val(formMonolitSetes.ComboBox4.Text)
            'Else
            '    Dim tempPillars As String() = formMonolitSetes.ComboBox4.DataSource
            '    If IsArray(tempPillars) = True Then
            '        For i As Integer = 0 To tempPillars.Length - 1
            '            ReDim Preserve arrayPillarsNumber(i)
            '            arrayPillarsNumber(i) = Val(tempPillars(i))
            '        Next
            '    End If
            'End If
            'panelFunc.CreateMonolithingPillar(ActivDocument, arrayPillarsNumber, idBridge, templateXML, thickness, typeMonolitPr)
            'End If
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создать крайнюю опору
        <cmd("CreateLastPillarsBridge")>
        Public Sub CreateLastPillarsBridge()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                MsgBox("Для запуска программы, пожалуйста активируйте окно ПЛАН!")
                Return
            End If
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'инициализируем новый проект
            Dim projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
            'получаем для данного проекта список подъобъектов
            Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures 'модели
            Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listEgSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
            'штшциализируем форму
            Dim FormCreateLastPillar As FormCreateLastPillars = New FormCreateLastPillars
            FormCreateLastPillar.civilStructuresProject = projectCivil
            FormCreateLastPillar.CBox_ListNamesArrProject.DataSource = projectCivil.listNameArrangementModels() 'проект для раскладки балок
            FormCreateLastPillar.CB_ProjectSurface.DataSource = projectCivil.listNameTerrainModels 'проектные поверхности
            FormCreateLastPillar.CB_EgSurface.DataSource = projectCivil.listNameTerrainModels 'фактические поверхности
            FormCreateLastPillar.CB_NameAlignment.DataSource = projectCivil.listNameRoadModels 'оси трассы
            'получаем шаблон оформления
            Dim dictionaryFilesTemlateXML As Dictionary(Of String, String) = ProjectCivilStructuresStyle.getTemplateXml()
            FormCreateLastPillar.CBox_ListNamesTemplateXML.DataSource = dictionaryFilesTemlateXML.Keys.ToList()
            FormCreateLastPillar.ShowDialog()
            If FormCreateLastPillar.boolShow = False Then
                Exit Sub
            End If
            Try
                ApplicationHost.Current.Plugins.Execute("redrawall")
            Catch ex As System.Exception
            End Try
        End Sub
        'создать промежуточную опору
        <cmd("CreateMiddlePillarsBridge")>
        Public Sub CreateMiddlePillarsBridge()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                MsgBox("Для запуска программы, пожалуйста активируйте окно ПЛАН!")
                Return
            End If
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'инициализируем новый проект
            Dim projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
            'получаем для данного проекта список подъобъектов
            Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures 'модели
            Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listEgSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
            'штшциализируем форму
            Dim FormCreateMiddlePillar As FormCreateMiddlePillars = New FormCreateMiddlePillars
            FormCreateMiddlePillar.civilStructuresProject = projectCivil
            FormCreateMiddlePillar.CBox_ListNamesArrProject.DataSource = projectCivil.listNameArrangementModels() 'проект для раскладки балок
            FormCreateMiddlePillar.CB_ProjectSurface.DataSource = projectCivil.listNameTerrainModels 'проектные поверхности
            FormCreateMiddlePillar.CB_EgSurface.DataSource = projectCivil.listNameTerrainModels 'фактические поверхности
            FormCreateMiddlePillar.CB_NameAlignment.DataSource = projectCivil.listNameRoadModels 'оси трассы
            'получаем шаблон оформления
            Dim dictionaryFilesTemlateXML As Dictionary(Of String, String) = ProjectCivilStructuresStyle.getTemplateXml()
            FormCreateMiddlePillar.CBox_ListNamesTemplateXML.DataSource = dictionaryFilesTemlateXML.Keys.ToList()
            FormCreateMiddlePillar.ShowDialog()
            If FormCreateMiddlePillar.boolShow = False Then
                Exit Sub
            End If
            Try
                ApplicationHost.Current.Plugins.Execute("redrawall")
            Catch ex As System.Exception
            End Try
        End Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создать конус мостоаого сооружения
        <cmd("CreateConesPillars")>
        Public Sub CreateConesPillars()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then
                MsgBox("Для запуска программы, пожалуйста активируйте окно ПЛАН!")
                Return
            End If
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then
                MsgBox("Для запуска программы, пожалуйста сделайте активируйте модель!")
                Return
            End If
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'Получение всех моделей проекта
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'инициализируем новый проект
            Dim projectCivil As ProjectCivilStructures = New ProjectCivilStructures()
            'получаем для данного проекта список подъобъектов
            Dim listProjectStructure As List(Of ArrangementModel) = projectCivil.ListModelStructures 'модели
            Dim listProjectSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listEgSurface As List(Of TerrainModel) = projectCivil.ListModelSurfaces
            Dim listProjectRoads As List(Of RoadModel) = projectCivil.ListModelRoads
            'штшциализируем форму
            Dim FormCreateConePillar As FormCreateConeLastPillars = New FormCreateConeLastPillars
            FormCreateConePillar.civilStructuresProject = projectCivil
            FormCreateConePillar.CBox_ListNamesArrProject.DataSource = projectCivil.listNameArrangementModels() 'проект для раскладки балок
            FormCreateConePillar.CB_ProjectSurface.DataSource = projectCivil.listNameTerrainModels 'проектные поверхности
            FormCreateConePillar.CB_EgSurface.DataSource = projectCivil.listNameTerrainModels 'фактические поверхности
            FormCreateConePillar.CB_NameAlignment.DataSource = projectCivil.listNameRoadModels 'оси трассы
            FormCreateConePillar.CB_NameSites.DataSource = projectCivil.listNameSitesModels 'площадки
            'получаем шаблон оформления
            Dim dictionaryFilesTemlateXML As Dictionary(Of String, String) = ProjectCivilStructuresStyle.getTemplateXml()
            FormCreateConePillar.CBox_ListNamesTemplateXML.DataSource = dictionaryFilesTemlateXML.Keys.ToList()
Line1:
            FormCreateConePillar.ShowDialog()
            If FormCreateConePillar.boolShow = False Then
                Exit Sub
            End If
        End Sub


        'удалить мост целиком
        <cmd("TestPl")>
        Public Sub TestPl()
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверка лицензии
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim activeData As Date = New Date
            activeData = Date.Now
            Dim countMinute As Long = DateDiff("n", activeData, timeModule)
            If Math.Abs(countMinute) > 61 Then
                Dim key As String = ""
                Dim userName As String = ""
                Dim hasp As String = PostNEt.FuncReadDataLicFile(key, userName)
                If hasp.Length > 0 And key.Length > 0 And userName.Length > 0 Then
                    boolIns = PostNEt.PostRequest(key, hasp)
                    If boolIns = "{""status"":""ok""}" Then
                        timeModule = Date.Now
                    Else
                        MsgBox("Лицензия не доступна!!!")
                        boolIns = "Error"
                    End If
                Else
                    MsgBox("Лицензия не доступна!!!")
                    boolIns = "Error"
                End If
            End If
            If boolIns = "Error" Then
                MsgBox("Лицензия не доступна!!!")
                Exit Sub
            End If
            'получаем активное окно
            Dim userCadView As CadView = Me.CadView
            If IsNothing(userCadView) Then Return
            'получаем активный слой
            Dim DrawLayer As DrawingLayer = DrawingLayer.GetDrawingLayer(userCadView)
            If IsNothing(DrawLayer) = True Then Return
            'получаем активный документ
            Dim ActivDocument As Topomatic.Dwg.Drawing = DrawLayer.Drawing
            If IsNothing(ActivDocument) = True Then Return
            Dim ent As DwgEntity = Nothing
            userCadView.SelectionSet.Clear()
            CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgEntity, "Выберите ось сооружения: ")
            For Each acEnt As Object In userCadView.SelectionSet
                If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgEntity Then
                    ent = acEnt
                    Exit For
                End If
            Next
            If IsNothing(ent) = True Then
                MsgBox("Объект не выбран. Макро прервано!!!")
                Exit Sub
            End If
            Dim dataObject As StructureElement = New StructureElement
            Dim boolFindData As Boolean = FuncXRecords.getXRecords(ent, dataObject)
            If IsNothing(dataObject) = False Then
                Dim idBridge As String = dataObject.IdStructure
                'получаем все элементы мостового сооружения
                Dim userBridge As Bridges = dataObject.getBridge()
                Dim dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)) = userBridge.getBridgeObjects(dataObject.DWGEntity)
                Dim dictBeams As Dictionary(Of Integer, Dictionary(Of Integer, StructureElement)) = userBridge.getBeams(dictionaryBridgeElements)

                Dim entBridgeObject As DwgLine = Nothing
                userCadView.SelectionSet.Clear()
                CadView.SelectionSet.SelectOneObjectAtScreen(Function(obj) TypeOf obj Is Topomatic.Dwg.Entities.DwgLine, "Выберите ось: ")
                For Each acEnt As Object In userCadView.SelectionSet
                    If TypeOf acEnt Is Topomatic.Dwg.Entities.DwgLine Then
                        entBridgeObject = acEnt
                        Exit For
                    End If
                Next
                'полчаем трассу
                Dim nameAlign As String = userBridge.AlignmentName
                Dim userAlign As Alignment = Nothing
                Dim boolFindAlign As Boolean = FuncAlignment.getAlignmentByName(nameAlign, userAlign)
                Dim nameSurface As String = userBridge.projectSurfaceName
                Dim userSurface As Surface = Nothing
                If nameSurface.Trim.Length > 0 Then
                    userSurface = FuncSurface.getSurfaceByName(nameSurface)
                End If
                If IsNothing(userSurface) = True Then
                    userSurface = FuncAlignment.getSurfaceToAlignment(nameAlign)
                End If
                'делаем расчет зазоров
                Dim dictRowBeam As Dictionary(Of Integer, String) = userBridge.getConditionalRows()
                Dim dictAxisPillar As Dictionary(Of Integer, List(Of StructureElement)) = userBridge.getPillars(dictionaryBridgeElements)
                For i As Integer = 0 To dictRowBeam.Count - 1
                    Dim numberRow As Integer = dictRowBeam.ElementAt(i).Key
                    'получаем крайние ряды балок
                    Dim listLastBeams As List(Of StructureElement) = CalculationBeams.getBeamsToRow(dictBeams, numberRow)
                    Dim rowBeamsCalculate As Dictionary(Of Integer, StructureElement) = dictBeams.ElementAt(i).Value
                    'Dim clearence = CalculationBeams.calculatePlacementBeams(listLastBeams, entBridgeObject, dictAxisPillar, userAlign, userSurface, userBridge)
                Next i

            End If
        End Sub
    End Class
End Namespace
