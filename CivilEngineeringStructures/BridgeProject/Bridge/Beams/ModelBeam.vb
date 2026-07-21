Imports System.ComponentModel
Imports System.IO
Imports CivilEnginStructures.CounterBeam
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime

Public Class ModelBeam
    Private _numberProlet As Integer 'номер пролета
    Private _numberRow As Integer 'номер ряда
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
    End Sub
    Public Sub New(numberProlet As Integer, numberRow As Integer)
        _numberProlet = numberProlet
        _numberRow = numberRow
    End Sub
    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер пролета")>
    <[ReadOnly](True)>
    Public Property numberProlet() As Integer
        Get
            Return _numberProlet
        End Get
        Set(value As Integer)
            _numberProlet = value
        End Set
    End Property
    <Browsable(True)>
    <Description("Номер ряда")>
    <Category("Свойства сооружения")>
    <DisplayName("Номер ряда")>
    <[ReadOnly](True)>
    Public Property numberRow() As Integer
        Get
            Return _numberRow
        End Get
        Set(value As Integer)
            _numberRow = value
        End Set
    End Property
    'создать новую структуру
    Public Shared Function createModelBeamI(ByVal idBridge As String) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementCounter.ClassObject = StructureElement.classStructure.BeamI
        elementCounter.Name = StructureElement.typeObject.modelBeam
        Dim deskBeam As String = StructureElement.GetDescription(StructureElement.typeObject.modelBeam)
        elementCounter.Description = deskBeam
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgModel3DElement
        Return elementCounter
    End Function
    'ищет модель балки (tlc)
    Public Shared Function getModelBeamI(ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer, ByVal numberRow As Integer) As StructureElement
        Dim dataModelBeamI As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberProlet < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelBeam) = True Then
            Dim listModelBeam = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelBeam)
            If IsNothing(listModelBeam) = False Then
                If listModelBeam.Count > 0 Then
                    For k As Integer = 0 To listModelBeam.Count - 1
                        Dim tempData As StructureElement = listModelBeam.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userModelBeam As ModelBeam = tempData.getModelBeam()
                            If IsNothing(userModelBeam) = False Then
                                If numberProlet = userModelBeam.numberProlet And numberRow = userModelBeam.numberRow Then
                                    dataModelBeamI = tempData
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataModelBeamI
    End Function
    'строит модель балки
    Public Shared Function drawModelBeamI(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userBeamI As BeamI, ByVal lineShortBeam As DwgLine, ByVal dictSections As Dictionary(Of Integer, List(Of Vector2D)), ByVal idBridge As String, Optional styleModelBeam As ProjectCivilStructuresStyle = Nothing, Optional ByVal templateXML As String = "") As DwgModel3DElement
        Dim modelBeamI As DwgModel3DElement = Nothing
        If IsNothing(activProjectDocument) = True Then Return Nothing
        If IsNothing(userBeamI) = True Then Return Nothing
        If IsNothing(lineShortBeam) = True Then Return Nothing
        If lineShortBeam.Length <= 0 Then Return Nothing
        If IsNothing(dictSections) = True Then Return Nothing
        If dictSections.Count = 0 Then Return Nothing

        Dim shellStartSectionShort As Shell = Nothing
        Dim shellSection As Shell = Nothing
        Dim shellEndSectionShort As Shell = Nothing
        Dim startLenghtMonolith As Double = userBeamI.startLenghtMonolith
        Dim endLenghtMonolith As Double = userBeamI.endLenghtMonolith

        If startLenghtMonolith < 0 Then startLenghtMonolith = 0
        If endLenghtMonolith < 0 Then endLenghtMonolith = 0

        Dim lenSite As Double = userBeamI.lenght - startLenghtMonolith - endLenghtMonolith
        If lenSite < 0 Then lenSite = 0
        'стиль оформления
        If IsNothing(styleModelBeam) = False And File.Exists(templateXML) = True Then
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Балка (модель)")
        End If

        Dim dataModel As StructureElement = createModelBeamI(idBridge)
        'участок омоноличивания балок
        If dictSections.ContainsKey(1) = True AndAlso startLenghtMonolith > 0 Then
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(1)
            If IsNothing(listPoint1) = False AndAlso listPoint1.Count > 2 Then
                Dim listCoordinates(listPoint1.Count - 1) As Vector3D

                For i As Integer = 0 To listPoint1.Count - 1
                    Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                    listCoordinates(i) = New Vector3D(startLenghtMonolith, v1.X, v1.Y)
                Next

                If listCoordinates.Length > 2 Then
                    shellStartSectionShort = Tools.Polygon(listCoordinates)
                    shellStartSectionShort = Tools.Extrude(startLenghtMonolith, shellStartSectionShort)
                End If
            End If
        End If

        'основной участок
        If dictSections.ContainsKey(2) = True AndAlso lenSite > 0 Then
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(2)
            If IsNothing(listPoint1) = False AndAlso listPoint1.Count > 2 Then
                Dim listCoordinates(listPoint1.Count - 1) As Vector3D
                Dim sectionPos As Double = startLenghtMonolith + lenSite

                For i As Integer = 0 To listPoint1.Count - 1
                    Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                    listCoordinates(i) = New Vector3D(sectionPos, v1.X, v1.Y)
                Next

                If listCoordinates.Length > 2 Then
                    shellSection = Tools.Polygon(listCoordinates)
                    shellSection = Tools.Extrude(lenSite, shellSection)
                End If
            End If
        End If

        'участок омоноличивания в конце
        If dictSections.ContainsKey(1) = True AndAlso endLenghtMonolith > 0 Then
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(1)
            If IsNothing(listPoint1) = False AndAlso listPoint1.Count > 2 Then
                Dim listCoordinates(listPoint1.Count - 1) As Vector3D

                For i As Integer = 0 To listPoint1.Count - 1
                    Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                    listCoordinates(i) = New Vector3D(userBeamI.lenght, v1.X, v1.Y)
                Next

                If listCoordinates.Length > 2 Then
                    shellEndSectionShort = Tools.Polygon(listCoordinates)
                    shellEndSectionShort = Tools.Extrude(endLenghtMonolith, shellEndSectionShort)
                End If
            End If
        End If

        'делаем объединение
        Dim unionShell As Shell = Nothing

        If IsNothing(shellStartSectionShort) = False Then
            unionShell = shellStartSectionShort
        End If

        If IsNothing(shellSection) = False Then
            If IsNothing(unionShell) = True Then
                unionShell = shellSection
            Else
                unionShell = Tools.Union(unionShell, shellSection)
            End If
        End If

        If IsNothing(shellEndSectionShort) = False Then
            If IsNothing(unionShell) = True Then
                unionShell = shellEndSectionShort
            Else
                unionShell = Tools.Union(unionShell, shellEndSectionShort)
            End If
        End If

        shellSection = unionShell
        If IsNothing(shellSection) = True Then Return Nothing

        Dim lStartPoint As Vector3D = lineShortBeam.StartPoint
        Dim lEndPoint As Vector3D = lineShortBeam.EndPoint
        Dim boolExtend As Boolean = MathFunction.FuncExtendPos(lStartPoint, lEndPoint, userBeamI.a, userBeamI.b)

        Dim ox As Vector3D = lEndPoint - lStartPoint
        ox.Normalize()

        Dim baseOz As Vector3D = Vector3D.UnitZ
        Dim dotOzOx As Double = Math.Abs(baseOz.X * ox.X + baseOz.Y * ox.Y + baseOz.Z * ox.Z)
        If dotOzOx > 0.999 Then
            baseOz = Vector3D.UnitY
        End If

        Dim oy As Vector3D = Vector3D.Cross(baseOz, ox)
        oy.Normalize()

        Dim oz As Vector3D = Vector3D.Cross(ox, oy)
        oz.Normalize()

        Dim largeCoordinates As Boolean =
            Math.Abs(lStartPoint.X) >= 1000000.0 OrElse
            Math.Abs(lStartPoint.Y) >= 1000000.0 OrElse
            Math.Abs(lStartPoint.Z) >= 1000000.0 OrElse
            Math.Abs(lEndPoint.X) >= 1000000.0 OrElse
            Math.Abs(lEndPoint.Y) >= 1000000.0 OrElse
            Math.Abs(lEndPoint.Z) >= 1000000.0

        If largeCoordinates = True Then
            'Для больших координат BRep остается локальным: перенос хранится в DwgModel3DElement.Position.
            Tools.Multmatrix(Matrix.CreateUCS(Vector3D.Empty, ox, oy, oz), shellSection)
        Else
            'Сохраняем прежнее поведение функции для обычных координат.
            Tools.Multmatrix(Matrix.CreateUCS(lStartPoint, ox, oy, oz), shellSection)
        End If

        Dim elementBeamI = New StaticSolidElement("Балка  двутавровая (модель)", "SmdxElement", New ImProperties(), shellSection, New ImDocuments())
        modelBeamI = New DwgModel3DElement()
        If largeCoordinates = True Then
            elementBeamI.Origin = Vector3D.Empty
            modelBeamI.Position = lStartPoint
        End If
        modelBeamI.Element = elementBeamI
        activProjectDocument.ActiveSpace.Add(modelBeamI)
        If activProjectDocument.ActiveSpace.Entities.Contains(modelBeamI) = True Then
            Dim userModel As ModelBeam = New ModelBeam(userBeamI.numberProlet, userBeamI.numberRow)
            Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
            dataModel.KeyParameter = strGSON
            dataModel.DWGEntity = modelBeamI
            Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelBeamI, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
            styleModelBeam.setObjectStyle(modelBeamI)
        End If
        Return modelBeamI
    End Function
    'удалить все модели балок
    Public Shared Function deleteAllModelBeams(ByRef drawing As Drawing, ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As Integer
        Dim countModelBeams As Integer = 0
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.modelBeam) = True Then
            Dim listTopCountersBeam = dictionaryObjectsBridge.Item(StructureElement.typeObject.modelBeam)
            If IsNothing(listTopCountersBeam) = False Then
                If listTopCountersBeam.Count > 0 Then
                    For k As Integer = 0 To listTopCountersBeam.Count - 1
                        Dim tempData As StructureElement = listTopCountersBeam.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim modelEntity As DwgModel3DElement = tempData.DWGEntity
                            If IsNothing(modelEntity) = False Then
                                If drawing.ActiveSpace.Entities.Contains(modelEntity) = True Then
                                    drawing.ActiveSpace.Entities.Remove(modelEntity)
                                    countModelBeams += 1
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return countModelBeams
    End Function

End Class
