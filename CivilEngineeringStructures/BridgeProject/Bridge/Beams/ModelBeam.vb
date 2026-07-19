Imports System.ComponentModel
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
    'создать новую модель
    Public Shared Function createModelBeamI(ByVal idBridge As String) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassObject = StructureElement.classStructure.BeamI
        elementCounter.Name = StructureElement.typeObject.modelBeam
        elementCounter.Description = "Балка двутавровая> (модель)"
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
    Public Shared Function drawModelBeamI(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userBeamI As BeamI, ByVal lineShortBeam As DwgLine, ByVal dictSections As Dictionary(Of Integer, List(Of Vector2D)), ByVal idBridge As String, ByVal templateXML As String) As DwgModel3DElement
        Dim modelBeamI As DwgModel3DElement = Nothing
        If IsNothing(dictSections) = True Then Return Nothing
        If dictSections.Count = 0 Then Return Nothing
        Dim shellStartSectionShort As Shell = Nothing
        Dim shellSection As Shell = Nothing
        Dim shellEndSectionShort As Shell = Nothing
        Dim startPos As Double = 0
        Dim startLenghtMonolith As Double = userBeamI.startLenghtMonolith
        Dim endLenghtMonolith As Double = userBeamI.endLenghtMonolith
        Dim categoryTables As String = "Искусственные сооружения"
        Dim styleModel As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
        styleModel.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Модель, "Балка двутавровая (модель)")
        Dim dataModel As StructureElement = createModelBeamI(idBridge)
        'участок омоноличивания балок
        If dictSections.Count > 0 And startLenghtMonolith > 0 Then
            Dim listCoordinates As Vector3D() = {}
            Dim count As Integer = 0
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(1)
            For i As Integer = 0 To listPoint1.Count - 1
                Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                ReDim Preserve listCoordinates(count)
                listCoordinates(count) = New Vector3D(startLenghtMonolith, v1.X, v1.Y)
                count += 1
            Next
            If listCoordinates.Length > 2 Then
                shellStartSectionShort = Tools.Polygon(listCoordinates)
                shellStartSectionShort = Tools.Extrude(startLenghtMonolith, shellStartSectionShort)
                startPos += startLenghtMonolith
            End If
        End If
        'основной участок
        If dictSections.Count > 1 Then
            Dim listCoordinates As Vector3D() = {}
            Dim count As Integer = 0
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(2)
            Dim lenSite As Double = userBeamI.lenght - startLenghtMonolith - endLenghtMonolith
            For i As Integer = 0 To listPoint1.Count - 1
                Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                ReDim Preserve listCoordinates(count)
                listCoordinates(count) = New Vector3D(lenSite + startLenghtMonolith, v1.X, v1.Y)
                count += 1
            Next
            If listCoordinates.Length > 2 Then
                shellSection = Tools.Polygon(listCoordinates)
                shellSection = Tools.Extrude(lenSite, shellSection)
                startPos = startLenghtMonolith + lenSite

            End If
        End If
        'участок омоноличивания в конце
        If dictSections.Count > 0 And endLenghtMonolith > 0 Then
            Dim listCoordinates As Vector3D() = {}
            Dim count As Integer = 0
            Dim listPoint1 As List(Of Cad.Foundation.Vector2D) = dictSections.Item(1)
            For i As Integer = 0 To listPoint1.Count - 1
                Dim v1 As Cad.Foundation.Vector2D = listPoint1.Item(i)
                ReDim Preserve listCoordinates(count)
                listCoordinates(count) = New Vector3D(userBeamI.lenght, v1.X, v1.Y)
                count += 1
            Next
            If listCoordinates.Length > 2 Then
                shellEndSectionShort = Tools.Polygon(listCoordinates)
                shellEndSectionShort = Tools.Extrude(endLenghtMonolith, shellEndSectionShort)
            End If
        End If
        'делаем объединение
        If IsNothing(shellSection) = True And IsNothing(shellStartSectionShort) = False Then
            shellSection = shellStartSectionShort
        Else
            If IsNothing(shellSection) = False Then
                If IsNothing(shellStartSectionShort) = False Then
                    shellSection = Tools.Union(shellStartSectionShort, shellSection)
                End If
                If IsNothing(shellEndSectionShort) = False Then
                    shellSection = Tools.Union(shellEndSectionShort, shellSection)
                End If
            End If
        End If

        Dim lStartPoint As Vector3D = lineShortBeam.StartPoint
        Dim lEndPoint As Vector3D = lineShortBeam.EndPoint
        Dim boolExtend As Boolean = MathFunction.FuncExtendPos(lStartPoint, lEndPoint, userBeamI.a, userBeamI.b)
        Dim ox = lEndPoint - lStartPoint
        ox.Normalize()
        Dim oz = Vector3D.UnitZ
        Dim oy = Vector3D.Cross(oz, ox)
        oy.Normalize()
        oz = Vector3D.Cross(ox, oy)
        oz.Normalize()
        Tools.Multmatrix(Matrix.CreateUCS(lStartPoint, ox, oy, oz), shellSection)
        If IsNothing(lineShortBeam) = False Then
            If lineShortBeam.Length > 0 Then
                'удлинняем балку
                Dim elementBeamI = New StaticSolidElement("Балка  двутавровая (модель)", "SmdxElement", New ImProperties(), shellSection, New ImDocuments())
                modelBeamI = New DwgModel3DElement()
                modelBeamI.Element = elementBeamI
                activProjectDocument.ActiveSpace.Add(modelBeamI)
                If activProjectDocument.ActiveSpace.Entities.Contains(modelBeamI) = True Then
                    Dim userModel As ModelBeam = New ModelBeam(userBeamI.numberProlet, userBeamI.numberRow)
                    Dim strGSON As String = Newtonsoft.Json.JsonConvert.SerializeObject(userModel)
                    dataModel.KeyParameter = strGSON
                    dataModel.DWGEntity = modelBeamI
                    Dim boolRecData As Boolean = FuncXRecords.setXRecords(modelBeamI, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataModel)
                    styleModel.setObjectStyle(modelBeamI)
                End If
            End If
        End If
        Return modelBeamI
    End Function
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
