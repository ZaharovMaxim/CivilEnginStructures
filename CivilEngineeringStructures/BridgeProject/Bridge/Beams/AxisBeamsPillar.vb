'ось опирания балок
Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Visualization.Runtime
Imports Drawing = Topomatic.Dwg.Drawing
'ось опирания болок
Public Class AxisBeamsPillars
    'Inherits Bridge
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private _numberPillar As Integer 'номер опоры
    Private _numberSubPillar As Integer 'номер подопоры (при ее наличии, иначе 0)
    Private _numberProlet As Integer 'номер пролета
    Public _elementBridgePoint As PointsCollections
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberPillar = 0
        _numberSubPillar = 0
        _numberProlet = 0
        _elementBridgePoint = New PointsCollections
    End Sub

    <Browsable(True)>
    <Description("Номер опоры")>
    <Category("Свойства")>
    <DisplayName("Номер опоры")>
    <[ReadOnly](True)>
    Public Property numberPillar() As Integer
        Get
            Return _numberPillar
        End Get
        Set(value As Integer)
            _numberPillar = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Номер элемента")>
    <Category("Свойства")>
    <DisplayName("Номер элемента")>
    <[ReadOnly](True)>
    Public Property numberSubPillar() As Integer
        Get
            Return _numberSubPillar
        End Get
        Set(value As Integer)
            _numberSubPillar = value
        End Set
    End Property

    <Browsable(True)>
    <Description("Номер пролета")>
    <Category("Свойства")>
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
    'создать структуру Ось опирания балок
    Public Shared Function createAxis(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.BeamI
        elementAxis.Name = StructureElement.typeObject.axisPillarBeams
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.axisPillarBeams)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgLine
        Return elementAxis
    End Function
    'ищет ось опирания балок
    Public Shared Function getAxisBeamsPillar(ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberProlet As Integer, Optional removeDictionary As Boolean = False) As StructureElement
        Dim dataAxisBeamsPillar As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillarBeams) = True Then
            Dim listAxisBeamsPillar = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillarBeams)
            If IsNothing(listAxisBeamsPillar) = False Then
                If listAxisBeamsPillar.Count > 0 Then
                    For k As Integer = 0 To listAxisBeamsPillar.Count - 1
                        Dim tempData As StructureElement = listAxisBeamsPillar.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userAxisBeamsPillar As AxisBeamsPillars = tempData.getAxisBeamsPillar
                            If IsNothing(userAxisBeamsPillar) = False Then
                                If numberPillar = userAxisBeamsPillar.numberPillar Then
                                    If numberProlet = userAxisBeamsPillar.numberProlet Then
                                        dataAxisBeamsPillar = tempData
                                        If removeDictionary = True Then
                                            listAxisBeamsPillar.RemoveAt(k)
                                        End If
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataAxisBeamsPillar
    End Function
    'рисование оси опирания балок и присваивание семантики
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleAxisBeam As ProjectCivilStructuresStyle = Nothing, Optional ByVal templateXML As String = "") As StructureElement
        Dim axisLineBeamsPillar As DwgLine = Nothing
        Dim dataStructureBeamsPillar As StructureElement = Nothing
        'ищем существующую ось насадки
        Dim dataAxisBeamsPillar As StructureElement = getAxisBeamsPillar(dictionaryBridgeElements, numberPillar, numberProlet)
        If IsNothing(dataAxisBeamsPillar) = True Then
            dataAxisBeamsPillar = createAxis(idBridge)
        End If
        If IsNothing(dataAxisBeamsPillar) Then Return Nothing
        axisLineBeamsPillar = dataAxisBeamsPillar.DWGEntity
        If IsNothing(axisLineBeamsPillar) = True Then Return Nothing
        If IsNothing(styleAxisBeam) = True And File.Exists(templateXML) = True Then
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            styleAxisBeam = New ProjectCivilStructuresStyle(activProjectDocument)
            styleAxisBeam.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опирания балок")
            styleAxisBeam.setObjectStyle(axisLineBeamsPillar)
        End If
        'ось опирания балок
        axisLineBeamsPillar.StartPoint = _elementBridgePoint.StartAxisPoint
        axisLineBeamsPillar.EndPoint = _elementBridgePoint.EndAxisPoint
        Dim lenghtAxis As Double = (axisLineBeamsPillar.StartPoint.Pos - axisLineBeamsPillar.EndPoint.Pos).Length
        If lenghtAxis = 0 Then
            MsgBox("Ось опирания балок имеет нулевое значение.")
            Return dataAxisBeamsPillar
        End If
        If activProjectDocument.ActiveSpace.Entities.Contains(axisLineBeamsPillar) = False Then
            activProjectDocument.ActiveSpace.Entities.Add(axisLineBeamsPillar)
        End If
        Dim strJson As String = JsonConvert.SerializeObject(Me, Formatting.Indented)
        dataAxisBeamsPillar.KeyParameter = strJson
        dataAxisBeamsPillar.DWGEntity = axisLineBeamsPillar
        Dim boolRecData As Boolean = FuncXRecords.setXRecords(axisLineBeamsPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataAxisBeamsPillar)
        If IsNothing(styleAxisBeam) = False Then
            styleAxisBeam.setObjectStyle(axisLineBeamsPillar)
        End If
        Return dataAxisBeamsPillar
    End Function
    'функция проверяет и корректирует балку еcли она против направления пикетажа
    Public Shared Function correctionAxisDirectionBeam(ByRef axisBeam As DwgLine, ByRef align As Alignment) As Boolean
        correctionAxisDirectionBeam = False
        If IsNothing(axisBeam) = True Then
            Return False
        End If
        If IsNothing(align) = True Then
            Return False
        End If
        If axisBeam.Length = 0 Then
            Return False
        End If
        'находим пикеты начала и конца предудущей балки
        Try
            Dim pkStart As Double = 0
            Dim offStart As Double = 0
            Dim boolFindPk As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisBeam.StartPoint.Pos, pkStart, offStart)
            Dim pkEnd As Double = 0
            Dim offend As Double = 0
            Dim boolFindPk1 As Boolean = align.Plan.CompoundLine.PosToStaOffset(axisBeam.EndPoint.Pos, pkEnd, offend)
            If boolFindPk = True And boolFindPk1 = True Then
                If pkStart > pkEnd Then
                    Dim tempStartPoint As Vector3D = axisBeam.StartPoint
                    Dim tempEndPoint As Vector3D = axisBeam.EndPoint
                    axisBeam.StartPoint = tempEndPoint
                    axisBeam.EndPoint = tempStartPoint
                    Return True
                End If
            End If
        Catch ex As System.Exception
            Return False
        End Try
    End Function
    'функция по номеру пролета возврящает 2 оси опирания балок
    Public Shared Function getAxisPillarBeamsInProlet(ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberProlet As Integer) As List(Of StructureElement)
        Dim result As List(Of StructureElement) = New List(Of StructureElement)
        'первыя ось опирания балок
        Dim numberFirstPillar As Integer = numberProlet
        If IsNothing(dictionaryObjectsBridge) = True Then Return result
        If numberFirstPillar > 1 Then
            Dim dataFirstAxisBeamsPillar As StructureElement = getAxisBeamsPillar(dictionaryObjectsBridge, numberFirstPillar, numberProlet)
            If IsNothing(dataFirstAxisBeamsPillar) = False Then
                result.Add(dataFirstAxisBeamsPillar)
            End If
        ElseIf numberFirstPillar = 1 Then 'это первая опора, забираем ось опоры
            Dim dataFirstAxisBeamsPillar As StructureElement = Pillar.getAxisPillar(dictionaryObjectsBridge, numberFirstPillar)
            If IsNothing(dataFirstAxisBeamsPillar) = False Then
                result.Add(dataFirstAxisBeamsPillar)
            End If
        End If
        Dim numberSecondPillar As Integer = numberProlet + 1
        If numberSecondPillar > 1 Then
            Dim dataSecondAxisBeamsPillar As StructureElement = getAxisBeamsPillar(dictionaryObjectsBridge, numberSecondPillar, numberProlet)
            If IsNothing(dataSecondAxisBeamsPillar) = False Then
                result.Add(dataSecondAxisBeamsPillar)
            Else
                dataSecondAxisBeamsPillar = Pillar.getAxisPillar(dictionaryObjectsBridge, numberSecondPillar)
                If IsNothing(dataSecondAxisBeamsPillar) = False Then
                    result.Add(dataSecondAxisBeamsPillar)
                End If
            End If
        End If
        Return result
    End Function
    'функция рисует оси опирания балок
    Public Shared Function drawAxisBeamsPillar(ByRef drawingDocument As Drawing, ByRef axisPillar As Dictionary(Of Integer, List(Of StructureElement)), ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleAxisBeamsPillar As ProjectCivilStructuresStyle = Nothing) As Boolean
        If IsNothing(drawingDocument) = False Then
            If IsNothing(axisPillar) = False Then
                If axisPillar.Count > 0 Then
                    For i As Integer = 0 To axisPillar.Count - 1
                        Try
                            Dim listAxisPillar As List(Of StructureElement) = axisPillar.ElementAt(i).Value
                            Dim dataPrevAxisBeamsPillar As StructureElement = listAxisPillar.Item(0)
                            Dim dataAxisPillar As StructureElement = listAxisPillar.Item(1)
                            Dim dataAxisBeamsPillar As StructureElement = listAxisPillar.Item(2)
                            If IsNothing(dataPrevAxisBeamsPillar) = False Then
                                Dim userAxisBeamsPillar As AxisBeamsPillars = dataAxisBeamsPillar.getAxisBeamsPillar()
                                Dim acLineAxisBeamsPillar As DwgLine = Nothing
                                Dim oldDataPrevAxisBeamsPillar As StructureElement = AxisBeamsPillars.getAxisBeamsPillar(dictionaryObjectsBridge, userAxisBeamsPillar.numberPillar, userAxisBeamsPillar.numberProlet, True)
                                If IsNothing(oldDataPrevAxisBeamsPillar) = False Then
                                    acLineAxisBeamsPillar = oldDataPrevAxisBeamsPillar.DWGEntity
                                    acLineAxisBeamsPillar.StartPoint = userAxisBeamsPillar._elementBridgePoint.StartAxisPoint
                                    acLineAxisBeamsPillar.EndPoint = userAxisBeamsPillar._elementBridgePoint.EndAxisPoint
                                Else
                                    acLineAxisBeamsPillar = dataPrevAxisBeamsPillar.DWGEntity
                                End If
                                If IsNothing(acLineAxisBeamsPillar) = True Then
                                    acLineAxisBeamsPillar = New DwgLine
                                    dataPrevAxisBeamsPillar.DWGEntity = acLineAxisBeamsPillar
                                End If
                                If acLineAxisBeamsPillar.Length = 0 Then
                                    acLineAxisBeamsPillar.StartPoint = userAxisBeamsPillar._elementBridgePoint.StartAxisPoint
                                    acLineAxisBeamsPillar.EndPoint = userAxisBeamsPillar._elementBridgePoint.EndAxisPoint
                                End If
                                If drawingDocument.ActiveSpace.Entities.Contains(acLineAxisBeamsPillar) = False Then
                                    drawingDocument.ActiveSpace.Entities.Add(acLineAxisBeamsPillar)
                                    Dim boolSetStyleBeam As Boolean = styleAxisBeamsPillar.setObjectStyle(acLineAxisBeamsPillar)
                                End If
                                Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisBeamsPillar)
                                dataAxisBeamsPillar.KeyParameter = keyParam
                                Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(acLineAxisBeamsPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataAxisBeamsPillar)
                            End If
                            If IsNothing(dataAxisBeamsPillar) = False Then
                                Dim userAxisBeamsPillar As AxisBeamsPillars = dataAxisBeamsPillar.getAxisBeamsPillar()
                                Dim acLineAxisBeamsPillar As DwgLine = Nothing
                                Dim oldDataAxisBeamsPillar As StructureElement = AxisBeamsPillars.getAxisBeamsPillar(dictionaryObjectsBridge, userAxisBeamsPillar.numberPillar, userAxisBeamsPillar.numberProlet, True)
                                If IsNothing(oldDataAxisBeamsPillar) = False Then
                                    acLineAxisBeamsPillar = oldDataAxisBeamsPillar.DWGEntity
                                    acLineAxisBeamsPillar.StartPoint = userAxisBeamsPillar._elementBridgePoint.StartAxisPoint
                                    acLineAxisBeamsPillar.EndPoint = userAxisBeamsPillar._elementBridgePoint.EndAxisPoint
                                Else
                                    acLineAxisBeamsPillar = dataAxisBeamsPillar.DWGEntity
                                End If
                                If IsNothing(acLineAxisBeamsPillar) = True Then
                                    acLineAxisBeamsPillar = New DwgLine
                                    dataAxisBeamsPillar.DWGEntity = acLineAxisBeamsPillar
                                End If
                                If acLineAxisBeamsPillar.Length = 0 Then
                                    acLineAxisBeamsPillar.StartPoint = userAxisBeamsPillar._elementBridgePoint.StartAxisPoint
                                    acLineAxisBeamsPillar.EndPoint = userAxisBeamsPillar._elementBridgePoint.EndAxisPoint
                                End If
                                If drawingDocument.ActiveSpace.Entities.Contains(acLineAxisBeamsPillar) = False Then
                                    drawingDocument.ActiveSpace.Entities.Add(acLineAxisBeamsPillar)
                                    Dim boolSetStyleBeam As Boolean = styleAxisBeamsPillar.setObjectStyle(acLineAxisBeamsPillar)
                                End If
                                Dim keyParam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userAxisBeamsPillar)
                                dataAxisBeamsPillar.KeyParameter = keyParam
                                Dim boolRecDatabeam As Boolean = FuncXRecords.setXRecords(acLineAxisBeamsPillar, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataAxisBeamsPillar)
                            End If
                        Catch ex As system.Exception
                        End Try
                    Next i
                End If
            End If
        End If
        'удаляем лишние элементы
        If dictionaryObjectsBridge.ContainsKey(StructureElement.typeObject.axisPillarBeams) = True Then
            Dim listAxisBeams As List(Of StructureElement) = dictionaryObjectsBridge.Item(StructureElement.typeObject.axisPillarBeams)
            If IsNothing(listAxisBeams) = False Then
                For Each dataBeam As StructureElement In listAxisBeams
                    If IsNothing(dataBeam) = False Then
                        Dim axisLine As DwgEntity = dataBeam.DWGEntity
                        If IsNothing(axisLine) = False Then
                            If drawingDocument.ActiveSpace.Entities.Contains(axisLine) = True Then
                                drawingDocument.ActiveSpace.Entities.Remove(axisLine)
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Return True
    End Function
End Class
