Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Media.Media3D
Imports Microsoft.Office.Interop.Excel
Imports Topomatic
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.Visualization.Runtime
Imports Vector3D = Topomatic.Cad.Foundation.Vector3D

Public Class CounterBeam
    Public Enum typeCounterBeam
        <Description("Верх плиты балки")> TopPlateBeam = 0
        <Description("Основание балки")> DownBeam = 1
        <Description("Не определено")> Notdefined = 2
    End Enum
    Private _numberProlet As Integer 'номер пролета
    Private _numberRow As Integer 'номер ряда
    Private _type As typeCounterBeam 'тип (низ балки, верх плиты балки)
    'Private _entityAxisBeams As DwgLine
    Public Sub New()
        _numberProlet = 0
        _numberRow = 0
        _type = typeCounterBeam.Notdefined
    End Sub
    ' Номер пролета
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
    ' Номер пролета
    <Browsable(True)>
    <Description("Номер ряда балок")>
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
    ' Номер пролета
    <Browsable(True)>
    <Description("Тип контура")>
    <Category("Свойства сооружения")>
    <DisplayName("Тип контура")>
    <[ReadOnly](True)>
    Public Property TypeCounter() As typeCounterBeam
        Get
            Return _type
        End Get
        Set(value As typeCounterBeam)
            _type = value
        End Set
    End Property
    'создать новую структуру данных для контура балки
    Public Shared Function createCounter(ByVal idBridge As String, ByVal typeCounter As StructureElement.typeObject) As StructureElement
        Dim elementCounter As StructureElement = New StructureElement()
        elementCounter.Label = "Мосты и путепроводы"
        elementCounter.ClassBridgeObject = StructureElement.classBridge.SpanStructures 'пролетные строение
        elementCounter.ClassObject = StructureElement.classStructure.BeamI 'балка
        elementCounter.Name = typeCounter
        Dim deskBeam As String = StructureElement.GetDescription(typeCounter)
        elementCounter.Description = deskBeam
        elementCounter.KeyParameter = ""
        elementCounter.IdElement = Guid.NewGuid.ToString
        elementCounter.IdStructure = idBridge
        elementCounter.Note = ""
        elementCounter.DWGEntity = New DwgPolyline3D
        Return elementCounter
    End Function
    'функция вычисляет крайние точки контура балки
    Public Function setCounterPoint(ByVal polyCounter As DwgPolyline3D, ByVal heightPlate As Double) As Dictionary(Of Integer, PointStructure)
        Dim result As Dictionary(Of Integer, PointStructure) = New Dictionary(Of Integer, PointStructure)
        If polyCounter.Count = 4 Then
            Dim pointStartLeft As Vector3D = polyCounter.Item(0)
            Dim pointEndLeft As Vector3D = polyCounter.Item(1)
            Dim pointDownStartLeft As Vector3D = BridgeGeometry.calculatePointP(pointStartLeft, pointEndLeft, 0, -1 * heightPlate)
            Dim pointDownEndLeft As Vector3D = BridgeGeometry.calculatePointP(pointEndLeft, pointStartLeft, 0, -1 * heightPlate)

            Dim pointStartRight As Vector3D = polyCounter.Item(2)
            Dim pointEndRight As Vector3D = polyCounter.Item(3)
            Dim pointDownStartRight As Vector3D = BridgeGeometry.calculatePointP(pointStartRight, pointEndRight, 0, -1 * heightPlate)
            Dim pointDownEndRight As Vector3D = BridgeGeometry.calculatePointP(pointEndRight, pointStartRight, 0, -1 * heightPlate)

            Dim numberPoint As Integer = 1
            Dim x As Double = Math.Round(pointStartLeft.X, 3)
            Dim y As Double = Math.Round(pointStartLeft.Y, 3)
            Dim z As Double = Math.Round(pointStartLeft.Z, 3)
            Dim h1 As Double = Math.Round(heightPlate, 3)
            Dim dx As Double = Math.Round(pointStartLeft.X - pointDownStartLeft.X, 3)
            Dim dy As Double = Math.Round(pointStartLeft.Y - pointDownStartLeft.Y, 3)
            Dim code As String = "leftPt1"
            result.Add(numberPoint, New PointStructure(x, y, z, dx, dy, -1 * h1, 0, code))
            numberPoint += 1

            x = Math.Round(pointEndLeft.X, 3)
            y = Math.Round(pointEndLeft.Y, 3)
            z = Math.Round(pointEndLeft.Z, 3)
            h1 = Math.Round(heightPlate, 3)
            dx = Math.Round(pointEndLeft.X - pointDownEndLeft.X, 3)
            dy = Math.Round(pointEndLeft.Y - pointDownEndLeft.Y, 3)
            code = "leftPt2"
            result.Add(numberPoint, New PointStructure(x, y, z, dx, dy, -1 * h1, 0, code))
            numberPoint += 1

            x = Math.Round(pointStartRight.X, 3)
            y = Math.Round(pointStartRight.Y, 3)
            z = Math.Round(pointStartRight.Z, 3)
            h1 = Math.Round(heightPlate, 3)
            dx = Math.Round(pointStartRight.X - pointDownStartRight.X, 3)
            dy = Math.Round(pointStartRight.Y - pointDownStartRight.Y, 3)
            code = "rightPt2"
            result.Add(numberPoint, New PointStructure(x, y, z, dx, dy, -1 * h1, 0, code))
            numberPoint += 1

            x = Math.Round(pointEndRight.X, 3)
            y = Math.Round(pointEndRight.Y, 3)
            z = Math.Round(pointEndRight.Z, 3)
            h1 = Math.Round(heightPlate, 3)
            dx = Math.Round(pointEndRight.X - pointDownEndRight.X, 3)
            dy = Math.Round(pointEndRight.Y - pointDownEndRight.Y, 3)
            code = "rightPt1"
            result.Add(numberPoint, New PointStructure(x, y, z, dx, dy, -1 * h1, 0, code))
        End If
        Return result
    End Function
    Public Function setCounterPoint(ByVal polyCounter As DwgPolyline3D) As Dictionary(Of Integer, PointStructure)
        Dim result As New Dictionary(Of Integer, PointStructure)
        If polyCounter.Count = 4 Then
            Dim pointStartLeft As Vector3D = polyCounter.Item(0)
            Dim pointEndLeft As Vector3D = polyCounter.Item(1)

            Dim pointStartRight As Vector3D = polyCounter.Item(2)
            Dim pointEndRight As Vector3D = polyCounter.Item(3)

            result.Add(1, New PointStructure(pointStartLeft.X, pointStartLeft.Y, pointStartLeft.Z, 0, 0, 0, 0, "leftPt1"))
            result.Add(2, New PointStructure(pointEndLeft.X, pointEndLeft.Y, pointEndLeft.Z, 0, 0, 0, 0, "leftPt2"))

            result.Add(3, New PointStructure(pointStartRight.X, pointStartRight.Y, pointStartRight.Z, 0, 0, 0, 0, "rightPt2"))
            result.Add(4, New PointStructure(pointEndRight.X, pointEndRight.Y, pointEndRight.Z, 0, 0, 0, 0, "rightPt1"))
        End If
        Return result
    End Function
    'ищет контур балки
    Public Shared Function getContour(ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal type As StructureElement.typeObject, ByVal numberProlet As Integer, ByVal numberRow As Integer) As StructureElement
        Dim dataCounterBeam As StructureElement = Nothing
        If IsNothing(dictionaryObjectsBridge) = True Then Return Nothing
        If numberProlet < 1 Then Return Nothing
        If dictionaryObjectsBridge.ContainsKey(type) = True Then
            Dim listTopCountersBeam = dictionaryObjectsBridge.Item(type)
            If IsNothing(listTopCountersBeam) = False Then
                If listTopCountersBeam.Count > 0 Then
                    For k As Integer = 0 To listTopCountersBeam.Count - 1
                        Dim tempData As StructureElement = listTopCountersBeam.Item(k)
                        If IsNothing(tempData) = False Then
                            Dim userCounterBeam As CounterBeam = tempData.getCounterBeam()
                            If IsNothing(userCounterBeam) = False Then
                                If numberProlet = userCounterBeam.numberProlet And numberRow = userCounterBeam.numberRow Then
                                    dataCounterBeam = tempData
                                    Exit For
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Return dataCounterBeam
    End Function
    'функция рисует контур по верху и низу балки и записывает семантику
    Public Shared Function drawContour(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal userBeam As BeamI, ByVal idBridge As String, ByRef dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), Optional styleCounterTopBeam As ProjectCivilStructuresStyle = Nothing, Optional styleCounterBottomBeam As ProjectCivilStructuresStyle = Nothing, Optional templateXML As String = "") As Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        drawContour = New Dictionary(Of StructureElement.typeObject, DwgPolyline3D)
        If IsNothing(userBeam) Then Return drawContour
        If IsNothing(userBeam._elementBridgePoint.StartAxisPoint) = True Then Return drawContour
        If IsNothing(userBeam._elementBridgePoint.EndAxisPoint) = True Then Return drawContour
        Dim acLineShortBeam As DwgLine = New DwgLine()
        acLineShortBeam.StartPoint = userBeam._elementBridgePoint.StartAxisPoint
        acLineShortBeam.EndPoint = userBeam._elementBridgePoint.EndAxisPoint
        'вспомогательные построения
        Dim layerBeam As DwgLayer = activProjectDocument.ActiveLayer
        Dim colorBeam As CadColor = New CadColor(7)
        Dim nameTypeLineBeam As DwgLinetype = activProjectDocument.ActiveLinetype
        Dim ScaleTypeLineBeam As Integer = 1
        Dim widthTypeLineBeam As Integer = 20
        'стиль
        Dim categoryTables As String = "Искусственные сооружения"
        If IsNothing(styleCounterTopBeam) = True And File.Exists(templateXML) = False Then
            styleCounterTopBeam = New ProjectCivilStructuresStyle(activProjectDocument)
            styleCounterTopBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Верх ребра плиты балки")
        End If
        Dim startPointElements As List(Of Vector3D) = New List(Of Vector3D)
        Dim endPointElements As List(Of Vector3D) = New List(Of Vector3D)
        Dim boolRestoreElements As Boolean = CalculationBeams.restoreElementsBeam(acLineShortBeam, userBeam, startPointElements, endPointElements, CalculationBeams.rectoreBeam.siteMonolit)
        If startPointElements.Count = 4 And endPointElements.Count = 4 Then
            Dim toplineBeam As DwgPolyline3D = New DwgPolyline3D
            'ищем в массиве уже существующий элемент
            Dim dataCounterBeam As StructureElement = getContour(dictionaryObjectsBridge, StructureElement.typeObject.counterTopBeam, userBeam.numberProlet, userBeam.numberRow)
            Dim userCounterTopBeam As CounterBeam = New CounterBeam()
            If IsNothing(dataCounterBeam) = False Then
                userCounterTopBeam = dataCounterBeam.getCounterBeam
                If IsNothing(dataCounterBeam.DWGEntity) = False Then
                    toplineBeam = dataCounterBeam.DWGEntity
                End If
            Else
                dataCounterBeam = createCounter(idBridge, StructureElement.typeObject.counterTopBeam)
            End If
            If toplineBeam.Count = 4 Then
                toplineBeam.Item(0) = startPointElements.Item(0)
                toplineBeam.Item(1) = endPointElements.Item(0)
                toplineBeam.Item(2) = endPointElements.Item(1)
                toplineBeam.Item(3) = startPointElements.Item(1)
            Else
                toplineBeam.Clear()
                toplineBeam.Add(startPointElements.Item(0))
                toplineBeam.Add(endPointElements.Item(0))
                toplineBeam.Add(endPointElements.Item(1))
                toplineBeam.Add(startPointElements.Item(1))
            End If

            userCounterTopBeam.numberProlet = userBeam.numberProlet
            userCounterTopBeam.numberRow = userBeam.numberRow
            userCounterTopBeam.TypeCounter = CounterBeam.typeCounterBeam.TopPlateBeam
            Dim strGSONTopBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounterTopBeam)
            dataCounterBeam.KeyParameter = strGSONTopBeam
            dataCounterBeam.DWGEntity = toplineBeam
            Dim styleBeam As Boolean = styleCounterTopBeam.setObjectStyle(toplineBeam)
            Dim boolRecDataPillar = FuncXRecords.setXRecords(toplineBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCounterBeam)
            If activProjectDocument.ActiveSpace.Entities.Contains(toplineBeam) = False Then
                activProjectDocument.ActiveSpace.Add(toplineBeam)
                toplineBeam.Closed = True
            End If
            drawContour.Add(StructureElement.typeObject.counterTopBeam, toplineBeam)
            '=================================================================================================================================================
            'низ контура
            Dim bottomLineBeam As DwgPolyline3D = New DwgPolyline3D
            If IsNothing(styleCounterBottomBeam) = True And File.Exists(templateXML) = False Then
                styleCounterBottomBeam = New ProjectCivilStructuresStyle(activProjectDocument)
                styleCounterBottomBeam.setObjectStyle(templateXML, categoryTables, "Балки мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Низ ребра балки")
            End If
            'ищем в массиве уже существующий элемент
            dataCounterBeam = getContour(dictionaryObjectsBridge, StructureElement.typeObject.counterBottomBeam, userBeam.numberProlet, userBeam.numberRow)
            Dim userCounterBottomBeam As CounterBeam = New CounterBeam()
            If IsNothing(dataCounterBeam) = False Then
                userCounterBottomBeam = dataCounterBeam.getCounterBeam
                If IsNothing(dataCounterBeam.DWGEntity) = False Then
                    bottomLineBeam = dataCounterBeam.DWGEntity
                End If
            Else
                dataCounterBeam = createCounter(idBridge, StructureElement.typeObject.counterBottomBeam)
            End If
            If bottomLineBeam.Count = 4 Then
                bottomLineBeam.Item(0) = startPointElements.Item(2)
                bottomLineBeam.Item(1) = endPointElements.Item(2)
                bottomLineBeam.Item(2) = endPointElements.Item(3)
                bottomLineBeam.Item(3) = startPointElements.Item(3)
            Else
                bottomLineBeam.Clear()
                bottomLineBeam.Add(startPointElements.Item(2))
                bottomLineBeam.Add(endPointElements.Item(2))
                bottomLineBeam.Add(endPointElements.Item(3))
                bottomLineBeam.Add(startPointElements.Item(3))
            End If
            userCounterBottomBeam.numberProlet = userBeam.numberProlet
            userCounterBottomBeam.numberRow = userBeam.numberRow
            userCounterBottomBeam.TypeCounter = CounterBeam.typeCounterBeam.DownBeam
            Dim strGSONBottomBeam As String = Newtonsoft.Json.JsonConvert.SerializeObject(userCounterBottomBeam)
            dataCounterBeam.KeyParameter = strGSONBottomBeam
            dataCounterBeam.DWGEntity = bottomLineBeam
            styleBeam = styleCounterBottomBeam.setObjectStyle(bottomLineBeam)
            boolRecDataPillar = FuncXRecords.setXRecords(bottomLineBeam, StructureElement.tableXRecords.PROJECT_STRUCTURES, dataCounterBeam)
            If activProjectDocument.ActiveSpace.Entities.Contains(bottomLineBeam) = False Then
                activProjectDocument.ActiveSpace.Add(bottomLineBeam)
                bottomLineBeam.Closed = True
            End If
            drawContour.Add(StructureElement.typeObject.counterBottomBeam, bottomLineBeam)
        End If
        Return drawContour
    End Function
End Class
