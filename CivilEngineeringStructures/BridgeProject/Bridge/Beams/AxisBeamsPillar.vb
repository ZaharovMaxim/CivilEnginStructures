'ось опирания балок
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Topomatic.Alg
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Lisp.LMath
Imports Topomatic.Visualization.Runtime
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
    'создать новую ось
    Public Shared Function createAxis(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
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
    'ищет ось раскладки балок
    Public Shared Function getAxisBeamsPillar(ByVal dictionaryObjectsBridge As Dictionary(Of StructureElement.typeObject, List(Of StructureElement)), ByVal numberPillar As Integer, ByVal numberProlet As Integer) As StructureElement
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
    'рисование оси 
    Public Function drawAxis(ByRef activProjectDocument As Topomatic.Dwg.Drawing, ByVal idBridge As String, ByVal templateXML As String, ByVal dictionaryBridgeElements As Dictionary(Of StructureElement.typeObject, List(Of StructureElement))) As StructureElement
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
        If axisLineBeamsPillar.Length = 0 Then
            Dim layerObject As DwgLayer = activProjectDocument.ActiveLayer
            Dim colorObject As CadColor = New CadColor(7)
            Dim nameTypeLineObject As DwgLinetype = activProjectDocument.ActiveLinetype
            Dim ScaleTypeLineObject As Integer = 1
            Dim widthTypeLineObject As Integer = 20
            'стиль
            Dim categoryTables As String = "Искусственные сооружения"
            Dim styleObject As ProjectCivilStructuresStyle = New ProjectCivilStructuresStyle(activProjectDocument)
            styleObject.setObjectStyle(templateXML, categoryTables, "Опоры мостовых сооружений", ProjectCivilStructuresStyle.typeEntity.Линия, "Ось опирания балок")
            styleObject.setObjectStyle(axisLineBeamsPillar)
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



End Class
