Imports System.IO
Imports System.Windows
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities
Public Class ProjectCivilStructuresStyle
    Public Enum typeEntity
        Линия
        Полилиния
        Модель
    End Enum
    Private _indexColor As CadColor
    Private _layer As DwgLayer
    Private _lineType As DwgLinetype
    Private _lineTypeScale As Double
    Private _lineWeight As Double
    Private _lineGWidth As Double
    Private _drawing As Drawing
    Public Sub New(drawing As Drawing)
        If IsNothing(drawing) = False Then
            _indexColor = New CadColor(7)
            _layer = drawing.ActiveLayer
            _lineType = drawing.ActiveLinetype
            _lineTypeScale = 1
            _lineWeight = 0.2
            _lineGWidth = 0
            _drawing = drawing
        End If
    End Sub
    Public Sub New(drawing As Drawing, IndexColor As CadColor, Layer As DwgLayer, LineType As DwgLinetype, LineTypeScale As Double, LineWeight As Double, LineGWidth As Double)
        If IsNothing(drawing) = False Then
            _indexColor = IndexColor
            _layer = Layer
            _lineType = LineType
            _lineTypeScale = LineTypeScale
            _lineWeight = LineWeight
            _lineGWidth = LineGWidth
            _drawing = drawing
        End If
    End Sub

    Public Property IndexColor As CadColor
        Get
            Return _indexColor
        End Get
        Set(value As CadColor)
            _indexColor = value
        End Set
    End Property

    'Получить или установить слой чертежа</summary>
    Public Property Layer As DwgLayer
        Get
            Return _layer
        End Get
        Set(value As DwgLayer)
            _layer = value
        End Set
    End Property

    'Получить или установить тип линии</summary>
    Public Property LineType As DwgLinetype
        Get
            Return _lineType
        End Get
        Set(value As DwgLinetype)
            _lineType = value
        End Set
    End Property

    'Получить или установить масштаб типа линии</summary>
    Public Property LineTypeScale As Double
        Get
            Return _lineTypeScale
        End Get
        Set(value As Double)
            _lineTypeScale = value
        End Set
    End Property

    'Получить или установить вес линии</summary>
    Public Property LineWeight As Double
        Get
            Return _lineWeight
        End Get
        Set(value As Double)
            _lineWeight = value
        End Set
    End Property

    'Получить или установить глобальную ширину линии</summary>
    Public Property LineGWidth As Double
        Get
            Return _lineGWidth
        End Get
        Set(value As Double)
            _lineGWidth = value
        End Set
    End Property

    'Получить или установить чертеж</summary>
    Public Property ActiveProjectDrawing As Drawing
        Get
            Return _drawing
        End Get
        Set(value As Drawing)
            _drawing = value
        End Set
    End Property

    ' Создает копию текущего стиля
    Public Function Clone() As ProjectCivilStructuresStyle
        Return New ProjectCivilStructuresStyle(_drawing, _indexColor, _layer, _lineType,
                                   _lineTypeScale, _lineWeight, _lineGWidth)
    End Function
    ' Сбрасывает стиль к значениям по умолчанию
    Public Sub ResetToDefault()
        If _drawing IsNot Nothing Then
            _indexColor = New CadColor(7)
            _layer = _drawing.ActiveLayer
            _lineType = _drawing.ActiveLinetype
            _lineTypeScale = 1
            _lineWeight = 0.2
            _lineGWidth = 0
        End If
    End Sub
    ' Проверяет валидность стиля
    Public Function IsValid() As Boolean
        Return _drawing IsNot Nothing
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'получаем шаблон оформления в формате xml
    Public Shared Function getTemplateXml() As Dictionary(Of String, String)
        Dim dictPutchXml As Dictionary(Of String, String) = New Dictionary(Of String, String)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.readDirectoriesSupport(arrayDirSupport)
        Dim parentFolder As String = ""
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("InfrastradaToolsUtility")
                If pos > -1 Then
                    parentFolder = Mid(nameFolder, 1, pos)
                    parentFolder = parentFolder & "InfrastradaToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim directorySupport As String = parentFolder & "\FileResources\Sample\"
        If Directory.Exists(directorySupport) = True Then
            Dim xmlFiles As String() = Directory.GetFiles(directorySupport, "*.xml")
            If IsArray(xmlFiles) = True Then
                For i As Integer = 0 To xmlFiles.Length - 1
                    Dim putchXml As String = xmlFiles(i)
                    Dim nameTemplate As String = IO.Path.GetFileNameWithoutExtension(putchXml)
                    If dictPutchXml.ContainsKey(nameTemplate) = False Then
                        dictPutchXml.Add(nameTemplate, putchXml)
                    End If
                Next
            End If
        End If
        Return dictPutchXml
    End Function
    'устанавливает стиль
    Public Sub setObjectStyle(ByVal putchTemplateXML As String, ByVal categoryTables As String, ByVal nameTable As String, ByVal typeEntity As typeEntity, ByVal nameObject As String)
        Dim arrayBeamProperties As String(,) = Nothing
        Dim boolFindParcelsStyle As Boolean = FuncXML.FuncReadDataObjectToXMLbyDesk(putchTemplateXML, categoryTables, nameTable, typeEntity.ToString, nameObject, arrayBeamProperties)
        If IsArray(arrayBeamProperties) = True Then
            Dim boolCreateStyle As Boolean = createStyleLinearObject(_drawing, putchTemplateXML, arrayBeamProperties, _layer, _lineType, _indexColor, _lineTypeScale, _lineWeight,,, _lineGWidth)
        End If
    End Sub
    Public Function setObjectStyle(ByRef entity As DwgEntity) As Boolean
        If IsNothing(entity) = False Then
            entity.Color = IndexColor
            entity.Layer = Layer
            entity.Linetype = LineType
            entity.LinetypeScale = LineTypeScale
            entity.Lineweight = LineWeight
            If TypeOf entity Is DwgPolyline And LineGWidth > 0 Then
                Dim pPoly As DwgPolyline = entity
                pPoly.Width = LineGWidth
            End If
        End If
        Return True
    End Function
    '===========================================================================================================================
    'фунция создает все необходимые стили для линейных объектов
    Public Function createStyleLinearObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "", Optional ByRef lineGWidth As Double = 0) As Boolean
        createStyleLinearObject = False
        '===================================================================================================================
        'оформляем откосы
        If IsArray(ArrayRez) = True Then
            Dim nameObject As String = ""
            Dim typeUserObject As String = ""
            Dim deskObject As String = ""
            Dim nameTypeLine As String = ""
            Dim lineTypeCode1 As String = ""
            Dim arrayMaskLine As String() = Nothing
            For i As Integer = 0 To ArrayRez.GetUpperBound(1)
                '=========================================================================================================
                'создаем слой
                If IsNothing(ArrayRez(0, i)) = True Then Continue For
                If ArrayRez(0, i).Trim Like "Layer" Then
                    Dim NameLayer As String = ArrayRez(1, i)
                    If IsNothing(NameLayer) = False Then
                        NameLayer = NameLayer.Trim
                        'ищем слой
                        Dim tempObjectLayer As DwgLayer = getLayerDwg(NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If tempObjectLayer.Name Like "0" Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = ""
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                            lineTypeCode = FuncXML.FuncFindLineTypeCodeByNameTypeLineToXML(namePutchXML, nameLayerTypeLine, strDeskTypeLine)
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    'штриовой тип линии
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                        'Имя объекта стиля
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = True Then
                        nameObject = "Continuous"
                    Else
                        If nameObject Like "ПоСлою" Then
                            nameObject = "ByLayer"
                        ElseIf nameObject Like "ПоБлоку" Then
                            nameObject = "ByBlock"
                        Else
                            nameObject = ArrayRez(1, i).Trim
                            lineTypeCode1 = FuncXML.FuncFindLineTypeCodeByNameTypeLineToXML(namePutchXML, nameObject, deskObject)
                        End If
                    End If

                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim strColor As String = ArrayRez(1, i).Trim
                        If strColor Like "ПоСлою" Then
                            colorObj = New CadColor(256)
                        ElseIf strColor Like "ПоБлоку" Then
                            colorObj = New CadColor(0)
                        Else
                            Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                            If ColorLine < 0 Then ColorLine = 0
                            If ColorLine > 255 Then ColorLine = 255
                            colorObj = New CadColor(ColorLine)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ByLineWeightDefault" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineGWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        lineGWidth = Val(ArrayRez(1, i))
                        lineGWidth = Math.Abs(lineGWidth)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Note" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Next i
            '===========================================================================================================
            arrayMaskLine = lineTypeCode1.Split(",")
            If IsArray(arrayMaskLine) = True Then
                If arrayMaskLine.Length > 2 Then
                    nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameObject, deskObject, arrayMaskLine, namePutchXML)
                End If
            End If
            Return True
        End If
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для штриховок
    Public Shared Function FuncCreateStyleHatchObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef namePattern As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean
        FuncCreateStyleHatchObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = "Непрерывный"
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            If ArrayRez(0, i).Trim Like "Layer" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    Dim NameLayer As String = ArrayRez(1, i).Trim
                    Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                    'слой не найден, ищем его в шаблоне и создаем его
                    If IsNothing(tempObjectLayer) = True Then
                        Dim ArrayRezLayer As String(,) = Nothing
                        If IO.File.Exists(namePutchXML) = True Then
                            Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                        End If
                        Dim indexColor As Integer = 7
                        Dim nameLayerTypeLine As String = "Continuous"
                        Dim strDeskTypeLine As String = "Непрерывный"
                        Dim lineTypeCode As String = ""
                        Dim widthLineLayer As Integer = 20
                        Dim TypeLineLayer As DwgLinetype = Nothing
                        If IsArray(ArrayRezLayer) = True Then
                            If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                    If ArrayRezLayer(0, k).Trim Like "Color" Then
                                        indexColor = Val(ArrayRezLayer(1, k).Trim)
                                        If indexColor <= 0 Then
                                            indexColor = 7
                                        ElseIf indexColor >= 255 Then
                                            indexColor = 7
                                        End If
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                        nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                        strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                        lineTypeCode = ArrayRezLayer(1, k).Trim
                                    ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                        If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                            widthLineLayer = -1
                                        ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                            widthLineLayer = -2
                                        ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                            widthLineLayer = -3
                                        Else
                                            widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                        End If
                                    End If
                                Next k
                                Dim arrayMask As String() = lineTypeCode.Split(",")
                                If IsArray(arrayMask) = True Then
                                    If arrayMask.Length > 2 Then
                                        TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                    End If
                                End If
                            End If
                        End If
                        layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                    Else
                        layerObj = tempObjectLayer
                    End If
                End If
                'Имя объекта стиля
            ElseIf ArrayRez(0, i).Trim Like "Name" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    namePattern = ArrayRez(1, i).Trim
                End If
                'дескриптор объекта
            ElseIf ArrayRez(0, i).Trim Like "Description" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    deskObject = ArrayRez(1, i).Trim
                End If
                'тип линии
            ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    nameTypeLine = ArrayRez(1, i).Trim
                    If nameTypeLine Like "ПоСлою" Then
                        nameTypeLine = "ByLayer"
                    ElseIf nameTypeLine Like "ПоБлоку" Then
                        nameTypeLine = "ByBlock"
                    End If
                End If
                'код типа линии
            ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                End If
                'цвет
            ElseIf ArrayRez(0, i).Trim Like "Color" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                    If ColorLine < 0 Then ColorLine = 0
                    If ColorLine > 255 Then ColorLine = 255
                    colorObj = New CadColor(ColorLine)
                End If
            ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    If ArrayRez(1, i).Trim Like "ПоСлою" Then
                        WidthLineObj = -1
                    ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                        WidthLineObj = -2
                    ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                        WidthLineObj = -3
                    Else
                        WidthLineObj = Val(ArrayRez(1, i).Trim)
                        WidthLineObj = Math.Abs(WidthLineObj)
                    End If
                End If
            ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    ScaleLineObj = Val(ArrayRez(1, i).Trim)
                    If ScaleLineObj <= 0 Then
                        ScaleLineObj = 1
                    End If
                End If
            ElseIf ArrayRez(0, i).Trim Like "Scale" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    ScaleObject = Val(ArrayRez(1, i).Trim)
                End If
            ElseIf ArrayRez(0, i).Trim Like "Angle" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    AngleObject = Val(ArrayRez(1, i).Trim)
                End If
            ElseIf ArrayRez(0, i).Trim Like "PS" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    PSTable = ArrayRez(1, i).Trim
                End If
            ElseIf ArrayRez(0, i).Trim Like "NOTE" Then
                If IsNothing(ArrayRez(1, i)) = False Then
                    note = ArrayRez(1, i).Trim
                End If
            End If
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для блоков
    Public Shared Function FuncCreateStylePointObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef ScaleObject As Double = 1, Optional ByRef AngleObject As Double = 0, Optional ByRef attributeHeight As Double = 2.5, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean
        FuncCreateStylePointObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            Try
                If ArrayRez(0, i).Trim Like "Layer" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim NameLayer As String = ArrayRez(1, i).Trim
                        Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(tempObjectLayer) = True Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = "Непрерывный"
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                            strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                            lineTypeCode = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                    End If
                    'Имя объекта стиля
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameObject = ArrayRez(1, i).Trim
                    End If
                    'дескриптор объекта
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                    'тип линии
                ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameTypeLine = ArrayRez(1, i).Trim
                        If nameTypeLine Like "ПоСлою" Then
                            nameTypeLine = "ByLayer"
                        ElseIf nameTypeLine Like "ПоБлоку" Then
                            nameTypeLine = "ByBlock"
                        End If
                    End If
                    'код типа линии
                ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                    End If
                    'цвет
                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                        If ColorLine < 0 Then ColorLine = 0
                        If ColorLine > 255 Then ColorLine = 255
                        colorObj = New CadColor(ColorLine)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Scale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleObject = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Angle" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        AngleObject = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "AttributeHeight" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        attributeHeight = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "NOTE" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        Return True
    End Function

    '===========================================================================================================================
    'функция создает все необходимые стили для выносок
    Public Shared Function FuncCreateStyleLeaderObject(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByVal ArrayRez As String(,), ByRef nameObject As String, ByRef layerObj As DwgLayer, ByRef nameTypeLineObj As DwgLinetype, ByRef colorObj As CadColor, Optional ByRef ScaleLineObj As Double = 1, Optional ByRef WidthLineObj As Double = 20, Optional ByRef textStyle As String = "Standard", Optional ByRef hText As Double = 2, Optional ByRef nameArrow As String = "Нет", Optional ByRef sizeArrow As Double = 2, Optional ByRef PSTable As String = "OBJECT", Optional ByRef note As String = "") As Boolean

        FuncCreateStyleLeaderObject = False
        Dim nameTypeLine As String = "Continuous"
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(ArrayRez) = False Then Return False
        For i As Integer = 0 To ArrayRez.GetUpperBound(1)
            'создаем слой
            Try
                If ArrayRez(0, i).Trim Like "Layer" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim NameLayer As String = ArrayRez(1, i).Trim
                        Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, NameLayer)
                        'слой не найден, ищем его в шаблоне и создаем его
                        If IsNothing(tempObjectLayer) = True Then
                            Dim ArrayRezLayer As String(,) = Nothing
                            If IO.File.Exists(namePutchXML) = True Then
                                Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, NameLayer, ArrayRezLayer)
                            End If
                            Dim indexColor As Integer = 7
                            Dim nameLayerTypeLine As String = "Continuous"
                            Dim strDeskTypeLine As String = "Непрерывный"
                            Dim lineTypeCode As String = ""
                            Dim widthLineLayer As Integer = 20
                            Dim TypeLineLayer As DwgLinetype = Nothing
                            If IsArray(ArrayRezLayer) = True Then
                                If ArrayRezLayer.GetUpperBound(1) > 0 Then
                                    For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                                        If ArrayRezLayer(0, k).Trim Like "Color" Then
                                            indexColor = Val(ArrayRezLayer(1, k).Trim)
                                            If indexColor <= 0 Then
                                                indexColor = 7
                                            ElseIf indexColor >= 255 Then
                                                indexColor = 7
                                            End If
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                            nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                            strDeskTypeLine = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                            lineTypeCode = ArrayRezLayer(1, k).Trim
                                        ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                            If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                                widthLineLayer = -1
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                                widthLineLayer = -2
                                            ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                                widthLineLayer = -3
                                            Else
                                                widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                            End If
                                        End If
                                    Next k
                                    Dim arrayMask As String() = lineTypeCode.Split(",")
                                    If IsArray(arrayMask) = True Then
                                        If arrayMask.Length > 2 Then
                                            TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                                        End If
                                    End If
                                End If
                            End If
                            layerObj = RoburFunc.FuncAddLayer(ActivDocument, NameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                        Else
                            layerObj = tempObjectLayer
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "Name" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameObject = ArrayRez(1, i).Trim
                    End If
                    'дескриптор объекта
                ElseIf ArrayRez(0, i).Trim Like "Description" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        deskObject = ArrayRez(1, i).Trim
                    End If
                    'тип линии
                ElseIf ArrayRez(0, i).Trim Like "LineType" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        nameTypeLine = ArrayRez(1, i).Trim
                        If nameTypeLine Like "ПоСлою" Then
                            nameTypeLine = "ByLayer"
                        ElseIf nameTypeLine Like "ПоБлоку" Then
                            nameTypeLine = "ByBlock"
                        End If
                    End If
                    'код типа линии
                ElseIf ArrayRez(0, i).Trim Like "LineTypeCode" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        arrayMaskLine = ArrayRez(1, i).Trim.Split(",")
                    End If
                    'цвет
                ElseIf ArrayRez(0, i).Trim Like "Color" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        Dim ColorLine As Integer = Val(ArrayRez(1, i).Trim)
                        If ColorLine < 0 Then ColorLine = 0
                        If ColorLine > 255 Then ColorLine = 255
                        colorObj = New CadColor(ColorLine)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineWidth" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        If ArrayRez(1, i).Trim Like "ПоСлою" Then
                            WidthLineObj = -1
                        ElseIf ArrayRez(1, i).Trim Like "ПоБлоку" Then
                            WidthLineObj = -2
                        ElseIf ArrayRez(1, i).Trim Like "ПоУмолчанию" Then
                            WidthLineObj = -3
                        Else
                            WidthLineObj = Val(ArrayRez(1, i).Trim)
                            WidthLineObj = Math.Abs(WidthLineObj)
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "LineScale" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        ScaleLineObj = Val(ArrayRez(1, i).Trim)
                        If ScaleLineObj <= 0 Then
                            ScaleLineObj = 1
                        End If
                    End If
                ElseIf ArrayRez(0, i).Trim Like "TextStyle" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        textStyle = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "TextHeight" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        hText = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "ArrowSize" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        sizeArrow = Val(ArrayRez(1, i).Trim)
                    End If
                ElseIf ArrayRez(0, i).Trim Like "PS" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        PSTable = ArrayRez(1, i).Trim
                    End If
                ElseIf ArrayRez(0, i).Trim Like "NOTE" Or ArrayRez(0, i).Trim Like "Note" Or ArrayRez(0, i).Trim Like "note" Then
                    If IsNothing(ArrayRez(1, i)) = False Then
                        note = ArrayRez(1, i).Trim
                    End If
                End If
            Catch ex As System.Exception
            End Try
        Next i
        'создаем тип линии
        If IsArray(arrayMaskLine) = True Then
            If arrayMaskLine.Length > 2 Then
                nameTypeLineObj = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameTypeLine, deskObject, arrayMaskLine, namePutchXML)
            End If
        End If
        If nameObject.Trim.Length > 0 Then
            Dim arrayPropertiesMLeader As String(,) = Nothing
            Dim boolStyleMleader As Boolean = FuncXML.FuncReadMLederStylePropertiesToXML(namePutchXML, nameObject, arrayPropertiesMLeader)
            If IsArray(arrayPropertiesMLeader) = True Then
                For i As Integer = 0 To arrayPropertiesMLeader.GetUpperBound(1)
                    Dim field As String = arrayPropertiesMLeader(0, i)
                    Dim valStr As String = arrayPropertiesMLeader(1, i)
                    If field Like "NameArrow" Then
                        nameArrow = valStr
                        Exit For
                    End If
                Next i
            End If
        End If
        Return True
    End Function


    '===========================================================================================================================
    'функция создает слой из XML
    Public Shared Function FuncCreateLayerByXML(ByVal ActivDocument As Drawing, ByVal namePutchXML As String, ByRef nameLayer As String) As DwgLayer
        FuncCreateLayerByXML = Nothing
        Dim layerObj As DwgLayer = Nothing
        'создаем слой
        Try
            Dim tempObjectLayer As DwgLayer = RoburFunc.FuncFindLayerDwg(ActivDocument, nameLayer)
            'слой не найден, ищем его в шаблоне и создаем его
            If IsNothing(tempObjectLayer) = True Then
                Dim ArrayRezLayer As String(,) = Nothing
                If IO.File.Exists(namePutchXML) = True Then
                    Dim boolFindLayerXML As Boolean = FuncXML.FuncReadPropertiesLayerToXML(namePutchXML, nameLayer, ArrayRezLayer)
                End If
                Dim indexColor As Integer = 7
                Dim nameLayerTypeLine As String = "Continuous"
                Dim strDeskTypeLine As String = "Непрерывный"
                Dim lineTypeCode As String = ""
                Dim widthLineLayer As Integer = 20
                Dim TypeLineLayer As DwgLinetype = Nothing
                If IsArray(ArrayRezLayer) = True Then
                    If ArrayRezLayer.GetUpperBound(1) > 0 Then
                        For k As Integer = 0 To ArrayRezLayer.GetUpperBound(1)
                            If ArrayRezLayer(0, k).Trim Like "Color" Then
                                indexColor = Val(ArrayRezLayer(1, k).Trim)
                                If indexColor <= 0 Then
                                    indexColor = 7
                                ElseIf indexColor >= 255 Then
                                    indexColor = 7
                                End If
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineType" Then
                                nameLayerTypeLine = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "DeskLineType" Then
                                strDeskTypeLine = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineTypeCode" Then
                                lineTypeCode = ArrayRezLayer(1, k).Trim
                            ElseIf ArrayRezLayer(0, k).Trim Like "LineWidth" Then
                                If ArrayRezLayer(1, k).Trim Like "ПоСлою" Then
                                    widthLineLayer = -1
                                ElseIf ArrayRezLayer(1, k).Trim Like "ПоБлоку" Then
                                    widthLineLayer = -2
                                ElseIf ArrayRezLayer(1, k).Trim Like "ПоУмолчанию" Then
                                    widthLineLayer = -3
                                Else
                                    widthLineLayer = Val(ArrayRezLayer(1, k).Trim)
                                End If
                            End If
                        Next k
                        Dim arrayMask As String() = lineTypeCode.Split(",")
                        If IsArray(arrayMask) = True Then
                            If arrayMask.Length > 2 Then
                                TypeLineLayer = FuncStyles.FuncCreateTypeLineByPattern(ActivDocument, nameLayerTypeLine, strDeskTypeLine, arrayMask, namePutchXML)
                            End If
                        End If
                    End If
                End If
                layerObj = RoburFunc.FuncAddLayer(ActivDocument, nameLayer, indexColor, nameLayerTypeLine, widthLineLayer)
                Return layerObj
            Else
                Return tempObjectLayer
            End If
        Catch ex As System.Exception
        End Try
    End Function
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    'функция возвращает слой
    Public Function getLayerDwg(ByVal NameLayer As String) As DwgLayer
        Dim result As DwgLayer = ActiveProjectDrawing.Layers.Item("0")
        If IsNothing(NameLayer) = False Then
            If NameLayer.Trim.Length > 0 Then
                Try
                    Dim DwgLayers As DwgLayers = ActiveProjectDrawing.Layers
                    For Each dwglayer As DwgLayer In DwgLayers
                        If dwglayer.Name Like NameLayer.Trim Then
                            result = dwglayer
                        End If
                    Next
                Catch ex As System.Exception
                End Try
            End If
        End If
        Return result
    End Function
End Class
