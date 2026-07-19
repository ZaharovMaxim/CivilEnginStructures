Imports System.ComponentModel
Imports System.Windows
Imports System.Xml
Imports Topomatic
Imports Topomatic.Alg
Imports Topomatic.Alg.Runtime.Communications
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.Foundation.Brep
Imports Topomatic.Dwg.Entities
Imports Topomatic.FoundationClasses.Vcs
Imports Topomatic.Sfc
Imports Topomatic.Visualization
Imports Topomatic.Visualization.Runtime
' Балка мостовая двутавровая
Public Class BeamI
    Private _numberProlet As Integer          ' Номер пролета
    Private _numberRow As Integer             ' Номер ряда
    Private _lenght As Double                 ' Полная длина балки
    Private _height As Double                 ' Полная высота балки
    Private _heightTopPlate As Double         ' Толщина плиты
    Private _widthTopPlateLeft As Double      ' Ширина плиты балки по верху влево
    Private _widthTopPlateRight As Double     ' Ширина плиты балки по верху влево
    Private _widthTop As Double               ' Ширина стойки балки по по верху
    Private _widthBottom As Double            ' Ширина стойки балки по низу
    Private _horizontalOffsetRibZone As Double ' величина уширения стойки балки по низу
    Private _deltaRib As Double               ' уклон по верху стойки (1:....)
    Private _radiusTop As Double              ' верхний радиус
    Private _radiusBottom As Double           ' нижний радиус
    Private _gWidth As Double                 ' толщина стойки 
    Private _verticalOffsetRibZone As Double              ' высота ребра
    ' Параметры покрытия и монтажа
    Private _offsetSurface As Double          ' наименишая толщина покрытия над балкой
    Private _clearence As Double              ' Зазор с предыдущей балкой
    Private _axisOffset As Double             ' Смещение от оси в начале балки
    Private _rotationAxisBridge As Double     ' угол поворота балки (в случае если смещение от оси конца бали и начала балки различное)
    ' Параметры опирания
    Private _a As Double                      ' Расстояние от начала балки до точки опирания в начале
    Private _b As Double                      ' Расстояние от конца балки до точки опирания в конце
    ' Параметры монолитных участков
    Private _startLenghtMonolith As Double    ' Участок омоноличивания в начале балки
    Private _endLenghtMonolith As Double      ' Участок омоноличивания в конце балки
    ' Информационные параметры
    Private _model As String                  ' Модель/артикул балки
    Private _nameAlbum As String              ' Альбом/каталог
    Private _mass As Double                   ' Масса балки
    Private _volume As Double                 ' Объем балки (строковый тип, может содержать единицы измерения)
    Private _modelTLS As String               ' Имя модели TLC
    Public _elementBridgePoint As PointsCollections
    ' Инициализация класса
    Public Sub New()
        ' Установка значений по умолчанию
        _numberProlet = 0
        _numberRow = 0
        _lenght = 0
        _height = 1.23
        _heightTopPlate = 0.15
        _widthTopPlateLeft = 0.7
        _widthTopPlateRight = 0.7
        _widthTop = 0.76
        _widthBottom = 0.59
        _horizontalOffsetRibZone = 0.01
        _deltaRib = 1
        _radiusTop = 300
        _radiusBottom = 200
        _gWidth = 0.16
        _verticalOffsetRibZone = 0.1
        _a = 0.3
        _b = 0.3
        _startLenghtMonolith = 0
        _endLenghtMonolith = 0

        _offsetSurface = 0
        _clearence = 0
        _axisOffset = 0
        _rotationAxisBridge = 0

        _model = ""
        _nameAlbum = ""
        _mass = 0
        _volume = 0
        _modelTLS = ""
        _elementBridgePoint = New PointsCollections
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
    ' Номер ряда
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
    ' Полная длина балки
    <Browsable(True)>
    <Description("Полная длина балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Длина балки")>
    <[ReadOnly](True)>
    Public Property lenght() As Double
        Get
            Return _lenght
        End Get
        Set(value As Double)
            _lenght = value
        End Set
    End Property
    ' полная высота балки
    <Browsable(True)>
    <Description("Полная высота балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Высота балки")>
    <[ReadOnly](True)>
    Public Property height() As Double
        Get
            Return _height
        End Get
        Set(value As Double)
            _height = value
        End Set
    End Property
    ' высота верхней плиты
    <Browsable(True)>
    <Description("Толщина верхней плиты балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Толщина плиты")>
    <[ReadOnly](True)>
    Public Property heightTopPlate() As Double
        Get
            Return _heightTopPlate
        End Get
        Set(value As Double)
            _heightTopPlate = value
        End Set
    End Property
    ' ширина верхней плиты
    <Browsable(True)>
    <Description("Ширина плиты балки влево от оси, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Ширина плиты влево")>
    <[ReadOnly](True)>
    Public Property widthTopPlateLeft() As Double
        Get
            Return _widthTopPlateLeft
        End Get
        Set(value As Double)
            _widthTopPlateLeft = value
        End Set
    End Property
    ' ширина верхней плиты
    <Browsable(True)>
    <Description("Ширина плиты балки вправо от оси, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Ширина плиты вправо")>
    <[ReadOnly](True)>
    Public Property widthTopPlateRight() As Double
        Get
            Return _widthTopPlateRight
        End Get
        Set(value As Double)
            _widthTopPlateRight = value
        End Set
    End Property
    ' ширина стойки балки по верху
    <Browsable(True)>
    <Description("Ширина стойки балки по верху под плитой, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Ширина верха стойки")>
    <[ReadOnly](True)>
    Public Property WidthTop() As Double
        Get
            Return _widthTop
        End Get
        Set(value As Double)
            _widthTop = value
        End Set
    End Property
    ' ширина стойки по низу
    <Browsable(True)>
    <Description("Ширина основания стойки балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Ширина основания стойки")>
    <[ReadOnly](True)>
    Public Property widthBottom() As Double
        Get
            Return _widthBottom
        End Get
        Set(value As Double)
            _widthBottom = value
        End Set
    End Property
    ' уширение стойки балки по низу
    <Browsable(True)>
    <Description("Величина уширения основания стойки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Уширение стойки")>
    <[ReadOnly](True)>
    Public Property HorizontalOffsetRibZone() As Double
        Get
            Return _horizontalOffsetRibZone
        End Get
        Set(value As Double)
            _horizontalOffsetRibZone = value
        End Set
    End Property
    ' уклон по верху основания балки
    <Browsable(True)>
    <Description("Относительный уклон ребра балки, 1:")>
    <Category("Свойства сооружения")>
    <DisplayName("Уклон ребра балки")>
    <[ReadOnly](True)>
    Public Property DeltaRib() As Double
        Get
            Return _deltaRib
        End Get
        Set(value As Double)
            _deltaRib = value
        End Set
    End Property
    ' верхний радиус
    <Browsable(True)>
    <Description("Верхний радиус, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Верхний радиус")>
    <[ReadOnly](True)>
    Public Property radiusTop() As Double
        Get
            Return _radiusTop
        End Get
        Set(value As Double)
            _radiusTop = value
        End Set
    End Property
    ' нижний радиус
    <Browsable(True)>
    <Description("Нижний радиус, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Нижний радиус")>
    <[ReadOnly](True)>
    Public Property radiusBottom() As Double
        Get
            Return _radiusBottom
        End Get
        Set(value As Double)
            _radiusBottom = value
        End Set
    End Property
    ' толщина стойки всередине
    <Browsable(True)>
    <Description("Толщина стойки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Толщина стойки")>
    <[ReadOnly](True)>
    Public Property gWidth() As Double
        Get
            Return _gWidth
        End Get
        Set(value As Double)
            _gWidth = value
        End Set
    End Property
    ' высота ребра балки
    <Browsable(True)>
    <Description("Высота ребра стойки балки в зоне уширения, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Высота ребра балки")>
    <[ReadOnly](True)>
    Public Property VerticalOffsetRibZone() As Double
        Get
            Return _verticalOffsetRibZone
        End Get
        Set(value As Double)
            _verticalOffsetRibZone = value
        End Set
    End Property
    ' расстояние от начала балки до точки опирания
    <Browsable(True)>
    <Description("Расстояние от начала балки до точки опирания (A), м")>
    <Category("Свойства сооружения")>
    <DisplayName("Точка опирания A")>
    <[ReadOnly](True)>
    Public Property a() As Double
        Get
            Return _a
        End Get
        Set(value As Double)
            _a = value
        End Set
    End Property
    ' расстояние от конца балки до точки опирания
    <Browsable(True)>
    <Description("Расстояние от конца балки до точки опирания (B), м")>
    <Category("Свойства сооружения")>
    <DisplayName("Точка опирания B")>
    <[ReadOnly](True)>
    Public Property b() As Double
        Get
            Return _b
        End Get
        Set(value As Double)
            _b = value
        End Set
    End Property
    ' участок омоноличивания внвчале балки
    <Browsable(True)>
    <Description("Участок омоноличивания начала балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Участок омоноличивания начала балки")>
    <[ReadOnly](True)>
    Public Property startLenghtMonolith() As Double
        Get
            Return _startLenghtMonolith
        End Get
        Set(value As Double)
            _startLenghtMonolith = value
        End Set
    End Property
    ' участок омоноличивания вконце балки
    <Browsable(True)>
    <Description("Участок омоноличивания конца балки, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Участок омоноличивания конца балки")>
    <[ReadOnly](True)>
    Public Property endLenghtMonolith() As Double
        Get
            Return _endLenghtMonolith
        End Get
        Set(value As Double)
            _endLenghtMonolith = value
        End Set
    End Property
    ' участок омоноличивания вконце балки
    <Browsable(True)>
    <Description("Толщина покрытия над балкой, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Толщина покрытия над балкой")>
    Public Property offsetSurface() As Double
        Get
            Return _offsetSurface
        End Get
        Set(value As Double)
            _offsetSurface = value
        End Set
    End Property
    ' минимальный зазор с перыдущей балкой
    <Browsable(True)>
    <Description("Минимальный зазор с предыдущей балкой, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Минимальный зазор")>
    Public Property clearence() As Double
        Get
            Return _clearence
        End Get
        Set(value As Double)
            _clearence = value
        End Set
    End Property
    ' смещение от оси
    <Browsable(True)>
    <Description("Смещение от оси трассы, м")>
    <Category("Свойства сооружения")>
    <DisplayName("Смещение от оси трассы")>
    Public Property axisOffset() As Double
        Get
            Return _axisOffset
        End Get
        Set(value As Double)
            _axisOffset = value
        End Set
    End Property
    ' смещение от оси
    <Browsable(True)>
    <Description("Угол отклонения конца балки от оси, град")>
    <Category("Свойства сооружения")>
    <DisplayName("Угол отклонения")>
    Public Property RotationAxisBridge() As Double
        Get
            Return _rotationAxisBridge
        End Get
        Set(value As Double)
            _rotationAxisBridge = value
        End Set
    End Property
    ' имя модели балки
    <Browsable(True)>
    <Description("Имя модели балки")>
    <Category("Свойства сооружения")>
    <DisplayName("Имя модели балки")>
    Public Property model() As String
        Get
            Return _model
        End Get
        Set(value As String)
            _model = value
        End Set
    End Property
    ' имя альбома балок
    <Browsable(True)>
    <Description("Имя альбома балок")>
    <Category("Свойства сооружения")>
    <DisplayName("Имя альбома балок")>
    Public Property nameAlbum() As String
        Get
            Return _nameAlbum
        End Get
        Set(value As String)
            _nameAlbum = value
        End Set
    End Property
    ' масса
    <Browsable(True)>
    <Description("Масса балки, г")>
    <Category("Свойства сооружения")>
    <DisplayName("Масса балки")>
    <[ReadOnly](True)>
    Public Property mass() As Double
        Get
            Return _mass
        End Get
        Set(value As Double)
            _mass = value
        End Set
    End Property
    ' объем
    <Browsable(True)>
    <Description("Объем балки, куб.м")>
    <Category("Свойства сооружения")>
    <DisplayName("Объем балки")>
    <[ReadOnly](True)>
    Public Property volume() As Double
        Get
            Return _volume
        End Get
        Set(value As Double)
            _volume = value
        End Set
    End Property
    ' имя модели TLC
    Public Property modelTLS() As String
        Get
            Return _modelTLS
        End Get
        Set(value As String)
            _modelTLS = value
        End Set
    End Property
    'создать класс БАЛКА ДВУТАВРОВАЯ
    Public Shared Function createAxis(ByVal idBridge As String) As StructureElement
        Dim elementAxis As StructureElement = New StructureElement()
        elementAxis.Label = "Мосты и путепроводы"
        elementAxis.ClassBridgeObject = StructureElement.classBridge.SpanStructures
        elementAxis.ClassObject = StructureElement.classStructure.BeamI
        elementAxis.Name = StructureElement.typeObject.axisBeam
        Dim deskObject As String = StructureElement.GetDescription(StructureElement.typeObject.axisBeam)
        elementAxis.Description = deskObject
        elementAxis.KeyParameter = ""
        elementAxis.IdElement = Guid.NewGuid.ToString
        elementAxis.IdStructure = idBridge
        elementAxis.Note = ""
        elementAxis.DWGEntity = New DwgLine
        Return elementAxis
    End Function



    'чтение характеристик двутавровой балки из файла xml
    Public Function setPropertiesFromXML(ByVal fullPatchXML As String, ByVal modelBeam As String) As BeamI
        Dim userBeam As BeamI = New BeamI
        If IO.File.Exists(fullPatchXML) = True Then
            Dim xDoc As XmlDocument = New XmlDocument()
            Dim reader As XmlTextReader = New XmlTextReader(fullPatchXML)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Beam" Then
                            If reader.HasAttributes = True Then
                                Dim boolFindBeam As Boolean = False
                                While (reader.MoveToNextAttribute())
                                    Dim name As String = reader.Name
                                    Dim valN As String = reader.Value
                                    If name Like "model" Then
                                        If Not (valN Like modelBeam) Then
                                            Exit While
                                        Else
                                            userBeam.model = valN
                                            userBeam.nameAlbum = IO.Path.GetFileNameWithoutExtension(fullPatchXML)
                                            boolFindBeam = True
                                        End If
                                    ElseIf name Like "fullLenght" Then
                                        userBeam.lenght = Val(valN) / 1000
                                    ElseIf name Like "widthTopPlateLeft" Then
                                        userBeam.widthTopPlateLeft = Val(valN) / 1000
                                    ElseIf name Like "widthTopPlateRight" Then
                                        userBeam.widthTopPlateRight = Val(valN) / 1000
                                    ElseIf name Like "heightTopPlate" Then
                                        userBeam.heightTopPlate = Val(valN) / 1000
                                    ElseIf name Like "WidthTop" Then
                                        userBeam.WidthTop = Val(valN) / 1000
                                    ElseIf name Like "widthBottom" Then
                                        userBeam.widthBottom = Val(valN) / 1000
                                    ElseIf name Like "height" Then
                                        userBeam.height = Val(valN) / 1000
                                    ElseIf name Like "a" Then
                                        userBeam.a = Val(valN) / 1000
                                    ElseIf name Like "b" Then
                                        userBeam.b = Val(valN) / 1000
                                    ElseIf name Like "verticalOffsetRibZone" Then
                                        userBeam.VerticalOffsetRibZone = Val(valN) / 1000
                                    ElseIf name Like "horizontalOffsetRibZone" Then
                                        userBeam.HorizontalOffsetRibZone = Val(valN) / 1000
                                    ElseIf name Like "deltaRib" Then
                                        userBeam.DeltaRib = Val(valN)
                                    ElseIf name Like "radiusTop" Then
                                        userBeam.radiusTop = Val(valN) / 1000
                                    ElseIf name Like "radiusBottom" Then
                                        userBeam.radiusBottom = Val(valN) / 1000
                                    ElseIf name Like "gWidth" Then
                                        userBeam.gWidth = Val(valN) / 1000
                                    ElseIf name Like "mass" Then
                                        userBeam.mass = Val(valN)
                                    ElseIf name Like "volume" Then
                                        userBeam.volume = Val(valN)
                                    ElseIf name Like "modelTLS" Then
                                        userBeam.modelTLS = valN
                                    End If
                                End While
                                If boolFindBeam = True Then
                                    Exit While
                                End If
                            End If
                        End If
                End Select
            End While
            reader.Close()
        End If
        Return userBeam
    End Function
    'расчет сечения балки (0-сечение без учета плиты балки, 1-сечение полное)
    Public Function getSection() As Dictionary(Of Integer, List(Of Vector2D))
        Dim polySection As List(Of Vector2D) = New List(Of Vector2D)
        Dim polyShortSection As List(Of Vector2D) = New List(Of Vector2D)
        'центр стойки балки по низу
        Dim x As Double = 0
        Dim y As Double = 0
        polySection.Add(New Vector2D(x, y))
        polyShortSection.Add(New Vector2D(x, y))
        'край стойки балки по низу влево
        x = x + -1 * (widthBottom / 2)
        polySection.Add(New Vector2D(x, y))
        polyShortSection.Add(New Vector2D(x, y))
        'левый край стойки балки по низу с учатом высоты ребра и уширения
        x = x + -1 * HorizontalOffsetRibZone
        y = y + _verticalOffsetRibZone
        polySection.Add(New Vector2D(x, y))
        polyShortSection.Add(New Vector2D(x, y))
        'наклонное ребро (предварительно)
        Dim lRib As Double = widthBottom / 2 + HorizontalOffsetRibZone - gWidth / 2
        Dim hrib As Double = lRib * DeltaRib
        x = x + lRib
        y = y + hrib
        'вписываем нижнюю дугу
        Dim endPoint1 As Vector2D = New Vector2D(x, y)
        Dim endPoint2 As Vector2D = New Vector2D(x, y + height)
        Dim listPoint As List(Of Vector2D) = MathFunction.InscribeCircleBetweenSegments(polySection.Item(2), endPoint1, endPoint1, endPoint2, radiusBottom, 5)
        If IsNothing(listPoint) = False Then
            For i As Integer = 0 To listPoint.Count - 1
                polySection.Add(listPoint.Item(i))
                polyShortSection.Add(listPoint.Item(i))
            Next i
        End If
        'вписываем верхнюю дугу
        Dim startPoint1 As Vector2D = polySection.Last
        endPoint1 = New Vector2D(startPoint1.X, height - heightTopPlate)
        endPoint2 = New Vector2D(startPoint1.X - (widthTopPlateLeft - gWidth / 2), height - heightTopPlate)
        listPoint = MathFunction.InscribeCircleBetweenSegments(startPoint1, endPoint1, endPoint1, endPoint2, radiusTop, 5)
        If IsNothing(listPoint) = False Then
            For i As Integer = 0 To listPoint.Count - 1
                polySection.Add(listPoint.Item(i))
                polyShortSection.Add(listPoint.Item(i))
            Next i
        End If
        'рисуем плиту влево по низу
        Dim lBottomPlate As Double = widthTopPlateLeft - gWidth / 2
        x = endPoint1.X - lBottomPlate
        y = endPoint1.Y
        polySection.Add(New Vector2D(x, y))
        y = y + heightTopPlate
        polySection.Add(New Vector2D(x, y))
        x = x + widthTopPlateLeft
        polySection.Add(New Vector2D(x, y))
        x = x + widthTopPlateRight
        polySection.Add(New Vector2D(x, y))
        y = y - heightTopPlate
        polySection.Add(New Vector2D(x, y))
        'вставляем третью дугу
        startPoint1 = polySection.Last
        endPoint1 = New Vector2D(startPoint1.X - (widthTopPlateRight - gWidth / 2), y)
        endPoint2 = New Vector2D(startPoint1.X - (widthTopPlateRight - gWidth / 2), y - height)
        listPoint = MathFunction.InscribeCircleBetweenSegments(startPoint1, endPoint1, endPoint1, endPoint2, radiusTop, 5)
        If IsNothing(listPoint) = False Then
            For i As Integer = 0 To listPoint.Count - 1
                polySection.Add(listPoint.Item(i))
                polyShortSection.Add(listPoint.Item(i))
            Next i
        End If
        'вставляет четвертую дугу
        startPoint1 = polySection.Last
        endPoint2 = New Vector2D(widthBottom / 2 + HorizontalOffsetRibZone, VerticalOffsetRibZone)
        endPoint1 = New Vector2D(endPoint2.X - lRib, endPoint2.Y + hrib)
        listPoint = MathFunction.InscribeCircleBetweenSegments(startPoint1, endPoint1, endPoint1, endPoint2, radiusBottom, 5)
        If IsNothing(listPoint) = False Then
            For i As Integer = 0 To listPoint.Count - 1
                polySection.Add(listPoint.Item(i))
                polyShortSection.Add(listPoint.Item(i))
            Next i
        End If
        polySection.Add(New Vector2D(widthBottom / 2 + HorizontalOffsetRibZone, VerticalOffsetRibZone))
        polyShortSection.Add(New Vector2D(widthBottom / 2 + HorizontalOffsetRibZone, VerticalOffsetRibZone))
        polySection.Add(New Vector2D(widthBottom / 2, 0))
        polyShortSection.Add(New Vector2D(widthBottom / 2, 0))

        Dim dictSections As Dictionary(Of Integer, List(Of Vector2D)) = New Dictionary(Of Integer, List(Of Vector2D))
        dictSections.Add(1, polyShortSection)
        dictSections.Add(2, polySection)
        Return dictSections
    End Function
    'функция возвращает точку контура по ее коду
    Public Function getPointBeam(ByVal code As String, Optional pointTop As Boolean = True) As Vector3D
        Dim result As Vector3D = New Vector3D
        Dim listPoint As Dictionary(Of Integer, PointStructure) = _elementBridgePoint.ListPointModel
        If IsNothing(listPoint) = True Then Return result
        If listPoint.Count = 0 Then Return result
        If IsNothing(code) = True Then Return result
        For i As Integer = 0 To listPoint.Count - 1
            Dim ptStructure As PointStructure = listPoint.ElementAt(i).Value
            If ptStructure.Code.Trim Like code.Trim Then
                If pointTop = True Then
                    result = New Vector3D(ptStructure.X, ptStructure.Y, ptStructure.Z)
                Else
                    result = New Vector3D(ptStructure.X - ptStructure.dx, ptStructure.Y - ptStructure.dy, ptStructure.Z + ptStructure.dz)
                End If
                Exit For
            End If
        Next i
        Return result
    End Function
End Class

