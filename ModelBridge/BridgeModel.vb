Imports Topomatic.FoundationClasses
Imports Topomatic.Stg

Public Class BridgeModel
    Inherits StateControllerObject
    Implements IStgSerializable

    Private m_ReadOnly As Boolean = False

    'Строковое значение
    Public StringValue As String = "Строка"

    'Булево значение
    Public BooleanValue As Boolean = False

    'Значение с плавающей точкой
    Public DoubleValue As Double = 10.5

    'Целое значение
    Public IntValue As Integer = 10

    'Список строковых значений
    Public ArrayValues As New List(Of String)()

    'флаг только для чтения
    Public Overrides Property [ReadOnly] As Boolean
        Get
            Return m_ReadOnly
        End Get
        Set(value As Boolean)
            m_ReadOnly = value
        End Set
    End Property

    'Загрузка из узла
    Public Sub LoadFromStg(node As StgNode) Implements IStgSerializable.LoadFromStg
        'Все значения загружаем с указанием значения по умолчанию
        BooleanValue = node.GetBoolean("BooleanValue", False)
        StringValue = node.GetString("StringValue", "Строка")
        DoubleValue = node.GetDouble("DoubleValue", 10.5)
        IntValue = node.GetInt32("IntValue", 10)
        ArrayValues.Clear()

        'При загрузке массива указывается тип составляющих массив значений
        Dim array = node.GetArray("ArrayValues", StgType.String)
        For i As Integer = 0 To array.Count - 1
            ArrayValues.Add(array.GetString(i))
        Next
    End Sub

    'Сохранение в узел
    Public Sub SaveToStg(node As StgNode) Implements IStgSerializable.SaveToStg
        'Сохраняем значения в узел
        node.AddBoolean("BooleanValue", BooleanValue)
        node.AddString("StringValue", StringValue)
        node.AddDouble("DoubleValue", DoubleValue)
        node.AddInt32("IntValue", IntValue)

        'Сохраняем массив с указанием типа значений
        Dim array = node.AddArray("ArrayValues", StgType.String)
        For i As Integer = 0 To ArrayValues.Count - 1
            array.AddString(ArrayValues(i))
        Next
    End Sub
End Class
