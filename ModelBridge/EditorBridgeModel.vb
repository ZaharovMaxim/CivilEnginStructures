Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Stg

Public Class EditorBridgeModel
    Inherits ModelEditor
    'Реализация загрузки модели по указанному пути, должна вернуть реализацию класса нашей модели
    Public Overrides Function LoadFromFile(fullpath As String) As Object
        'создаем экземпляр класса модели
        Dim model = New BridgeModel()
        'если fullpath null - то необходимо просто вернуть экземпляр класса модели, без загрузки данных
        If fullpath IsNot Nothing Then
            'создаем файловый поток
            Using stream = New FileStream(fullpath, FileMode.Open, FileAccess.Read, FileShare.Read)
                'создаем документ для работы с Topomatic.Stg
                Dim document = New StgDocument()
                'загружаем документ из потока в бинарном виде
                document.LoadFromStreamAsBinary(stream)
                'загружаем данные нашей модели из документа
                model.LoadFromStg(document.Body)
            End Using
        End If
        'всегда возвращаем экземпляр модели
        Return model
    End Function

    'Реализация сохранения модели по указанному пути
    Public Overrides Sub SaveToFile(model As Object, fullpath As String)
        'в качестве параметра model приходит наша модель данных
        Dim m = TryCast(model, BridgeModel)
        If m IsNot Nothing Then
            'создаем файловый поток
            Using stream = New FileStream(fullpath, FileMode.Create, FileAccess.Write, FileShare.None)
                'создаем документ для работы с Topomatic.Stg
                Dim document = New StgDocument()
                'сохраняем данные нашей модели в документ
                m.SaveToStg(document.Body)
                'сохраняем документ в потока в бинарном виде
                document.SaveToStreamAsBinary(stream)
            End Using
        End If
    End Sub

    'Реализация открытия модели по команде "open"
    Public Overrides Function Open(model As IProjectModel) As IEditorResult
        Dim cursor = System.Windows.Forms.Cursor.Current
        Cursor.Current = Cursors.WaitCursor
        Try
            'В нашем поросто возвращаем реализацию интерфеса IEditorResult
            Return New EditorResult()
        Finally
            Cursor.Current = cursor
        End Try
    End Function

    'Реализация интерфейса IEditorResult
    Private Class EditorResult
        Implements IEditorResult

        Private m_Opened As Boolean

        Public Sub New()
            m_Opened = True
        End Sub

        'Необходимо реализовать флаг, показывающий открыта модель или нет
        Public ReadOnly Property Opened As Boolean Implements IEditorResult.Opened
            Get
                Return m_Opened
            End Get
        End Property

        'Необходимо реализовать метод закрытия модели
        Public Sub Close() Implements IEditorResult.Close
            'В нашем случае мы просто управляем флагом и все
            m_Opened = False
        End Sub

        'И метод перезагрузки модели
        Public Sub Reload() Implements IEditorResult.Reload
            'Здесь нам ничего не нужно делать
        End Sub
    End Class
End Class

