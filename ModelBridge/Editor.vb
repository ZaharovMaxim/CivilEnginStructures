Imports System.IO
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.ApplicationPlatform.Core.DocumentModelEditor
Imports Topomatic.Cad.View
Imports Topomatic.Stg
Public Class Editor
    Inherits PlanModelEditor

    ' Реализация загрузки модели по указанному пути, должна вернуть реализацию класса нашей модели
    Public Overrides Function LoadFromFile(fullpath As String) As Object
        ' Создаем экземпляр класса модели
        Dim model As New Model()

        ' Если fullpath null - то необходимо просто вернуть экземпляр класса модели, без загрузки данных
        If fullpath IsNot Nothing Then
            ' Создаем файловый поток
            Using stream As New FileStream(fullpath, FileMode.Open, FileAccess.Read, FileShare.Read)
                ' Создаем документ для работы с Topomatic.Stg
                Dim document As New StgDocument()
                ' Загружаем документ из потока в бинарном виде
                document.LoadFromStreamAsBinary(stream)
                ' Загружаем данные нашей модели из документа
                model.LoadFromStg(document.Body)
            End Using
        End If

        ' Всегда возвращаем экземпляр модели
        Return model
    End Function

    ' Реализация сохранения модели по указанному пути
    Public Overrides Sub SaveToFile(model As Object, fullpath As String)
        ' В качестве параметра model приходит наша модель данных
        Dim m As Model = TryCast(model, Model)

        If m IsNot Nothing Then
            ' Создаем файловый поток
            Using stream As New FileStream(fullpath, FileMode.Create, FileAccess.Write, FileShare.None)
                ' Создаем документ для работы с Topomatic.Stg
                Dim document As New StgDocument()
                ' Сохраняем данные нашей модели в документ
                m.SaveToStg(document.Body)
                ' Сохраняем документ в поток в бинарном виде
                document.SaveToStreamAsBinary(stream)
            End Using
        End If
    End Sub

    Protected Overrides Function CreatePlanLayer(model As IProjectModel) As CadViewLayer
        Dim m As Model = TryCast(model.LockRead(), Model)

        If m IsNot Nothing Then
            Return New ModelLayer() With {.Model = m}
        End If

        Return Nothing
    End Function

    Protected Overrides Sub ReloadModel(model As IProjectModel, editorResult As EditorResult)
        Dim m As Model = TryCast(model.LockRead(), Model)

        If m IsNot Nothing Then
            Dim layer As ModelLayer = TryCast(editorResult.PlanLayer, ModelLayer)
            layer.Model = m
        End If
    End Sub

    Protected Overrides Sub RemovePlanLayer(model As IProjectModel, layer As CadViewLayer)
        Dim model_layer As ModelLayer = TryCast(layer, ModelLayer)

        If model_layer IsNot Nothing Then
            model_layer.Model = Nothing
        End If
    End Sub
End Class
