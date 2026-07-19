Imports Topomatic.ApplicationPlatform
Imports Topomatic.ApplicationPlatform.Core
Imports Topomatic.Arrangements

Public Class FormReportBeams
    Public userActivDocument As Topomatic.Dwg.Drawing = Nothing
    Public showDialog1 As Boolean = False

    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayNameReport As String() = {"Отчет по точкам опирания балок", "Отчет по деформационным зазорам", "Ведомость по осям опор", "Ведомость верха плиты балок и толщины покрытия", "Экспорт элементов в dwg"}
        ' Добавить код инициализации после вызова InitializeComponent().
        ComboBox4.DataSource = arrayNameReport
    End Sub
    Private Sub FormReportBeams_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        showDialog1 = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        showDialog1 = True
        Me.Hide()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ComboBox3.Enabled = False
        Else
            ComboBox3.Enabled = True
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбрать модель
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Получение всех моделей проекта
        Dim Project As ModelProject = ApplicationHost.Current.ActiveProject
        Dim userSubObjectBearm As ArrangementModel = New ArrangementModel()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub
End Class