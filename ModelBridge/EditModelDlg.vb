Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports Topomatic.Cad.Foundation
Imports Topomatic.Controls.Dialogs
Partial Public Class EditModelDlg
    Private m_Model As BridgeModel
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub EditModelDlg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Инициализация при загрузке формы
        If m_Model IsNot Nothing Then
            tbStringValue.Text = m_Model.StringValue
            tbDoubleValue.Text = ValueConverter.FloatToStr(m_Model.DoubleValue)
            tbIntValue.Text = m_Model.IntValue.ToString()
            cbBooleanValue.Checked = m_Model.BooleanValue
            For i As Integer = 0 To m_Model.ArrayValues.Count - 1
                tbArrayValues.Text += m_Model.ArrayValues(i)
                If i < m_Model.ArrayValues.Count - 1 Then
                    tbArrayValues.Text += ";"c
                End If
            Next
        End If
    End Sub
    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        If CommitChanges() Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Function CommitChanges() As Boolean
        'При нажатии клавиши ОК присваиваем все значения
        If m_Model IsNot Nothing Then
            m_Model.StringValue = tbStringValue.Text

            ' Проверка Double значения
            Dim doubleResult As Double
            If Double.TryParse(tbDoubleValue.Text, doubleResult) Then
                m_Model.DoubleValue = doubleResult
            Else
                MessageBox.Show("Некорректное число с плавающей точкой", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            ' Проверка Int значения
            Dim intResult As Integer
            If Integer.TryParse(tbIntValue.Text, intResult) Then
                m_Model.IntValue = intResult
            Else
                MessageBox.Show("Некорректное целое число", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            m_Model.BooleanValue = cbBooleanValue.Checked
            m_Model.ArrayValues.Clear()

            Dim splitValues() As String = tbArrayValues.Text.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            m_Model.ArrayValues.AddRange(splitValues)
        End If

        Return True
    End Function

    'Статическая функция для вызова диалога
    Public Shared Function Execute(model As BridgeModel) As Boolean
        Using dlg As New EditModelDlg()
            dlg.m_Model = model
            Return dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK
        End Using
    End Function
End Class