Public Class dlgSortTable

    Const Cap = "Сортировка данных"

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim vsoObj As Visio.Selection = winObj.Selection
        Dim sh As Visio.Shape

        For Each sh In vsoObj
            If Not sh.Name Like "ClW*" Then vsoApp.ActiveWindow.Select(sh, 1)
        Next

        If vsoApp.ActiveWindow.Selection.Count > 1 Then
            Call SortTableData(Num_Column.Value, cb_DigitOrText.Checked, cb_SortingDirection.Checked)
        Else
            MsgBox("Нет выделенных ячеек." & vbNewLine & "Или выделенно меньше двух ячеек.", vbExclamation, Cap)
        End If
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Close()
    End Sub

    'Private Sub dlgSortTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    Num_Column.Maximum = SelColRow(1)
    'End Sub
End Class