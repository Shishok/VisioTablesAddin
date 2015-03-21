Public Class dlgLinkData

    Dim lngRowIDs() As Integer

    Private Sub dlgLinkData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim vsoDataRecordset As Visio.DataRecordset, arrDataRecordset() As String
        Dim i As Byte

        For i = 1 To vsoApp.ActiveDocument.DataRecordsets.Count
            vsoDataRecordset = vsoApp.ActiveDocument.DataRecordsets.Item(i)
            arrDataRecordset = Split(vsoDataRecordset.DataConnection.ConnectionString, ";")
            cmbSourseData.Items.Add(Strings.Right(arrDataRecordset(2), Len(arrDataRecordset(2)) - 12) & " - " & vsoDataRecordset.Name)
        Next

        cmbSourseData.SelectedIndex = 0
        txtNameTable.Enabled = ckbInsertName.Checked
        txtColEndSourse.Value = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).DataColumns.Count
        txtRowEndSourse.Value = UBound(lngRowIDs) + 1
        ckbInsertName.Checked = False
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        If vsoApp.ActiveWindow.Selection.Count = 0 Then GoTo err
        'Me.Hide()
        Call LinkToDataInShapes(cmbSourseData.SelectedIndex, ckbInsertName.Checked, txtNameTable.Text, _
        ckbTitleColumns.Checked, txtRowStartSourse.Value, ckbInvisibleZero.Checked, _
        UBound(lngRowIDs), vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).DataColumns.Count, ckbFontBold.Checked)

        Me.Close()
        Exit Sub

err:
        MsgBox("Необходимо выделить любую ячейку/ячейки в нужной таблице", 64, "Внимание!")

    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub cmdRefreshAll_Click(sender As Object, e As EventArgs) Handles cmdRefreshAll.Click
        Call RefreshDataInShapes()
    End Sub

    Private Sub cmbSourseData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSourseData.SelectedIndexChanged
        txtNameTable.Text = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).Name
        lngRowIDs = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).GetDataRowIDs("")
        lblCountRow.Text = "Источник содержит " & _
        vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).DataColumns.Count & " столбцов и " _
        & UBound(lngRowIDs) + 1 & " строк данных"
        txtColEndSourse.Value = vsoApp.ActiveDocument.DataRecordsets.Item(cmbSourseData.SelectedIndex + 1).DataColumns.Count
        txtRowEndSourse.Value = UBound(lngRowIDs) + 1
    End Sub

    Private Sub ckbInsertName_CheckedChanged(sender As Object, e As EventArgs) Handles ckbInsertName.CheckedChanged
        txtNameTable.Enabled = ckbInsertName.Checked
    End Sub

    Sub RefreshDataInShapes() ' Обновление источников внешних данных

        If vsoApp.ActiveDocument.DataRecordsets.Count = 0 Then
            Exit Sub
        End If

        Dim vsoDataRecordset As Visio.DataRecordset
        Dim i As Byte

        Call RecUndo("Обновить все данные")

        On Error GoTo ErrorHandler

        For i = 1 To vsoApp.ActiveDocument.DataRecordsets.Count
            vsoDataRecordset = vsoApp.ActiveDocument.DataRecordsets.Item(i)
            vsoDataRecordset.Refresh()
        Next
        Call RecUndo("0")

        MsgBox("Обновлено источников внешних данных - " & vsoApp.ActiveDocument.DataRecordsets.Count, vbInformation, "Обновление данных")
        Exit Sub

ErrorHandler:
        MsgBox("Внешний источник:" & vbCrLf & vsoDataRecordset.Name & vbCrLf & Err.Description, vbExclamation, "Обновление данных")
    End Sub

End Class