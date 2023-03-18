Public Class dlgSelectFromList

    Dim sh As Visio.Selection

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Dim i As Integer

        Call RecUndo("Вставить из списка")

        For i = 1 To sh.Count ' Вставка выбранного значения в выделенные ячейки
            'If InStr(1, sh(i).NameU, "ClW", 1) <> 0 Then
            sh(i).Characters.Text = cmbSelectValue.Text
            'End If
        Next

        Call RecUndo("0")

        'winObj.Selection = sh
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub dlgSelectFromList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call InitArrShapeID(NT)
        sh = winObj.Selection
        Call LoadData()
    End Sub

    Private Sub LoadData()
        Dim NoDupes As New Collection, Shp As Visio.Shape
        Dim intStartCol As Integer = sh(1).Cells(UTC).Result(""), intStartRow As Integer = 1
        Dim intEndCol As Integer = intStartCol, intEndRow As Integer = UBound(ArrShapeID, 2)

        If rdbFromRow.Checked Then
            intStartCol = 1 : intEndCol = UBound(ArrShapeID, 1)
            intStartRow = sh(1).Cells(UTR).Result("") : intEndRow = intStartRow
        End If

        On Error Resume Next
        For i = intStartCol To intEndCol
            For j = intStartRow To intEndRow
                With winObj.Page.Shapes.ItemFromID(ArrShapeID(i, j))
                    NoDupes.Add(.Characters.Text, CStr(.Characters.Text))
                End With
            Next
        Next
        On Error GoTo 0

        cmbSelectValue.Items.Clear()
        For i = 1 To NoDupes.Count ' Добавление коллекции в ComboBox (cmbSelectValue)
            If StrComp(NoDupes(i), Trim("")) <> 0 Then cmbSelectValue.Items.Add(NoDupes(i))
        Next

        If cmbSelectValue.Items.Count <> 0 Then cmbSelectValue.SelectedIndex = cmbSelectValue.Items.Count - 1
    End Sub

    Private Sub rdbFromCol_Click(sender As Object, e As EventArgs) Handles rdbFromCol.Click
        Call LoadData()
    End Sub

    Private Sub rdbFromRow_Click(sender As Object, e As EventArgs) Handles rdbFromRow.Click
        Call LoadData()
    End Sub
End Class