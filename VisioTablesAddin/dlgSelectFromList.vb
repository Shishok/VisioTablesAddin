Public Class dlgSelectFromList

    Dim sh As Visio.Selection

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Dim i As Integer

        Call RecUndo("Вставить из списка")

        Select Case ckbOnlyActiveCell.Checked
            Case True ' Вставка выбранного значения в выделенные ячейки
                If InStr(1, winObj.Selection(1).NameU, "ClW", 1) <> 0 Then winObj.Selection(1).Characters.Text = cmbSelectValue.Text
            Case False ' Вставка выбранного значения в активную ячейку
                For i = 1 To sh.Count ' Вставка выбранного значения в выделенные ячейки
                    If InStr(1, sh(i).NameU, "ClW", 1) <> 0 Then sh(i).Characters.Text = cmbSelectValue.Text
                Next
        End Select

        Call RecUndo("0")

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    'Private Sub dlgSelectFromList_Activated(sender As Object, e As EventArgs) Handles Me.Activated
    '    cmbSelectValue.DroppedDown = True
    'End Sub

    Private Sub dlgSelectFromList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer, NoDupes As New Collection, Shp As Visio.Shape
        'Dim j As Integer, Swap1 As String, Swap2 As String

        sh = winObj.Selection
        Call SelCell(4)

        On Error Resume Next
        For Each Shp In winObj.Selection ' Заполнение коллекции значениями
            NoDupes.Add(Shp.Characters.Text, CStr(Shp.Characters.Text))
        Next
        On Error GoTo 0

        'For i = 1 To NoDupes.Count - 1 ' Сортировка коллекции
        '    For j = i + 1 To NoDupes.Count
        '        If NoDupes(i) > NoDupes(j) Then
        '            Swap1 = NoDupes(i)
        '            Swap2 = NoDupes(j)
        '            NoDupes.Add(Swap1, Before:=j)
        '            NoDupes.Add(Swap2, Before:=i)
        '            NoDupes.Remove(i + 1)
        '            NoDupes.Remove(j + 1)
        '        End If
        '    Next j
        'Next i

        For i = 1 To NoDupes.Count ' Добавление коллекции в ComboBox (cmbSelectValue)
            If StrComp(NoDupes(i), Trim("")) <> 0 Then cmbSelectValue.Items.Add(NoDupes(i))
        Next
        'cmbSelectValue.ListIndex = 0

        winObj.DeselectAll() : winObj.Selection = sh ' cmbSelectValue.DroppedDown = True
        If cmbSelectValue.Items.Count <> 0 Then cmbSelectValue.SelectedIndex = cmbSelectValue.Items.Count - 1
    End Sub

End Class