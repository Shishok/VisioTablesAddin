Public Class dlgPictures

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Dim hAL As Byte = 1, Val As Byte = 2, shN As Boolean = False, lF As Boolean = False

        If optAlignCenterH.Checked Then hAL = 2 ' hAL - выравнивание по горизонтали(1-3)
        If optAlignRightH.Checked Then hAL = 3

        If optAlignTopV.Checked Then Val = 1 ' vAL - выравнивание по вертикали(1-3)
        If optAlignBottomV.Checked Then Val = 3

        If ckbInsertName.Checked Then shN = True ' shN - помещать названия в ячейку(0,1)

        If ckbLockFormulas.Checked Then lF = True ' lF - блокировать формулы(True,False)
        Call LockPicture(hAL, Val, shN, lF, True)

        Me.Close()
    End Sub

    Private Sub dlgPictures_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With Me.ToolTip1
            .SetToolTip(ckbInsertName, "Вставка имени фигуры в ту же ячейку таблицы ")
            .SetToolTip(ckbLockFormulas, "Блокировка фигур(формул PinX и PinY) в ячейках таблицы")
        End With
    End Sub
End Class