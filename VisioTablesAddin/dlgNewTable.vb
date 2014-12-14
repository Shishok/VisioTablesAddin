Imports System.Windows.Forms

Public Class dlgNewTable
    '    Private ctl As System.Windows.Forms.Control
    '    Private vsoApp As Visio.Application
    '    Private RadioButtonIsPress As String


    '    Private Sub dlgNewTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '        vsoApp = Globals.ThisAddIn.Application
    '        Const FRM = "###0.0###"

    '        With Me
    '            .txtTableCusHeight.Text = vsoApp.FormatResult(200, "mm", "mm", FRM)
    '            .txtTableCusWidth.Text = vsoApp.FormatResult(200, "mm", "mm", FRM)
    '            .txtCellDefHeight.Text = vsoApp.FormatResult(10, "mm", "mm", FRM)
    '            .txtCellDefWidth.Text = vsoApp.FormatResult(20, "mm", "mm", FRM)

    '            For Each .ctl In .Controls
    '                If .ctl.Name Like "*Dim*" Then
    '                    .ctl.Text = "mm"
    '                End If
    '            Next


    '            For Each Me.ctl In .Controls
    '                If .ctl.Tag = "1" Then .ctl.Enabled = False
    '                If .ctl.Tag = "0" Then .ctl.Enabled = True
    '            Next
    '        End With

    '        Call ToolTipfrm()
    '        RadioButtonIsPress = "optDefault"
    '        Call DelShape()
    '    End Sub

    '    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    '        With Me
    '            .DialogResult = System.Windows.Forms.DialogResult.OK
    '            'Call MainModule.GreatTable(RadioButtonIsPress, .nudColumns.Value, .nudRows.Value, .txtCellDefWidth.Text, _
    '            '.txtCellDefHeight.Text, .txtTableCusWidth.Text, .txtTableCusHeight.Text, .ckbDelShape.Checked)
    '        End With

    '        ' Me.nudColumns.Value
    '        ' Me.nudRows.Value
    '        ' Me.txtCellDefWidth.text
    '        ' Me.txtCellDefHeight.text
    '        ' Me.txtTableCusWidth.text
    '        ' Me.txtTableCusHeight.text
    '        ' Me.ckbDelShape.Checked
    '        Me.close()
    '    End Sub

    '    Private Sub optInside_CheckedChanged(sender As Object, e As EventArgs) Handles optInside.CheckedChanged
    '        With Me
    '            For Each .ctl In .Controls
    '                If .ctl.Tag = "1" Then .ctl.Enabled = False
    '                If .ctl.Tag = "0" Then .ctl.Enabled = False
    '            Next
    '        End With
    '        RadioButtonIsPress = "optInside"
    '        Call DelShape()
    '    End Sub

    '    Private Sub optDefault_CheckedChanged(sender As Object, e As EventArgs) Handles optDefault.CheckedChanged
    '        With Me
    '            For Each .ctl In .Controls
    '                If .ctl.Tag = "1" Then .ctl.Enabled = False
    '                If .ctl.Tag = "0" Then .ctl.Enabled = True
    '            Next
    '        End With
    '        RadioButtonIsPress = "optDefault"
    '        Call DelShape()
    '    End Sub

    '    Private Sub optPage_CheckedChanged(sender As Object, e As EventArgs) Handles optPage.CheckedChanged
    '        With Me
    '            For Each .ctl In .Controls
    '                If .ctl.Tag = "1" Then .ctl.Enabled = False
    '                If .ctl.Tag = "0" Then .ctl.Enabled = False
    '            Next
    '        End With
    '        RadioButtonIsPress = "optPage"
    '        Call DelShape()
    '    End Sub

    '    Private Sub optCustom_CheckedChanged(sender As Object, e As EventArgs) Handles optCustom.CheckedChanged
    '        With Me
    '            For Each .ctl In .Controls
    '                If .ctl.Tag = "1" Then .ctl.Enabled = True
    '                If .ctl.Tag = "0" Then .ctl.Enabled = False
    '            Next
    '        End With
    '        RadioButtonIsPress = "optCustom"
    '        Call DelShape()
    '    End Sub

    '    Private Sub DelShape()
    '        With Me
    '            .ckbDelShape.Enabled = .optInside.Checked
    '            .ckbDelShape.Checked = .optInside.Checked
    '        End With
    '    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    '    Private Sub ToolTipfrm()
    '        With Me
    '            .ToolTip1.SetToolTip(.optDefault, "Новая таблица по умолчанию")
    '            .ToolTip1.SetToolTip(.optCustom, "Новая таблица по вашим размерам")
    '            .ToolTip1.SetToolTip(.optPage, "Новая таблица по размерам листа, исключая поля листа")
    '            .ToolTip1.SetToolTip(.optInside, "Новая таблица по размерам выделенной фигуры")
    '            .ToolTip1.SetToolTip(.ckbDelShape, "Удалить фигуру-контур")
    '            .ToolTip1.SetToolTip(.nudColumns, "Количество столбцов в новой таблице")
    '            .ToolTip1.SetToolTip(.nudRows, "Количество строк в новой таблице")
    '            .ToolTip1.SetToolTip(.txtCellDefWidth, "Ширина ячеек в новой таблице")
    '            .ToolTip1.SetToolTip(.txtCellDefHeight, "Высота ячеек в новой таблице")
    '            .ToolTip1.SetToolTip(.txtTableCusWidth, "Ширина новой таблицы")
    '            .ToolTip1.SetToolTip(.txtTableCusHeight, "Высота новой таблицы")
    '        End With
    '    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        MessageBox.Show("Создание новой таблицы")
        Me.Close()
    End Sub

    Protected Overrides Sub OnClosed(ByVal e As EventArgs)
        booOpenForm = False
        MyBase.OnClosed(e)
    End Sub
End Class
