Imports System.Windows.Forms

Public Class dlgNewTable
    Dim ctl As System.Windows.Forms.Control
    Dim bytInsertType As Byte, booDeleteTargetShape As Boolean

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        strNameTable = txtNameTable.Text
        If txtCellDefWidth.Text = Trim("") Then txtCellDefWidth.Text = 0
        If txtCellDefHeight.Text = Trim("") Then txtCellDefHeight.Text = 0
        If txtTableCusWidth.Text = Trim("") Then txtTableCusWidth.Text = 0
        If txtTableCusHeight.Text = Trim("") Then txtTableCusHeight.Text = 0
        Me.Hide()
        Call CreatTable(strNameTable, bytInsertType, nudColumns.Value, nudRows.Value, txtCellDefWidth.Text, _
        txtCellDefHeight.Text, txtTableCusWidth.Text, txtTableCusHeight.Text, ckbDelShape.Checked, True)
        SaveSettings(1)
        Me.Close()
    End Sub

    Private Sub dlgNewTable_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmNewTable = Nothing
    End Sub

    Private Sub dlgNewTable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vsoApp = Globals.ThisAddIn.Application
        'booOpenForm = True
        Const FRM = "###0.0######"
        Dim strDrawingScale As String

        strDrawingScale = vsoApp.ActivePage.PageSheet.Cells("DrawingScale").ResultStrU("")
        strDrawingScale = Strings.Right(strDrawingScale, Strings.Len(strDrawingScale) - Strings.InStr(1, strDrawingScale, " ", 1))

        With Me
            .txtTableCusHeight.Text = vsoApp.FormatResult(200, "mm", strDrawingScale, FRM)
            .txtTableCusWidth.Text = vsoApp.FormatResult(200, "mm", strDrawingScale, FRM)
            .txtCellDefHeight.Text = vsoApp.FormatResult(10, "mm", strDrawingScale, FRM)
            .txtCellDefWidth.Text = vsoApp.FormatResult(20, "mm", strDrawingScale, FRM)

            For Each .ctl In .Controls
                If .ctl.Name Like "*Dim*" Then
                    .ctl.Text = strDrawingScale
                End If
            Next


            For Each Me.ctl In .Controls
                If .ctl.Tag = "1" Then .ctl.Enabled = False
                If .ctl.Tag = "0" Then .ctl.Enabled = True
            Next
        End With

        Call ToolTipfrm()
        bytInsertType = 1
        SaveSettings(0)
        Call DelShape()
    End Sub

    Private Sub optInside_CheckedChanged(sender As Object, e As EventArgs) Handles optInside.CheckedChanged
        With Me
            For Each .ctl In .Controls
                If .ctl.Tag = "1" Then .ctl.Enabled = False
                If .ctl.Tag = "0" Then .ctl.Enabled = False
            Next
        End With

        Call DelShape()
        bytInsertType = 4
        booDeleteTargetShape = ckbDelShape.Checked
    End Sub

    Private Sub optDefault_CheckedChanged(sender As Object, e As EventArgs) Handles optDefault.CheckedChanged
        With Me
            For Each .ctl In .Controls
                If .ctl.Tag = "1" Then .ctl.Enabled = False
                If .ctl.Tag = "0" Then .ctl.Enabled = True
            Next
        End With
        Call DelShape()
        bytInsertType = 1
    End Sub

    Private Sub optPage_CheckedChanged(sender As Object, e As EventArgs) Handles optPage.CheckedChanged
        With Me
            For Each .ctl In .Controls
                If .ctl.Tag = "1" Then .ctl.Enabled = False
                If .ctl.Tag = "0" Then .ctl.Enabled = False
            Next
        End With
        Call DelShape()
        bytInsertType = 2
    End Sub

    Private Sub optCustom_CheckedChanged(sender As Object, e As EventArgs) Handles optCustom.CheckedChanged
        With Me
            For Each .ctl In .Controls
                If .ctl.Tag = "1" Then .ctl.Enabled = True
                If .ctl.Tag = "0" Then .ctl.Enabled = False
            Next
        End With
        Call DelShape()
        bytInsertType = 3
    End Sub

    Private Sub DelShape()
        With Me
            .ckbDelShape.Enabled = .optInside.Checked
            '.ckbDelShape.Checked = .optInside.Checked
        End With
    End Sub

    Private Sub ToolTipfrm()
        With Me
            .ToolTip1.SetToolTip(.txtNameTable, "Имя Главной Управляющей ячейки, для идентификации ячеек таблицы")
            .ToolTip1.SetToolTip(.optDefault, "Новая таблица по умолчанию")
            .ToolTip1.SetToolTip(.optCustom, "Новая таблица по вашим размерам")
            .ToolTip1.SetToolTip(.optPage, "Новая таблица по размерам листа, исключая поля листа")
            .ToolTip1.SetToolTip(.optInside, "Новая таблица по размерам выделенной фигуры")
            .ToolTip1.SetToolTip(.ckbDelShape, "Удалить фигуру-контур")
            .ToolTip1.SetToolTip(.nudColumns, "Количество столбцов в новой таблице")
            .ToolTip1.SetToolTip(.nudRows, "Количество строк в новой таблице")
            .ToolTip1.SetToolTip(.txtCellDefWidth, "Ширина ячеек в новой таблице")
            .ToolTip1.SetToolTip(.txtCellDefHeight, "Высота ячеек в новой таблице")
            .ToolTip1.SetToolTip(.txtTableCusWidth, "Ширина новой таблицы")
            .ToolTip1.SetToolTip(.txtTableCusHeight, "Высота новой таблицы")
        End With
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub SaveSettings(arg)
        Const An As String = "VisioTableAddin", Sc As String = "NewTableWindow"
        Select Case arg
            Case 0
                If GetSetting(AppName:=An, Section:=Sc, Key:="Top") <> "" Then
                    Me.Top = GetSetting(AppName:=An, Section:=Sc, Key:="Top")
                    Me.Left = GetSetting(AppName:=An, Section:=Sc, Key:="Left")
                    txtNameTable.Text = GetSetting(AppName:=An, Section:=Sc, Key:="Name")
                    nudColumns.Value = GetSetting(AppName:=An, Section:=Sc, Key:="Columns")
                    nudRows.Value = GetSetting(AppName:=An, Section:=Sc, Key:="Rows")
                    txtCellDefWidth.Text = GetSetting(AppName:=An, Section:=Sc, Key:="WidthCell")
                    txtCellDefHeight.Text = GetSetting(AppName:=An, Section:=Sc, Key:="HeightCell")
                    txtTableCusWidth.Text = GetSetting(AppName:=An, Section:=Sc, Key:="WidthTable")
                    txtTableCusHeight.Text = GetSetting(AppName:=An, Section:=Sc, Key:="HeightTable")
                    ckbDelShape.Checked = GetSetting(AppName:=An, Section:=Sc, Key:="DeleteShape")
                    bytInsertType = GetSetting(AppName:=An, Section:=Sc, Key:="Method")
                    Select Case bytInsertType
                        Case 1 : optDefault.Checked = True
                        Case 2 : optPage.Checked = True
                        Case 3 : optCustom.Checked = True
                        Case 4 : optInside.Checked = True
                    End Select
                End If
            Case 1
                SaveSetting(AppName:=An, Section:=Sc, Key:="Top", Setting:=Me.Top)
                SaveSetting(AppName:=An, Section:=Sc, Key:="Left", Setting:=Me.Left)
                SaveSetting(AppName:=An, Section:=Sc, Key:="Name", Setting:=txtNameTable.Text)
                SaveSetting(AppName:=An, Section:=Sc, Key:="Columns", Setting:=nudColumns.Value)
                SaveSetting(AppName:=An, Section:=Sc, Key:="Rows", Setting:=nudRows.Value)
                SaveSetting(AppName:=An, Section:=Sc, Key:="WidthCell", Setting:=txtCellDefWidth.Text)
                SaveSetting(AppName:=An, Section:=Sc, Key:="HeightCell", Setting:=txtCellDefHeight.Text)
                SaveSetting(AppName:=An, Section:=Sc, Key:="WidthTable", Setting:=txtTableCusWidth.Text)
                SaveSetting(AppName:=An, Section:=Sc, Key:="HeightTable", Setting:=txtTableCusHeight.Text)
                SaveSetting(AppName:=An, Section:=Sc, Key:="DeleteShape", Setting:=ckbDelShape.Checked)
                SaveSetting(AppName:=An, Section:=Sc, Key:="Method", Setting:=bytInsertType)
        End Select
    End Sub

End Class
