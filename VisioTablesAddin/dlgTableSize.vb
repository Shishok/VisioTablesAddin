Public Class dlgTableSize

    Dim m_bytCellsOrTable As Byte, m_booWidth As Boolean, m_booHeight As Boolean

    Private Sub dlgTableSize_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strDrawingScale As String
        Me.TabControl1.SelectedIndex = FlagPage
        If FlagPage = 0 Then Call InitSize()

        strDrawingScale = winObj.Page.PageSheet.Cells("DrawingScale").ResultStrU("")
        strDrawingScale = Strings.Right(strDrawingScale, Len(strDrawingScale) - InStr(1, strDrawingScale, " ", 1))

        lblCellLDim.Text = strDrawingScale : lblCellHDim.Text = strDrawingScale
        lblTableLDim.Text = strDrawingScale : lblTableHDim.Text = strDrawingScale
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        Select Case TabControl1.SelectedIndex

            Case 0 ' Размеры
                If Val(txtCellL.Text) = 0 Or Val(txtTableL.Text) = 0 Or Val(txtCellH.Text) = 0 Or Val(txtTableH.Text) = 0 Then Exit Sub
                Dim TableW As Single, TableH As Single
                TableW = txtTableL.Text : TableH = txtTableH.Text
                With winObj.Page.PageSheet
                    If ckbToWidthPage.Checked Then TableW = .Cells("PageWidth").Result(64) - .Cells("PageRightMargin").Result(64) - .Cells("PageLeftMargin").Result(64)
                    If ckbToHeightPage.Checked Then TableH = .Cells("PageHeight").Result(64) - .Cells("PageTopMargin").Result(64) - .Cells("PageBottomMargin").Result(64)
                End With
                If txtCellL.Tag <> txtCellL.Text Or txtTableL.Tag <> txtTableL.Text Or ckbToWidthPage.Checked Then m_booWidth = True
                If txtCellH.Tag <> txtCellH.Text Or txtTableH.Tag <> txtTableH.Text Or ckbToHeightPage.Checked Then m_booHeight = True
                Call ResizeCells(m_bytCellsOrTable, ckbWithActivateCell.Checked, txtCellL.Text, txtCellH.Text, TableW, TableH, m_booWidth, m_booHeight)

            Case 1 ' Авторазмеры
                Dim bytNothingOrAutoOrLockColumns As Byte, bytNothingOrAutoOrLockRows As Byte
                If ckbAllWidth.Checked + ckbAllHeight.Checked = 0 Then Exit Sub
                bytNothingOrAutoOrLockColumns = 0 ' 0-nothing, 1-auto, 2-lock
                bytNothingOrAutoOrLockRows = 0
                If optAllWidthAuto.Checked Then bytNothingOrAutoOrLockColumns = 1
                If optAllWidthLock.Checked Then bytNothingOrAutoOrLockColumns = 2
                If optAllHeightAuto.Checked Then bytNothingOrAutoOrLockRows = 1
                If optAllHeightLock.Checked Then bytNothingOrAutoOrLockRows = 2
                Call AllAlignOnText(ckbAllWidth.Checked, ckbAllHeight.Checked, bytNothingOrAutoOrLockColumns, bytNothingOrAutoOrLockRows, ckbOnlySelectW.Checked, ckbOnlySelectH.Checked)

        End Select

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.Close()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then Call InitSize()
    End Sub

    Private Sub InitSize()
        Const frm = "####0.####"
        If TabControl1.SelectedIndex = 1 Then m_bytCellsOrTable = 3
        If TabControl1.SelectedIndex = 0 AndAlso winObj.Selection.Count > 0 Then
            txtTableL.Text = Format(fSTWH(winObj.Selection(1), 1, True), frm)
            txtTableH.Text = Format(fSTWH(winObj.Selection(1), 2, False), frm)
            txtCellL.Text = Format(winObj.Selection(1).Cells("Width").Result(64), frm)
            txtCellH.Text = Format(winObj.Selection(1).Cells("Height").Result(64), frm)
            txtTableL.Tag = txtTableL.Text
            txtTableH.Tag = txtTableH.Text
            txtCellL.Tag = txtCellL.Text
            txtCellH.Tag = txtCellH.Text
        End If
    End Sub

#Region "Tab-0"

    Private Sub txtCellL_MouseDown(sender As Object, e As Windows.Forms.MouseEventArgs) Handles txtCellL.MouseDown
        m_bytCellsOrTable = 1
    End Sub

    Private Sub txtCellH_MouseDown(sender As Object, e As Windows.Forms.MouseEventArgs) Handles txtCellH.MouseDown
        m_bytCellsOrTable = 1
    End Sub

    Private Sub txtTableL_MouseDown(sender As Object, e As Windows.Forms.MouseEventArgs) Handles txtTableL.MouseDown
        m_bytCellsOrTable = 2
    End Sub

    Private Sub txtTableH_MouseDown(sender As Object, e As Windows.Forms.MouseEventArgs) Handles txtTableH.MouseDown
        m_bytCellsOrTable = 2
    End Sub

    Private Sub ckbToWidthPage_Click(sender As Object, e As EventArgs) Handles ckbToWidthPage.Click
        m_bytCellsOrTable = 2
    End Sub

    Private Sub ckbToHeightPage_Click(sender As Object, e As EventArgs) Handles ckbToHeightPage.Click
        m_bytCellsOrTable = 2
    End Sub

#End Region

#Region "Tab-1"

    Private Sub ckbAllWidth_CheckedChanged(sender As Object, e As EventArgs) Handles ckbAllWidth.CheckedChanged
        If ckbAllWidth.Checked Then
            optAllWidthAuto.Checked = False
            optAllWidthLock.Checked = False
        End If
        optAllWidthAuto.Enabled = ckbAllWidth.Checked
        optAllWidthLock.Enabled = ckbAllWidth.Checked
        ckbOnlySelectW.Enabled = ckbAllWidth.Checked
    End Sub

    Private Sub ckbAllHeight_CheckedChanged(sender As Object, e As EventArgs) Handles ckbAllHeight.CheckedChanged
        If ckbAllHeight.Checked Then
            optAllHeightAuto.Checked = False
            optAllHeightLock.Checked = False
        End If
        optAllHeightAuto.Enabled = ckbAllHeight.Checked
        optAllHeightLock.Enabled = ckbAllHeight.Checked
        ckbOnlySelectH.Enabled = ckbAllWidth.Checked
    End Sub

#End Region

End Class