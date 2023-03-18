<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgTableSize
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.ckbToHeightPage = New System.Windows.Forms.CheckBox()
        Me.lblTableHDim = New System.Windows.Forms.Label()
        Me.txtTableH = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ckbToWidthPage = New System.Windows.Forms.CheckBox()
        Me.lblTableLDim = New System.Windows.Forms.Label()
        Me.txtTableL = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblCellHDim = New System.Windows.Forms.Label()
        Me.txtCellH = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ckbWithActivateCell = New System.Windows.Forms.CheckBox()
        Me.lblCellLDim = New System.Windows.Forms.Label()
        Me.txtCellL = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.optAllHeightLock = New System.Windows.Forms.RadioButton()
        Me.optAllHeightAuto = New System.Windows.Forms.RadioButton()
        Me.ckbOnlySelectH = New System.Windows.Forms.CheckBox()
        Me.ckbAllHeight = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optAllWidthLock = New System.Windows.Forms.RadioButton()
        Me.optAllWidthAuto = New System.Windows.Forms.RadioButton()
        Me.ckbOnlySelectW = New System.Windows.Forms.CheckBox()
        Me.ckbAllWidth = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(313, 255)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Transparent
        Me.TabPage1.Controls.Add(Me.ckbToHeightPage)
        Me.TabPage1.Controls.Add(Me.lblTableHDim)
        Me.TabPage1.Controls.Add(Me.txtTableH)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.ckbToWidthPage)
        Me.TabPage1.Controls.Add(Me.lblTableLDim)
        Me.TabPage1.Controls.Add(Me.txtTableL)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.lblCellHDim)
        Me.TabPage1.Controls.Add(Me.txtCellH)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.ckbWithActivateCell)
        Me.TabPage1.Controls.Add(Me.lblCellLDim)
        Me.TabPage1.Controls.Add(Me.txtCellL)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(305, 229)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Размеры"
        '
        'ckbToHeightPage
        '
        Me.ckbToHeightPage.AutoSize = True
        Me.ckbToHeightPage.Location = New System.Drawing.Point(13, 194)
        Me.ckbToHeightPage.Name = "ckbToHeightPage"
        Me.ckbToHeightPage.Size = New System.Drawing.Size(156, 17)
        Me.ckbToHeightPage.TabIndex = 16
        Me.ckbToHeightPage.Text = "Таблица по высоте листа"
        Me.ckbToHeightPage.UseVisualStyleBackColor = True
        '
        'lblTableHDim
        '
        Me.lblTableHDim.Location = New System.Drawing.Point(229, 141)
        Me.lblTableHDim.Name = "lblTableHDim"
        Me.lblTableHDim.Size = New System.Drawing.Size(40, 13)
        Me.lblTableHDim.TabIndex = 15
        Me.lblTableHDim.Text = "мм"
        '
        'txtTableH
        '
        Me.txtTableH.Location = New System.Drawing.Point(118, 138)
        Me.txtTableH.Name = "txtTableH"
        Me.txtTableH.Size = New System.Drawing.Size(100, 20)
        Me.txtTableH.TabIndex = 14
        Me.txtTableH.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(10, 141)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Высота таблицы"
        '
        'ckbToWidthPage
        '
        Me.ckbToWidthPage.AutoSize = True
        Me.ckbToWidthPage.Location = New System.Drawing.Point(13, 171)
        Me.ckbToWidthPage.Name = "ckbToWidthPage"
        Me.ckbToWidthPage.Size = New System.Drawing.Size(157, 17)
        Me.ckbToWidthPage.TabIndex = 12
        Me.ckbToWidthPage.Text = "Таблица по ширине листа"
        Me.ckbToWidthPage.UseVisualStyleBackColor = True
        '
        'lblTableLDim
        '
        Me.lblTableLDim.Location = New System.Drawing.Point(229, 114)
        Me.lblTableLDim.Name = "lblTableLDim"
        Me.lblTableLDim.Size = New System.Drawing.Size(40, 13)
        Me.lblTableLDim.TabIndex = 10
        Me.lblTableLDim.Text = "мм"
        '
        'txtTableL
        '
        Me.txtTableL.Location = New System.Drawing.Point(118, 111)
        Me.txtTableL.Name = "txtTableL"
        Me.txtTableL.Size = New System.Drawing.Size(100, 20)
        Me.txtTableL.TabIndex = 9
        Me.txtTableL.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(10, 114)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(92, 13)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Ширина таблицы"
        '
        'lblCellHDim
        '
        Me.lblCellHDim.Location = New System.Drawing.Point(229, 43)
        Me.lblCellHDim.Name = "lblCellHDim"
        Me.lblCellHDim.Size = New System.Drawing.Size(40, 13)
        Me.lblCellHDim.TabIndex = 7
        Me.lblCellHDim.Text = "мм"
        '
        'txtCellH
        '
        Me.txtCellH.Location = New System.Drawing.Point(118, 40)
        Me.txtCellH.Name = "txtCellH"
        Me.txtCellH.Size = New System.Drawing.Size(100, 20)
        Me.txtCellH.TabIndex = 6
        Me.txtCellH.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Высота строки"
        '
        'ckbWithActivateCell
        '
        Me.ckbWithActivateCell.AutoSize = True
        Me.ckbWithActivateCell.Checked = True
        Me.ckbWithActivateCell.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbWithActivateCell.Location = New System.Drawing.Point(13, 73)
        Me.ckbWithActivateCell.Name = "ckbWithActivateCell"
        Me.ckbWithActivateCell.Size = New System.Drawing.Size(216, 17)
        Me.ckbWithActivateCell.TabIndex = 4
        Me.ckbWithActivateCell.Text = "Только выделенные строки/столбцы"
        Me.ckbWithActivateCell.UseVisualStyleBackColor = True
        '
        'lblCellLDim
        '
        Me.lblCellLDim.Location = New System.Drawing.Point(229, 16)
        Me.lblCellLDim.Name = "lblCellLDim"
        Me.lblCellLDim.Size = New System.Drawing.Size(40, 13)
        Me.lblCellLDim.TabIndex = 2
        Me.lblCellLDim.Text = "мм"
        '
        'txtCellL
        '
        Me.txtCellL.Location = New System.Drawing.Point(118, 13)
        Me.txtCellL.Name = "txtCellL"
        Me.txtCellL.Size = New System.Drawing.Size(100, 20)
        Me.txtCellL.TabIndex = 1
        Me.txtCellL.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Ширина столбца"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Transparent
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(305, 229)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Авторазмеры"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.optAllHeightLock)
        Me.GroupBox2.Controls.Add(Me.optAllHeightAuto)
        Me.GroupBox2.Controls.Add(Me.ckbOnlySelectH)
        Me.GroupBox2.Controls.Add(Me.ckbAllHeight)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 115)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(293, 100)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = " Высота "
        '
        'optAllHeightLock
        '
        Me.optAllHeightLock.AutoSize = True
        Me.optAllHeightLock.Location = New System.Drawing.Point(80, 42)
        Me.optAllHeightLock.Name = "optAllHeightLock"
        Me.optAllHeightLock.Size = New System.Drawing.Size(103, 17)
        Me.optAllHeightLock.TabIndex = 3
        Me.optAllHeightLock.TabStop = True
        Me.optAllHeightLock.Text = "Заблокировать"
        Me.optAllHeightLock.UseVisualStyleBackColor = True
        '
        'optAllHeightAuto
        '
        Me.optAllHeightAuto.AutoSize = True
        Me.optAllHeightAuto.Location = New System.Drawing.Point(10, 42)
        Me.optAllHeightAuto.Name = "optAllHeightAuto"
        Me.optAllHeightAuto.Size = New System.Drawing.Size(49, 17)
        Me.optAllHeightAuto.TabIndex = 2
        Me.optAllHeightAuto.TabStop = True
        Me.optAllHeightAuto.Text = "Авто"
        Me.optAllHeightAuto.UseVisualStyleBackColor = True
        '
        'ckbOnlySelectH
        '
        Me.ckbOnlySelectH.AutoSize = True
        Me.ckbOnlySelectH.Location = New System.Drawing.Point(10, 70)
        Me.ckbOnlySelectH.Name = "ckbOnlySelectH"
        Me.ckbOnlySelectH.Size = New System.Drawing.Size(130, 17)
        Me.ckbOnlySelectH.TabIndex = 1
        Me.ckbOnlySelectH.Text = "Только выделенные"
        Me.ckbOnlySelectH.UseVisualStyleBackColor = True
        '
        'ckbAllHeight
        '
        Me.ckbAllHeight.AutoSize = True
        Me.ckbAllHeight.Checked = True
        Me.ckbAllHeight.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbAllHeight.Location = New System.Drawing.Point(10, 19)
        Me.ckbAllHeight.Name = "ckbAllHeight"
        Me.ckbAllHeight.Size = New System.Drawing.Size(175, 17)
        Me.ckbAllHeight.TabIndex = 0
        Me.ckbAllHeight.Text = "Все строки по высоте текста"
        Me.ckbAllHeight.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optAllWidthLock)
        Me.GroupBox1.Controls.Add(Me.optAllWidthAuto)
        Me.GroupBox1.Controls.Add(Me.ckbOnlySelectW)
        Me.GroupBox1.Controls.Add(Me.ckbAllWidth)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(293, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = " Ширина "
        '
        'optAllWidthLock
        '
        Me.optAllWidthLock.AutoSize = True
        Me.optAllWidthLock.Location = New System.Drawing.Point(80, 42)
        Me.optAllWidthLock.Name = "optAllWidthLock"
        Me.optAllWidthLock.Size = New System.Drawing.Size(103, 17)
        Me.optAllWidthLock.TabIndex = 3
        Me.optAllWidthLock.TabStop = True
        Me.optAllWidthLock.Text = "Заблокировать"
        Me.optAllWidthLock.UseVisualStyleBackColor = True
        '
        'optAllWidthAuto
        '
        Me.optAllWidthAuto.AutoSize = True
        Me.optAllWidthAuto.Location = New System.Drawing.Point(10, 42)
        Me.optAllWidthAuto.Name = "optAllWidthAuto"
        Me.optAllWidthAuto.Size = New System.Drawing.Size(49, 17)
        Me.optAllWidthAuto.TabIndex = 2
        Me.optAllWidthAuto.TabStop = True
        Me.optAllWidthAuto.Text = "Авто"
        Me.optAllWidthAuto.UseVisualStyleBackColor = True
        '
        'ckbOnlySelectW
        '
        Me.ckbOnlySelectW.AutoSize = True
        Me.ckbOnlySelectW.Location = New System.Drawing.Point(10, 70)
        Me.ckbOnlySelectW.Name = "ckbOnlySelectW"
        Me.ckbOnlySelectW.Size = New System.Drawing.Size(130, 17)
        Me.ckbOnlySelectW.TabIndex = 1
        Me.ckbOnlySelectW.Text = "Только выделенные"
        Me.ckbOnlySelectW.UseVisualStyleBackColor = True
        '
        'ckbAllWidth
        '
        Me.ckbAllWidth.AutoSize = True
        Me.ckbAllWidth.Checked = True
        Me.ckbAllWidth.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbAllWidth.Location = New System.Drawing.Point(10, 19)
        Me.ckbAllWidth.Name = "ckbAllWidth"
        Me.ckbAllWidth.Size = New System.Drawing.Size(184, 17)
        Me.ckbAllWidth.TabIndex = 0
        Me.ckbAllWidth.Text = "Все столбцы по ширине текста"
        Me.ckbAllWidth.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(175, 277)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 12
        Me.OK_Button.Text = "ОК"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 13
        Me.Cancel_Button.Text = "Отмена"
        '
        'dlgTableSize
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(337, 317)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.TabControl1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(353, 356)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(353, 356)
        Me.Name = "dlgTableSize"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Настройка размеров"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lblCellHDim As System.Windows.Forms.Label
    Friend WithEvents txtCellH As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ckbWithActivateCell As System.Windows.Forms.CheckBox
    Friend WithEvents lblCellLDim As System.Windows.Forms.Label
    Friend WithEvents txtCellL As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ckbToHeightPage As System.Windows.Forms.CheckBox
    Friend WithEvents lblTableHDim As System.Windows.Forms.Label
    Friend WithEvents txtTableH As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ckbToWidthPage As System.Windows.Forms.CheckBox
    Friend WithEvents lblTableLDim As System.Windows.Forms.Label
    Friend WithEvents txtTableL As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ckbAllWidth As System.Windows.Forms.CheckBox
    Friend WithEvents optAllWidthLock As System.Windows.Forms.RadioButton
    Friend WithEvents optAllWidthAuto As System.Windows.Forms.RadioButton
    Friend WithEvents ckbOnlySelectW As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents optAllHeightLock As System.Windows.Forms.RadioButton
    Friend WithEvents optAllHeightAuto As System.Windows.Forms.RadioButton
    Friend WithEvents ckbOnlySelectH As System.Windows.Forms.CheckBox
    Friend WithEvents ckbAllHeight As System.Windows.Forms.CheckBox
End Class
