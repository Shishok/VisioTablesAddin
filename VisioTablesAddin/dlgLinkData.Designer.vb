<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgLinkData
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
        Me.lblCountRow = New System.Windows.Forms.Label()
        Me.lblSourseData = New System.Windows.Forms.Label()
        Me.cmbSourseData = New System.Windows.Forms.ComboBox()
        Me.txtNameTable = New System.Windows.Forms.TextBox()
        Me.ckbInsertName = New System.Windows.Forms.CheckBox()
        Me.ckbTitleColumns = New System.Windows.Forms.CheckBox()
        Me.ckbFontBold = New System.Windows.Forms.CheckBox()
        Me.ckbInvisibleZero = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtRowStartSourse = New System.Windows.Forms.NumericUpDown()
        Me.txtRowEndSourse = New System.Windows.Forms.NumericUpDown()
        Me.txtColStartSourse = New System.Windows.Forms.NumericUpDown()
        Me.txtColEndSourse = New System.Windows.Forms.NumericUpDown()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.cmdRefreshAll = New System.Windows.Forms.Button()
        CType(Me.txtRowStartSourse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRowEndSourse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtColStartSourse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtColEndSourse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblCountRow
        '
        Me.lblCountRow.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblCountRow.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.lblCountRow.Location = New System.Drawing.Point(12, 9)
        Me.lblCountRow.Name = "lblCountRow"
        Me.lblCountRow.Size = New System.Drawing.Size(513, 23)
        Me.lblCountRow.TabIndex = 0
        Me.lblCountRow.Text = "Источник"
        '
        'lblSourseData
        '
        Me.lblSourseData.AutoSize = True
        Me.lblSourseData.Location = New System.Drawing.Point(12, 31)
        Me.lblSourseData.Name = "lblSourseData"
        Me.lblSourseData.Size = New System.Drawing.Size(135, 13)
        Me.lblSourseData.TabIndex = 1
        Me.lblSourseData.Text = "Выбор источника данных"
        '
        'cmbSourseData
        '
        Me.cmbSourseData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbSourseData.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSourseData.FormattingEnabled = True
        Me.cmbSourseData.Location = New System.Drawing.Point(12, 47)
        Me.cmbSourseData.Name = "cmbSourseData"
        Me.cmbSourseData.Size = New System.Drawing.Size(513, 21)
        Me.cmbSourseData.TabIndex = 2
        '
        'txtNameTable
        '
        Me.txtNameTable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNameTable.Location = New System.Drawing.Point(12, 107)
        Me.txtNameTable.Name = "txtNameTable"
        Me.txtNameTable.Size = New System.Drawing.Size(513, 20)
        Me.txtNameTable.TabIndex = 3
        '
        'ckbInsertName
        '
        Me.ckbInsertName.AutoSize = True
        Me.ckbInsertName.Checked = True
        Me.ckbInsertName.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbInsertName.Location = New System.Drawing.Point(12, 84)
        Me.ckbInsertName.Name = "ckbInsertName"
        Me.ckbInsertName.Size = New System.Drawing.Size(259, 17)
        Me.ckbInsertName.TabIndex = 4
        Me.ckbInsertName.Text = "Вставить название таблицы в первую строку "
        Me.ckbInsertName.UseVisualStyleBackColor = True
        '
        'ckbTitleColumns
        '
        Me.ckbTitleColumns.AutoSize = True
        Me.ckbTitleColumns.Checked = True
        Me.ckbTitleColumns.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbTitleColumns.Location = New System.Drawing.Point(300, 84)
        Me.ckbTitleColumns.Name = "ckbTitleColumns"
        Me.ckbTitleColumns.Size = New System.Drawing.Size(130, 17)
        Me.ckbTitleColumns.TabIndex = 5
        Me.ckbTitleColumns.Text = "Заголовки столбцов"
        Me.ckbTitleColumns.UseVisualStyleBackColor = True
        '
        'ckbFontBold
        '
        Me.ckbFontBold.AutoSize = True
        Me.ckbFontBold.Location = New System.Drawing.Point(12, 142)
        Me.ckbFontBold.Name = "ckbFontBold"
        Me.ckbFontBold.Size = New System.Drawing.Size(227, 17)
        Me.ckbFontBold.TabIndex = 6
        Me.ckbFontBold.Text = "Выделять заголовки жирным шрифтом"
        Me.ckbFontBold.UseVisualStyleBackColor = True
        '
        'ckbInvisibleZero
        '
        Me.ckbInvisibleZero.AutoSize = True
        Me.ckbInvisibleZero.Location = New System.Drawing.Point(245, 142)
        Me.ckbInvisibleZero.Name = "ckbInvisibleZero"
        Me.ckbInvisibleZero.Size = New System.Drawing.Size(280, 17)
        Me.ckbInvisibleZero.TabIndex = 7
        Me.ckbInvisibleZero.Text = "Не показывать нулевые и пустые значения ячеек"
        Me.ckbInvisibleZero.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 177)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Начинать с"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 212)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Завершить на"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(190, 177)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "строки и"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(190, 212)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(51, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "строке и"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(333, 177)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(143, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "столбца источника данных"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(333, 212)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(143, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "столбце источника данных"
        '
        'txtRowStartSourse
        '
        Me.txtRowStartSourse.Location = New System.Drawing.Point(104, 173)
        Me.txtRowStartSourse.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.txtRowStartSourse.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtRowStartSourse.Name = "txtRowStartSourse"
        Me.txtRowStartSourse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRowStartSourse.Size = New System.Drawing.Size(80, 20)
        Me.txtRowStartSourse.TabIndex = 14
        Me.txtRowStartSourse.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtRowEndSourse
        '
        Me.txtRowEndSourse.Location = New System.Drawing.Point(104, 208)
        Me.txtRowEndSourse.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.txtRowEndSourse.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtRowEndSourse.Name = "txtRowEndSourse"
        Me.txtRowEndSourse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRowEndSourse.Size = New System.Drawing.Size(80, 20)
        Me.txtRowEndSourse.TabIndex = 15
        Me.txtRowEndSourse.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtColStartSourse
        '
        Me.txtColStartSourse.Location = New System.Drawing.Point(247, 173)
        Me.txtColStartSourse.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.txtColStartSourse.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtColStartSourse.Name = "txtColStartSourse"
        Me.txtColStartSourse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColStartSourse.Size = New System.Drawing.Size(80, 20)
        Me.txtColStartSourse.TabIndex = 16
        Me.txtColStartSourse.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtColEndSourse
        '
        Me.txtColEndSourse.Location = New System.Drawing.Point(247, 208)
        Me.txtColEndSourse.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.txtColEndSourse.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtColEndSourse.Name = "txtColEndSourse"
        Me.txtColEndSourse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtColEndSourse.Size = New System.Drawing.Size(80, 20)
        Me.txtColEndSourse.TabIndex = 17
        Me.txtColEndSourse.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(379, 254)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 18
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
        'cmdRefreshAll
        '
        Me.cmdRefreshAll.AutoSize = True
        Me.cmdRefreshAll.Location = New System.Drawing.Point(12, 257)
        Me.cmdRefreshAll.Name = "cmdRefreshAll"
        Me.cmdRefreshAll.Size = New System.Drawing.Size(182, 23)
        Me.cmdRefreshAll.TabIndex = 19
        Me.cmdRefreshAll.Text = "Обновить все источники данных"
        Me.cmdRefreshAll.UseVisualStyleBackColor = True
        '
        'dlgLinkData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 290)
        Me.Controls.Add(Me.cmdRefreshAll)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.txtColEndSourse)
        Me.Controls.Add(Me.txtColStartSourse)
        Me.Controls.Add(Me.txtRowEndSourse)
        Me.Controls.Add(Me.txtRowStartSourse)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ckbInvisibleZero)
        Me.Controls.Add(Me.ckbFontBold)
        Me.Controls.Add(Me.ckbTitleColumns)
        Me.Controls.Add(Me.ckbInsertName)
        Me.Controls.Add(Me.txtNameTable)
        Me.Controls.Add(Me.cmbSourseData)
        Me.Controls.Add(Me.lblSourseData)
        Me.Controls.Add(Me.lblCountRow)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(700, 328)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(551, 328)
        Me.Name = "dlgLinkData"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Связь с внешними данными"
        Me.TopMost = True
        CType(Me.txtRowStartSourse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRowEndSourse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtColStartSourse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtColEndSourse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblCountRow As System.Windows.Forms.Label
    Friend WithEvents lblSourseData As System.Windows.Forms.Label
    Friend WithEvents cmbSourseData As System.Windows.Forms.ComboBox
    Friend WithEvents txtNameTable As System.Windows.Forms.TextBox
    Friend WithEvents ckbInsertName As System.Windows.Forms.CheckBox
    Friend WithEvents ckbTitleColumns As System.Windows.Forms.CheckBox
    Friend WithEvents ckbFontBold As System.Windows.Forms.CheckBox
    Friend WithEvents ckbInvisibleZero As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtRowStartSourse As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtRowEndSourse As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtColStartSourse As System.Windows.Forms.NumericUpDown
    Friend WithEvents txtColEndSourse As System.Windows.Forms.NumericUpDown
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents cmdRefreshAll As System.Windows.Forms.Button
End Class
