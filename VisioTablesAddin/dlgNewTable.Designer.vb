<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgNewTable
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
        Me.components = New System.ComponentModel.Container()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.optDefault = New System.Windows.Forms.RadioButton()
        Me.optPage = New System.Windows.Forms.RadioButton()
        Me.optCustom = New System.Windows.Forms.RadioButton()
        Me.optInside = New System.Windows.Forms.RadioButton()
        Me.ckbDelShape = New System.Windows.Forms.CheckBox()
        Me.txtCellDefHeight = New System.Windows.Forms.TextBox()
        Me.txtCellDefWidth = New System.Windows.Forms.TextBox()
        Me.txtTableCusHeight = New System.Windows.Forms.TextBox()
        Me.txtTableCusWidth = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblTableLDim = New System.Windows.Forms.Label()
        Me.lblTableHDim = New System.Windows.Forms.Label()
        Me.lblCellLDim = New System.Windows.Forms.Label()
        Me.lblCellHDim = New System.Windows.Forms.Label()
        Me.nudColumns = New System.Windows.Forms.NumericUpDown()
        Me.nudRows = New System.Windows.Forms.NumericUpDown()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtNameTable = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.nudColumns, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudRows, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(221, 292)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
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
        'optDefault
        '
        Me.optDefault.AutoSize = True
        Me.optDefault.Checked = True
        Me.optDefault.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.optDefault.Location = New System.Drawing.Point(15, 75)
        Me.optDefault.Name = "optDefault"
        Me.optDefault.Size = New System.Drawing.Size(108, 18)
        Me.optDefault.TabIndex = 1
        Me.optDefault.TabStop = True
        Me.optDefault.Tag = ""
        Me.optDefault.Text = "По умолчанию"
        Me.optDefault.UseVisualStyleBackColor = True
        '
        'optPage
        '
        Me.optPage.AutoSize = True
        Me.optPage.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.optPage.Location = New System.Drawing.Point(15, 128)
        Me.optPage.Name = "optPage"
        Me.optPage.Size = New System.Drawing.Size(126, 18)
        Me.optPage.TabIndex = 2
        Me.optPage.Tag = ""
        Me.optPage.Text = "По размеру листа"
        Me.optPage.UseVisualStyleBackColor = True
        '
        'optCustom
        '
        Me.optCustom.AutoSize = True
        Me.optCustom.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.optCustom.Location = New System.Drawing.Point(15, 175)
        Me.optCustom.Name = "optCustom"
        Me.optCustom.Size = New System.Drawing.Size(104, 18)
        Me.optCustom.TabIndex = 3
        Me.optCustom.Tag = ""
        Me.optCustom.Text = "Свои размеры"
        Me.optCustom.UseVisualStyleBackColor = True
        '
        'optInside
        '
        Me.optInside.AutoSize = True
        Me.optInside.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.optInside.Location = New System.Drawing.Point(15, 223)
        Me.optInside.Name = "optInside"
        Me.optInside.Size = New System.Drawing.Size(112, 18)
        Me.optInside.TabIndex = 4
        Me.optInside.Tag = ""
        Me.optInside.Text = "Внутри фигуры"
        Me.optInside.UseVisualStyleBackColor = True
        '
        'ckbDelShape
        '
        Me.ckbDelShape.AutoSize = True
        Me.ckbDelShape.Checked = True
        Me.ckbDelShape.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ckbDelShape.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.ckbDelShape.Location = New System.Drawing.Point(15, 254)
        Me.ckbDelShape.Name = "ckbDelShape"
        Me.ckbDelShape.Size = New System.Drawing.Size(117, 18)
        Me.ckbDelShape.TabIndex = 5
        Me.ckbDelShape.Text = "Удалить фигуру"
        Me.ckbDelShape.UseVisualStyleBackColor = True
        '
        'txtCellDefHeight
        '
        Me.txtCellDefHeight.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtCellDefHeight.Location = New System.Drawing.Point(271, 176)
        Me.txtCellDefHeight.Name = "txtCellDefHeight"
        Me.txtCellDefHeight.Size = New System.Drawing.Size(60, 22)
        Me.txtCellDefHeight.TabIndex = 9
        Me.txtCellDefHeight.Tag = "0"
        Me.txtCellDefHeight.Text = "10"
        Me.txtCellDefHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCellDefWidth
        '
        Me.txtCellDefWidth.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtCellDefWidth.Location = New System.Drawing.Point(271, 148)
        Me.txtCellDefWidth.Name = "txtCellDefWidth"
        Me.txtCellDefWidth.Size = New System.Drawing.Size(60, 22)
        Me.txtCellDefWidth.TabIndex = 8
        Me.txtCellDefWidth.Tag = "0"
        Me.txtCellDefWidth.Text = "20"
        Me.txtCellDefWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTableCusHeight
        '
        Me.txtTableCusHeight.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtTableCusHeight.Location = New System.Drawing.Point(271, 253)
        Me.txtTableCusHeight.Name = "txtTableCusHeight"
        Me.txtTableCusHeight.Size = New System.Drawing.Size(60, 22)
        Me.txtTableCusHeight.TabIndex = 11
        Me.txtTableCusHeight.Tag = "1"
        Me.txtTableCusHeight.Text = "100"
        Me.txtTableCusHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTableCusWidth
        '
        Me.txtTableCusWidth.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtTableCusWidth.Location = New System.Drawing.Point(271, 225)
        Me.txtTableCusWidth.Name = "txtTableCusWidth"
        Me.txtTableCusWidth.Size = New System.Drawing.Size(60, 22)
        Me.txtTableCusWidth.TabIndex = 10
        Me.txtTableCusWidth.Tag = "1"
        Me.txtTableCusWidth.Text = "200"
        Me.txtTableCusWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(204, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 14)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Столбцов"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(225, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 14)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Строк"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(180, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 14)
        Me.Label3.TabIndex = 15
        Me.Label3.Tag = "0"
        Me.Label3.Text = "Высота ячеек"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.Location = New System.Drawing.Point(176, 151)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(89, 14)
        Me.Label4.TabIndex = 14
        Me.Label4.Tag = "0"
        Me.Label4.Text = "Ширина ячеек"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(166, 256)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(99, 14)
        Me.Label5.TabIndex = 17
        Me.Label5.Tag = "1"
        Me.Label5.Text = "Высота таблицы"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label6.Location = New System.Drawing.Point(162, 230)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(103, 14)
        Me.Label6.TabIndex = 16
        Me.Label6.Tag = "1"
        Me.Label6.Text = "Ширина таблицы"
        '
        'lblTableLDim
        '
        Me.lblTableLDim.AutoSize = True
        Me.lblTableLDim.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblTableLDim.Location = New System.Drawing.Point(340, 230)
        Me.lblTableLDim.Name = "lblTableLDim"
        Me.lblTableLDim.Size = New System.Drawing.Size(27, 14)
        Me.lblTableLDim.TabIndex = 18
        Me.lblTableLDim.Tag = "1"
        Me.lblTableLDim.Text = "mm"
        '
        'lblTableHDim
        '
        Me.lblTableHDim.AutoSize = True
        Me.lblTableHDim.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblTableHDim.Location = New System.Drawing.Point(340, 256)
        Me.lblTableHDim.Name = "lblTableHDim"
        Me.lblTableHDim.Size = New System.Drawing.Size(27, 14)
        Me.lblTableHDim.TabIndex = 19
        Me.lblTableHDim.Tag = "1"
        Me.lblTableHDim.Text = "mm"
        '
        'lblCellLDim
        '
        Me.lblCellLDim.AutoSize = True
        Me.lblCellLDim.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblCellLDim.Location = New System.Drawing.Point(340, 151)
        Me.lblCellLDim.Name = "lblCellLDim"
        Me.lblCellLDim.Size = New System.Drawing.Size(27, 14)
        Me.lblCellLDim.TabIndex = 20
        Me.lblCellLDim.Tag = "0"
        Me.lblCellLDim.Text = "mm"
        '
        'lblCellHDim
        '
        Me.lblCellHDim.AutoSize = True
        Me.lblCellHDim.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.lblCellHDim.Location = New System.Drawing.Point(340, 179)
        Me.lblCellHDim.Name = "lblCellHDim"
        Me.lblCellHDim.Size = New System.Drawing.Size(27, 14)
        Me.lblCellHDim.TabIndex = 21
        Me.lblCellHDim.Tag = "0"
        Me.lblCellHDim.Text = "mm"
        '
        'nudColumns
        '
        Me.nudColumns.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.nudColumns.Location = New System.Drawing.Point(271, 75)
        Me.nudColumns.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudColumns.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudColumns.Name = "nudColumns"
        Me.nudColumns.Size = New System.Drawing.Size(60, 22)
        Me.nudColumns.TabIndex = 6
        Me.nudColumns.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudColumns.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'nudRows
        '
        Me.nudRows.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.nudRows.Location = New System.Drawing.Point(271, 102)
        Me.nudRows.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudRows.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudRows.Name = "nudRows"
        Me.nudRows.Size = New System.Drawing.Size(60, 22)
        Me.nudRows.TabIndex = 7
        Me.nudRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudRows.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label7.Location = New System.Drawing.Point(15, 10)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(182, 14)
        Me.Label7.TabIndex = 22
        Me.Label7.Tag = ""
        Me.Label7.Text = "Имя таблицы (необязательно)"
        '
        'txtNameTable
        '
        Me.txtNameTable.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.txtNameTable.Location = New System.Drawing.Point(15, 34)
        Me.txtNameTable.Name = "txtNameTable"
        Me.txtNameTable.Size = New System.Drawing.Size(349, 22)
        Me.txtNameTable.TabIndex = 23
        Me.txtNameTable.Tag = ""
        Me.txtNameTable.Text = "TbL"
        '
        'dlgNewTable
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(379, 333)
        Me.Controls.Add(Me.txtNameTable)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.nudRows)
        Me.Controls.Add(Me.nudColumns)
        Me.Controls.Add(Me.lblCellHDim)
        Me.Controls.Add(Me.lblCellLDim)
        Me.Controls.Add(Me.lblTableHDim)
        Me.Controls.Add(Me.lblTableLDim)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtTableCusHeight)
        Me.Controls.Add(Me.txtTableCusWidth)
        Me.Controls.Add(Me.txtCellDefHeight)
        Me.Controls.Add(Me.txtCellDefWidth)
        Me.Controls.Add(Me.ckbDelShape)
        Me.Controls.Add(Me.optInside)
        Me.Controls.Add(Me.optCustom)
        Me.Controls.Add(Me.optPage)
        Me.Controls.Add(Me.optDefault)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgNewTable"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Создание новой таблицы"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.nudColumns, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudRows, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents optDefault As System.Windows.Forms.RadioButton
    Friend WithEvents optPage As System.Windows.Forms.RadioButton
    Friend WithEvents optCustom As System.Windows.Forms.RadioButton
    Friend WithEvents optInside As System.Windows.Forms.RadioButton
    Friend WithEvents ckbDelShape As System.Windows.Forms.CheckBox
    Friend WithEvents txtCellDefHeight As System.Windows.Forms.TextBox
    Friend WithEvents txtCellDefWidth As System.Windows.Forms.TextBox
    Friend WithEvents txtTableCusHeight As System.Windows.Forms.TextBox
    Friend WithEvents txtTableCusWidth As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblTableLDim As System.Windows.Forms.Label
    Friend WithEvents lblTableHDim As System.Windows.Forms.Label
    Friend WithEvents lblCellLDim As System.Windows.Forms.Label
    Friend WithEvents lblCellHDim As System.Windows.Forms.Label
    Friend WithEvents nudColumns As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudRows As System.Windows.Forms.NumericUpDown
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtNameTable As System.Windows.Forms.TextBox

End Class
