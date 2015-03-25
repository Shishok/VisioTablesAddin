<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgPictures
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optAlignRightH = New System.Windows.Forms.RadioButton()
        Me.optAlignCenterH = New System.Windows.Forms.RadioButton()
        Me.optAlignLeftH = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.optAlignBottomV = New System.Windows.Forms.RadioButton()
        Me.optAlignCenterV = New System.Windows.Forms.RadioButton()
        Me.optAlignTopV = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.ckbInsertName = New System.Windows.Forms.CheckBox()
        Me.ckbLockFormulas = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.optAlignRightH)
        Me.GroupBox1.Controls.Add(Me.optAlignCenterH)
        Me.GroupBox1.Controls.Add(Me.optAlignLeftH)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(256, 60)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Выравнивание по горизонтали"
        '
        'optAlignRightH
        '
        Me.optAlignRightH.AutoSize = True
        Me.optAlignRightH.Location = New System.Drawing.Point(186, 24)
        Me.optAlignRightH.Name = "optAlignRightH"
        Me.optAlignRightH.Size = New System.Drawing.Size(62, 17)
        Me.optAlignRightH.TabIndex = 2
        Me.optAlignRightH.Text = "Справа"
        Me.optAlignRightH.UseVisualStyleBackColor = True
        '
        'optAlignCenterH
        '
        Me.optAlignCenterH.AutoSize = True
        Me.optAlignCenterH.Location = New System.Drawing.Point(89, 24)
        Me.optAlignCenterH.Name = "optAlignCenterH"
        Me.optAlignCenterH.Size = New System.Drawing.Size(76, 17)
        Me.optAlignCenterH.TabIndex = 1
        Me.optAlignCenterH.Text = "По центру"
        Me.optAlignCenterH.UseVisualStyleBackColor = True
        '
        'optAlignLeftH
        '
        Me.optAlignLeftH.AutoSize = True
        Me.optAlignLeftH.Checked = True
        Me.optAlignLeftH.Location = New System.Drawing.Point(12, 24)
        Me.optAlignLeftH.Name = "optAlignLeftH"
        Me.optAlignLeftH.Size = New System.Drawing.Size(56, 17)
        Me.optAlignLeftH.TabIndex = 0
        Me.optAlignLeftH.TabStop = True
        Me.optAlignLeftH.Text = "Слева"
        Me.optAlignLeftH.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.optAlignBottomV)
        Me.GroupBox2.Controls.Add(Me.optAlignCenterV)
        Me.GroupBox2.Controls.Add(Me.optAlignTopV)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 76)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(256, 60)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Выравнивание по вертикали"
        '
        'optAlignBottomV
        '
        Me.optAlignBottomV.AutoSize = True
        Me.optAlignBottomV.Location = New System.Drawing.Point(186, 24)
        Me.optAlignBottomV.Name = "optAlignBottomV"
        Me.optAlignBottomV.Size = New System.Drawing.Size(55, 17)
        Me.optAlignBottomV.TabIndex = 5
        Me.optAlignBottomV.Text = "Снизу"
        Me.optAlignBottomV.UseVisualStyleBackColor = True
        '
        'optAlignCenterV
        '
        Me.optAlignCenterV.AutoSize = True
        Me.optAlignCenterV.Checked = True
        Me.optAlignCenterV.Location = New System.Drawing.Point(89, 24)
        Me.optAlignCenterV.Name = "optAlignCenterV"
        Me.optAlignCenterV.Size = New System.Drawing.Size(76, 17)
        Me.optAlignCenterV.TabIndex = 4
        Me.optAlignCenterV.TabStop = True
        Me.optAlignCenterV.Text = "По центру"
        Me.optAlignCenterV.UseVisualStyleBackColor = True
        '
        'optAlignTopV
        '
        Me.optAlignTopV.AutoSize = True
        Me.optAlignTopV.Location = New System.Drawing.Point(12, 24)
        Me.optAlignTopV.Name = "optAlignTopV"
        Me.optAlignTopV.Size = New System.Drawing.Size(60, 17)
        Me.optAlignTopV.TabIndex = 3
        Me.optAlignTopV.Text = "Сверху"
        Me.optAlignTopV.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(118, 206)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 8
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
        Me.Cancel_Button.TabIndex = 9
        Me.Cancel_Button.Text = "Отмена"
        '
        'ckbInsertName
        '
        Me.ckbInsertName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckbInsertName.AutoSize = True
        Me.ckbInsertName.Location = New System.Drawing.Point(8, 150)
        Me.ckbInsertName.Name = "ckbInsertName"
        Me.ckbInsertName.Size = New System.Drawing.Size(194, 17)
        Me.ckbInsertName.TabIndex = 6
        Me.ckbInsertName.Text = "Вставлять имена фигур в ячейки"
        Me.ckbInsertName.UseVisualStyleBackColor = True
        '
        'ckbLockFormulas
        '
        Me.ckbLockFormulas.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ckbLockFormulas.AutoSize = True
        Me.ckbLockFormulas.Location = New System.Drawing.Point(8, 173)
        Me.ckbLockFormulas.Name = "ckbLockFormulas"
        Me.ckbLockFormulas.Size = New System.Drawing.Size(168, 17)
        Me.ckbLockFormulas.TabIndex = 7
        Me.ckbLockFormulas.Text = "Заблокировать координаты"
        Me.ckbLockFormulas.UseVisualStyleBackColor = True
        '
        'dlgPictures
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(272, 247)
        Me.Controls.Add(Me.ckbLockFormulas)
        Me.Controls.Add(Me.ckbInsertName)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(288, 285)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(288, 285)
        Me.Name = "dlgPictures"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Настройки"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optAlignRightH As System.Windows.Forms.RadioButton
    Friend WithEvents optAlignCenterH As System.Windows.Forms.RadioButton
    Friend WithEvents optAlignLeftH As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents optAlignBottomV As System.Windows.Forms.RadioButton
    Friend WithEvents optAlignCenterV As System.Windows.Forms.RadioButton
    Friend WithEvents optAlignTopV As System.Windows.Forms.RadioButton
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents ckbInsertName As System.Windows.Forms.CheckBox
    Friend WithEvents ckbLockFormulas As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
