<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgIntellInput
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
        Me.cmbText = New System.Windows.Forms.ComboBox()
        Me.optCol = New System.Windows.Forms.RadioButton()
        Me.optRow = New System.Windows.Forms.RadioButton()
        Me.ckbSkipNotEmpty = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'cmbText
        '
        Me.cmbText.FormattingEnabled = True
        Me.cmbText.Location = New System.Drawing.Point(12, 35)
        Me.cmbText.Name = "cmbText"
        Me.cmbText.Size = New System.Drawing.Size(387, 21)
        Me.cmbText.TabIndex = 0
        '
        'optCol
        '
        Me.optCol.AutoSize = True
        Me.optCol.Location = New System.Drawing.Point(12, 12)
        Me.optCol.Name = "optCol"
        Me.optCol.Size = New System.Drawing.Size(91, 17)
        Me.optCol.TabIndex = 1
        Me.optCol.TabStop = True
        Me.optCol.Text = "По столбцам"
        Me.optCol.UseVisualStyleBackColor = True
        '
        'optRow
        '
        Me.optRow.AutoSize = True
        Me.optRow.Location = New System.Drawing.Point(119, 12)
        Me.optRow.Name = "optRow"
        Me.optRow.Size = New System.Drawing.Size(85, 17)
        Me.optRow.TabIndex = 2
        Me.optRow.TabStop = True
        Me.optRow.Text = "По строкам"
        Me.optRow.UseVisualStyleBackColor = True
        '
        'ckbSkipNotEmpty
        '
        Me.ckbSkipNotEmpty.AutoSize = True
        Me.ckbSkipNotEmpty.Location = New System.Drawing.Point(221, 12)
        Me.ckbSkipNotEmpty.Name = "ckbSkipNotEmpty"
        Me.ckbSkipNotEmpty.Size = New System.Drawing.Size(178, 17)
        Me.ckbSkipNotEmpty.TabIndex = 3
        Me.ckbSkipNotEmpty.Text = "Пропускать не пустые ячейки"
        Me.ckbSkipNotEmpty.UseVisualStyleBackColor = True
        '
        'dlgIntellInput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(412, 72)
        Me.Controls.Add(Me.ckbSkipNotEmpty)
        Me.Controls.Add(Me.optRow)
        Me.Controls.Add(Me.optCol)
        Me.Controls.Add(Me.cmbText)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgIntellInput"
        Me.Text = "Ввод текста с переходом по ячейкам"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbText As System.Windows.Forms.ComboBox
    Friend WithEvents optCol As System.Windows.Forms.RadioButton
    Friend WithEvents optRow As System.Windows.Forms.RadioButton
    Friend WithEvents ckbSkipNotEmpty As System.Windows.Forms.CheckBox
End Class
