<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgSortTable
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cb_SortingDirection = New System.Windows.Forms.CheckBox()
        Me.cb_DigitOrText = New System.Windows.Forms.CheckBox()
        Me.Num_Column = New System.Windows.Forms.NumericUpDown()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        CType(Me.Num_Column, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(181, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Целевой столбец для сортировки:"
        '
        'cb_SortingDirection
        '
        Me.cb_SortingDirection.AutoSize = True
        Me.cb_SortingDirection.Location = New System.Drawing.Point(12, 40)
        Me.cb_SortingDirection.Name = "cb_SortingDirection"
        Me.cb_SortingDirection.Size = New System.Drawing.Size(160, 17)
        Me.cb_SortingDirection.TabIndex = 1
        Me.cb_SortingDirection.Text = "Сортировать по убыванию"
        Me.cb_SortingDirection.UseVisualStyleBackColor = True
        '
        'cb_DigitOrText
        '
        Me.cb_DigitOrText.AutoSize = True
        Me.cb_DigitOrText.Location = New System.Drawing.Point(12, 63)
        Me.cb_DigitOrText.Name = "cb_DigitOrText"
        Me.cb_DigitOrText.Size = New System.Drawing.Size(205, 17)
        Me.cb_DigitOrText.TabIndex = 2
        Me.cb_DigitOrText.Text = "Сортировать как числовые данные"
        Me.cb_DigitOrText.UseVisualStyleBackColor = True
        '
        'Num_Column
        '
        Me.Num_Column.Location = New System.Drawing.Point(234, 8)
        Me.Num_Column.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.Num_Column.Name = "Num_Column"
        Me.Num_Column.Size = New System.Drawing.Size(66, 20)
        Me.Num_Column.TabIndex = 3
        Me.Num_Column.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.Num_Column.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'btn_OK
        '
        Me.btn_OK.Location = New System.Drawing.Point(142, 114)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(75, 23)
        Me.btn_OK.TabIndex = 4
        Me.btn_OK.Text = "OK"
        Me.btn_OK.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Cancel.Location = New System.Drawing.Point(225, 114)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.btn_Cancel.TabIndex = 5
        Me.btn_Cancel.Text = "Закрыть"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'dlgSortTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btn_Cancel
        Me.ClientSize = New System.Drawing.Size(312, 150)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.Num_Column)
        Me.Controls.Add(Me.cb_DigitOrText)
        Me.Controls.Add(Me.cb_SortingDirection)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgSortTable"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Сортировка данных"
        CType(Me.Num_Column, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cb_SortingDirection As System.Windows.Forms.CheckBox
    Friend WithEvents cb_DigitOrText As System.Windows.Forms.CheckBox
    Friend WithEvents Num_Column As System.Windows.Forms.NumericUpDown
    Friend WithEvents btn_OK As System.Windows.Forms.Button
    Friend WithEvents btn_Cancel As System.Windows.Forms.Button
End Class
