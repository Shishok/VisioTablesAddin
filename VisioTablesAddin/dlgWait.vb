Imports System.Windows.Forms

Public Class dlgWait

    Private Sub dlgWait_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.Close()
    End Sub
End Class
