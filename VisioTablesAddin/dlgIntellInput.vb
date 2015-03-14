Imports System.Windows.Forms

Public Class dlgIntellInput
    Private bytColOrRow As Byte
    Private ReturnPressed As Boolean
    Private winObj As Visio.Window, vsoObj As Visio.Selection
    Private iC As Integer, iR As Integer, arr1 As Integer, arr2 As Integer
    Private NoDupes As New Collection

    Private Sub Enter_down()
        '        vsoObj = winObj.Selection

        '        Dim SelectShape = Sub() winObj.Select(winObj.Page.Shapes.ItemFromID(ArrShapeID(iC, iR)), 258)

        '        Dim NextRow = Sub()
        '                          If iR <> arr2 Then
        '                              iR = iR + 1
        '                          Else
        '                              iR = 1
        '                              If iC <> arr1 Then
        '                                  iC = iC + 1
        '                              Else
        '                                  iC = 1
        '                              End If
        '                          End If
        '                          GoTo Line1
        '                      End Sub

        '        Dim NextColumn = Sub()
        '                             If iC <> arr1 Then
        '                                 iC = iC + 1
        '                             Else
        '                                 iC = 1
        '                                 If iR <> arr2 Then
        '                                     iR = iR + 1
        '                                 Else
        '                                     iR = 1
        '                                 End If
        '                             End If
        '                             GoTo Line2
        '                         End Sub
        '        On Error GoTo ex

        '        If ReturnPressed Then
        '            iC = vsoObj(1).Cells("User.TableCol").Result("") : iR = vsoObj(1).Cells("User.TableRow").Result("")
        '            arr1 = UBound(ArrShapeID, 1) : arr2 = UBound(ArrShapeID, 2)

        '            If Trim(cmbText.Text) <> "" Then
        '                If ckbSkipNotEmpty.Checked Then
        '                    If Trim(vsoObj(1).Characters.Text) = "" Then vsoObj(1).Characters.Text = cmbText.Text
        '                    Call SaveText()
        '                Else
        '                    vsoObj(1).Characters.Text = cmbText.Text
        '                    Call SaveText()
        '                End If
        '            End If


        '            Select Case bytColOrRow
        '                Case 1
        '                    NextRow()
        'Line1:
        '                    If ArrShapeID(iC, iR) <> 0 Then
        '                        SelectShape()
        '                    Else
        '                        NextRow()
        '                    End If
        '                Case 2
        '                    NextColumn()
        'Line2:
        '                    If ArrShapeID(iC, iR) <> 0 Then
        '                        SelectShape()
        '                    Else
        '                        NextColumn()
        '                    End If
        '            End Select
        '        End If

        '        Exit Sub
        'ex:
        '        'End
        '    End Sub

        'Private Sub cmbText_Exit(ByVal Cancel As MSForms.ReturnBoolean)
        '    Cancel = ReturnPressed
    End Sub

    Private Sub dlgIntellInput_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        winObj = Nothing : vsoObj = Nothing
        Erase ArrShapeID
    End Sub

    Private Sub dlgIntellInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'vsoApp = Globals.ThisAddIn.Application
        winObj = vsoApp.ActiveWindow
        optCol.Checked = True
    End Sub

    Private Sub optCol_Click(sender As Object, e As EventArgs) Handles optCol.Click
        If bytColOrRow = 0 Then Call InitArrShapeID(NT)
        bytColOrRow = 1
    End Sub

    Private Sub optRow_Click(sender As Object, e As EventArgs) Handles optRow.Click
        If bytColOrRow = 0 Then Call InitArrShapeID(NT)
        bytColOrRow = 2
    End Sub

    Private Sub cmbText_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbText.KeyDown
        If (e.KeyCode = Keys.Return) Then
            ReturnPressed = True
            Call Enter_down()
            cmbText.SelectionStart = 0
            cmbText.SelectionLength = Len(cmbText.Text)
        Else
            ReturnPressed = False
        End If
    End Sub

    Private Sub SaveText()
        On Error GoTo ex
        NoDupes.Add(cmbText.Text, CStr(cmbText.Text))
        cmbText.Items.Add(NoDupes(NoDupes.Count))
ex:
    End Sub

End Class