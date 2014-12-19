Imports System.Drawing
Imports System.Windows.Forms

Module CreatingTable

#Region "LIST OF VARIABLES AND CONSTANTS"

    Public vsoApp As Visio.Application = Globals.ThisAddIn.Application
    Public UndoScopeID As Long = 0
    Public booOpenForm As Boolean = False
    Public strNameTable As String = ""
    Public FlagPage As Byte = 0
    Public ArrShapeID(,) As Integer

    Dim frmNewTable As System.Windows.Forms.Form = New dlgNewTable
    Dim docObj As Visio.Document = vsoApp.ActiveDocument
    Dim winObj As Visio.Window = vsoApp.ActiveWindow
    Dim pagObj As Visio.Page = vsoApp.ActivePage
    Dim shpsObj As Visio.Shapes = pagObj.Shapes
    Dim selObj As Visio.Selection
    Dim vsoSelection As Visio.Selection


    Dim NT As String = ""
    'Dim LayerVisible As String = ""
    'Dim LayerLock As String = ""

    Const UTN = "User.TableName"
    Const UTR = "User.TableRow"
    Const UTC = "User.TableCol"
    'Const PX = "PinX"
    'Const PY = "PinY"
    'Const WI = "Width"
    'Const HE = "Height"
    'Const CA = "Angle"
    Const LD = "LockDelete"
    'Const GU = "=GUARD("
    'Const strATC = "!Actions.Titles.Checked=1,"
    'Const strACC = "!Actions.Comments.Checked=1,"
    'Const strThGu000 = "THEMEGUARD(MSOTINT(RGB(0,0,0),50))"
    'Const strThGu255 = "THEMEGUARD(RGB(255,255,255))"
    'Const strThGu191 = "THEMEGUARD(RGB(191,191,191))"
    'Const GS = "=GUARD(Sheet."
    'Const GI = "=Guard(IF("
    'Const sh = "Sheet."
    'Const frm = "###0.0###"
    'Const GU5 = "=GUARD(10 mm)" ' Переделать на DrawUn
    'Const P50 = "50%"
    'Const GT = "GUARD(TRUE)"
    'Const G1 = "Guard(1)"


#End Region

    Public Sub QuickTable(strV)
        Dim nC As Integer = 0, nR As Integer = 0, w As Single = 0, h As Single = 0
        strV = Strings.Right(strV, Strings.Len(strV) - 1)
        nC = Val(Strings.Left(strV, Strings.InStr(1, strV, "x", 1) - 1))
        nR = Val(Strings.Right(strV, Strings.Len(strV) - Strings.InStr(1, strV, "x", 1)))
        w = vsoApp.FormatResult(20, 70, 64, "#.0000")
        h = vsoApp.FormatResult(8, 70, 64, "#.0000")
        strNameTable = "TbL"
        Call CreatTable(strNameTable, 1, nC, nR, w, h, 200, 150, False, False)
    End Sub

    Sub Load_dlgNewTable()
        If Not booOpenForm Then
            frmNewTable = New dlgNewTable
            frmNewTable.Show()
        End If
    End Sub

#Region "CREATTABLE"

    Public Sub CreatTable(strNameTable, bytInsertType, intColumnsCount, intRowsCount, sngWidthCells, sngHeightCells, _
            sngWidthTable, sngHeightTable, booDeleteTargetShape, booVisibleProgressBar)

        On Error GoTo errD

        Dim NewTable As VisioTable = New VisioTable(strNameTable, bytInsertType, intColumnsCount, intRowsCount, _
                sngWidthCells, sngHeightCells, sngWidthTable, sngHeightTable, booDeleteTargetShape, booVisibleProgressBar)


        NewTable.CreatTable()

        NewTable = Nothing

        Exit Sub
errD:
        MessageBox.Show("CreatTable1" & vbNewLine & Err.Description)
    End Sub

#End Region

    Public Sub DelTab() ' Удаление активной таблицы. Основная процедура
        On Error GoTo errD
        Dim Response As Byte = 0
        ' 6 - Да, 7 - нет, 2 - отмена
        'Call CheckSelCells()

        'If Response = 0 Then
        Response = MsgBox("Уверены что хотите удалить эту таблицу?", 67, "Удаление!")
        'End If

        If Response = 6 Then
            winObj = vsoApp.ActiveWindow
            shpsObj = winObj.Page.Shapes
            NT = winObj.Selection(1).Cells(UTN).ResultStr("")
            'Call RecUndo("Удалить таблицу")

            Dim frm As New dlgWait
            frm.Label1.Text = " " & vbCrLf & "Удаление таблицы..."
            frm.Show() : frm.Refresh()

            winObj.DeselectAll()
            vsoApp.ShowChanges = False

            Dim dblW As Integer, iCount As Integer
            dblW = shpsObj.Count

            For iCount = shpsObj.Count To 1 Step -1
                frm.lblProgressBar.Width = (300 / dblW) * iCount : frm.lblProgressBar.Refresh() : Application.DoEvents()
                With shpsObj.Item(iCount)
                    If .CellExistsU(UTN, 0) Then
                        If StrComp(.Cells(UTN).ResultStr(""), NT) = 0 Then
                            .Cells(LD).FormulaForceU = 0
                            .Delete()
                        End If
                    End If
                End With
            Next

            frm.Close()
            vsoApp.ShowChanges = True
            'Call RecUndo(0)
        End If
        Exit Sub
errD:
        MessageBox.Show("DelTab" & vbNewLine & Err.Description)
    End Sub

#Region "MiscProcedure"

    Sub RecUndo(index) ' Сохранение данных для операций Undo, Redo

        If index <> "0" Then
            UndoScopeID = vsoApp.BeginUndoScope(index)
        Else
            vsoApp.EndUndoScope(UndoScopeID, True)
        End If

    End Sub

    Public Sub InitArrShapeID(strNameShape As String)  ' Заполнение массива шейпами активной таблицы
        'strNameTable - строковая переменая, значение ячейки "User.TableName" любого шейпа из активной таблицы
        ' Переделать. Здесь определять имя ГУЯ

        Dim shObj As Visio.Shape
        Dim cMax As Integer, rMax As Integer
        'shpsObj = winObj.Page.Shapes

        rMax = shpsObj.Item(strNameShape).Cells(UTR).Result("")
        cMax = shpsObj.Item(strNameShape).Cells(UTC).Result("")

        ReDim ArrShapeID(cMax, rMax)

        For Each shObj In shpsObj
            With shObj
                If .CellExistsU(UTN, 0) Then
                    If StrComp(.Cells(UTN).ResultStr(""), strNameShape) = 0 Then
                        ArrShapeID(.Cells(UTC).Result(""), .Cells(UTR).Result("")) = .ID
                    End If
                End If
            End With
        Next
        ArrShapeID(0, 0) = shpsObj.Item(strNameShape).ID
        If ArrShapeID(0, 0) = ArrShapeID(cMax, rMax) Then ArrShapeID(cMax, rMax) = 0
    End Sub

#End Region

End Module
