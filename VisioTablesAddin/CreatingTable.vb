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
    Dim NoDupes As New Collection

    Dim frmNewTable As System.Windows.Forms.Form = New dlgNewTable
    Public docObj As Visio.Document = vsoApp.ActiveDocument
    Public winObj As Visio.Window = vsoApp.ActiveWindow
    Public pagObj As Visio.Page = vsoApp.ActivePage
    Public shpsObj As Visio.Shapes = pagObj.Shapes
    Public selObj As Visio.Selection
    Public vsoSelection As Visio.Selection
    Dim MemSel As Visio.Shape

    Public NT As String = ""
    Dim LayerVisible As String = ""
    Dim LayerLock As String = ""

    Public Const UTN = "User.TableName"
    Public Const UTR = "User.TableRow"
    Public Const UTC = "User.TableCol"
    Public Const PX = "PinX"
    Public Const PY = "PinY"
    Public Const WI = "Width"
    Public Const HE = "Height"
    'Const CA = "Angle"
    Public Const LD = "LockDelete"
    Public Const GU = "=GUARD("
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

#Region "New Table Sub"

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

    Public Sub CreatTable(a, b, c, d, e, f, g, h, i, j)

        On Error GoTo errD

        Dim NewTable As VisioTable = New VisioTable(a, b, c, d, e, f, g, h, i, j)
        NewTable.CreatTable()
        NewTable = Nothing

        Exit Sub
errD:
        MessageBox.Show("CreatTable" & vbNewLine & Err.Description)
    End Sub

#End Region

#Region "Friend Sub"

    Sub AllAlignOnText(booOnWidth As Boolean, booOnHeight As Boolean, bytNothingOrAutoOrLockColumns As Byte, _
        bytNothingOrAutoOrLockRows As Byte, booOnlySelectedColumns As Boolean, booOnlySelectedRow As Boolean)
        ' Выравнивание/автовыравнивание ячеек таблицы по ширине/высоте текста. Предварительная процедура

        If Not CheckSelCells() Then Exit Sub

        Dim vsoSel As Visio.Selection 'NoDupes As New Collection ', strNameMainCtlCell As String
        Dim Shp As Visio.Shape, ShNum As Integer, iCount As Integer, bytColumnOrRow As Byte

        shpsObj = winObj.Page.Shapes : vsoSel = winObj.Selection
        Call InitArrShapeID(NT)

        vsoApp.ShowChanges = False

        With shpsObj

            If booOnWidth And Not booOnlySelectedColumns Then ' Автовыравнивание всех столбцов
                bytColumnOrRow = 4
                Call RecUndo("Выровнять все по ширине текста")
                For iCount = 1 To UBound(ArrShapeID, 1)
                    ShNum = .ItemFromID(ArrShapeID(iCount, 0)).Cells(UTC).Result("") : Call AlignOnText(ShNum, bytColumnOrRow, bytNothingOrAutoOrLockColumns)
                Next
                Call RecUndo("0")
            End If

            If booOnHeight And Not booOnlySelectedRow Then ' Автовыравнивание всех строк
                bytColumnOrRow = 5
                Call RecUndo("Выровнять все по высоте текста")
                For iCount = 1 To UBound(ArrShapeID, 2)
                    ShNum = .ItemFromID(ArrShapeID(0, iCount)).Cells(UTR).Result("") : Call AlignOnText(ShNum, bytColumnOrRow, bytNothingOrAutoOrLockRows)
                Next
                Call RecUndo("0")
            End If

            If booOnWidth And booOnlySelectedColumns Then ' Автовыравнивание только выделенных столбцов
                bytColumnOrRow = 4 : NotDub(vsoSel, UTC)
                Call RecUndo("Выровнять по ширине текста")
                For iCount = 1 To NoDupes.Count
                    If NoDupes(iCount) <> 0 Then Call AlignOnText(NoDupes(iCount), bytColumnOrRow)
                Next
                Call RecUndo("0")
                NoDupes.Clear()
            End If

            If booOnHeight And booOnlySelectedRow Then ' Автовыравнивание только выделенных строк
                bytColumnOrRow = 5 : NotDub(vsoSel, UTR)
                Call RecUndo("Выровнять по высоте текста")
                For iCount = 1 To NoDupes.Count
                    If NoDupes(iCount) <> 0 Then Call AlignOnText(NoDupes(iCount), bytColumnOrRow)
                Next
                Call RecUndo("0")
                NoDupes.Clear()
            End If

        End With

        vsoApp.ShowChanges = True
        winObj.Selection = vsoSel

    End Sub

    Sub AlignOnSize(arg As Byte)
        If Not CheckSelCells() Then Exit Sub

        Dim i As Integer, strCellWH As String = "", dblResult As Double
        Dim vsoSel As Visio.Selection
        shpsObj = winObj.Page.Shapes

        Call InitArrShapeID(NT)

        Select Case arg
            Case 4
                ClearControlCells(UTC)
            Case 5
                ClearControlCells(UTR)
        End Select

        If winObj.Selection.Count = 0 Then GoTo err

        vsoSel = winObj.Selection
        If vsoSel.Count = 1 Then Call SelCell(2)

        vsoSelection = winObj.Selection

        Select Case arg
            Case 4
                NotDub(vsoSelection, UTC) : strCellWH = WI
                Call RecUndo("Выровнять ширину столбцов")
            Case 5
                NotDub(vsoSelection, UTR) : strCellWH = HE
                Call RecUndo("Выровнять высоту строк")
        End Select

        With shpsObj
            For i = 1 To NoDupes.Count
                Select Case arg
                    Case 4
                        dblResult = dblResult + .ItemFromID(ArrShapeID(NoDupes(i), 0)).Cells(strCellWH).Result(64)
                    Case 5
                        dblResult = dblResult + .ItemFromID(ArrShapeID(0, NoDupes(i))).Cells(strCellWH).Result(64)
                End Select
            Next
            dblResult = dblResult / NoDupes.Count

            For i = 1 To NoDupes.Count
                Select Case arg
                    Case 4
                        .ItemFromID(ArrShapeID(NoDupes(i), 0)).Cells(strCellWH).Result(64) = dblResult
                    Case 5
                        .ItemFromID(ArrShapeID(0, NoDupes(i))).Cells(strCellWH).Result(64) = dblResult
                End Select
            Next
        End With
        NoDupes.Clear()
        Call RecUndo("0")
        winObj.DeselectAll() : winObj.Selection = vsoSel

err:
    End Sub

    Sub AllRotateText(iAng) 'Поворот текста
        If Not CheckSelCells() Then Exit Sub

        Dim i As Integer
        vsoSelection = winObj.Selection
        Call RecUndo("Поворот текста")

        For i = 1 To vsoSelection.Count
            With vsoSelection(i)
                If .CellExistsU(UTN, 0) Then
                    If InStr(1, .Name, "ClW", 1) <> 0 Then
                        .Cells("TxtPinX").FormulaU = "Width*0.5"
                        .Cells("TxtPinY").FormulaU = "Height*0.5"
                        .Cells("TxtLocPinX").FormulaU = "TxtWidth*0.5"
                        .Cells("TxtLocPinY").FormulaU = "TxtHeight*0.5"
                        If iAng = 0 Or iAng = 180 Then
                            .Cells("TxtWidth").FormulaU = "Width*1"
                            .Cells("TxtHeight").FormulaU = "Height*1"
                        Else
                            .Cells("TxtWidth").FormulaU = "Height*1"
                            .Cells("TxtHeight").FormulaU = "Width*1"
                        End If
                        .Cells("TxtAngle").FormulaU = Val(iAng) & " deg"
                    End If
                End If
            End With
        Next

        Call RecUndo("0")

    End Sub

    Sub AlternatLines(iAlt As Byte)  'Чередование цвета строк/столбцов
        If Not CheckSelCells() Then Exit Sub
        Dim i As Integer, j As Integer, strCellWH As String = UTC
        shpsObj = winObj.Page.Shapes : MemSel = winObj.Selection(1)

        Call InitArrShapeID(NT)

        If iAlt = 5 Then strCellWH = UTR

        Call RecUndo("Чередование строк/столбцов")

        With shpsObj
            For i = 1 To UBound(ArrShapeID, 1)
                For j = 1 To UBound(ArrShapeID, 2)
                    If ArrShapeID(i, j) <> 0 Then
                        If .ItemFromID(ArrShapeID(i, j)).Cells(strCellWH).Result("") Mod 2 = 0 Then .ItemFromID(ArrShapeID(i, j)).Cells("FillForegnd").FormulaU = "THEMEGUARD(MSOTINT(RGB(255,255,255),-10))"
                    End If
                Next
            Next
        End With

        Call RecUndo("0")

        winObj.DeselectAll() : winObj.Select(MemSel, 2)

    End Sub

    Sub ClearControlCells(arg)   ' Deselect УЯ столбцов или строк
        Dim shObj As Visio.Shape

        With winObj
            For Each shObj In .Selection
                If shObj.Cells(arg).Result("") = 0 Then .Select(shObj, 1)
                If shObj.Name = NT Then .Select(shObj, 1)
            Next
        End With

    End Sub

    Sub ConvertInto1Shape() ' Преобразование таблицы в одну сгруппированную фигуру
        If Not CheckSelCells() Then Exit Sub

        Dim visWorkCells As Visio.Selection, i As Integer
        Dim dblTop As Double, dblBottom As Double, dblLeft As Double, dblRight As Double
        'winObj = ActiveWindow

        Call InitArrShapeID(winObj.Selection(1).Cells(UTN).ResultStr(""))
        winObj.Page.Shapes.ItemU(NT).BoundingBox(1, dblLeft, dblBottom, dblRight, dblTop)

        Call SelCell(2) : visWorkCells = winObj.Selection
        winObj.DeselectAll()

        Call SelectCells(0, UBound(ArrShapeID, 1), 0, 0)
        Call SelectCells(0, 0, 1, UBound(ArrShapeID, 2))

        Call RecUndo("Преобразовать в 1 фигуру")
        On Error GoTo err
        With winObj
            .Selection.Group()
            .Selection.DeleteEx(0)
            .Selection = visWorkCells
            .Selection.Group()
        End With

        For i = 1 To visWorkCells.Count
            With visWorkCells(i)
                .DeleteSection(242)
                .DeleteSection(240)
                '.DeleteSection visSectionProp ' Как насчет таблицы с внешними данными?
                .CellsSRC(1, 17, 16).FormulaU = ""
                .CellsSRC(1, 15, 5).FormulaForceU = "0"
                .CellsSRC(1, 15, 8).FormulaForceU = "0"
                .CellsSRC(1, 2, 3).FormulaForceU = ""
                .CellsSRC(1, 6, 0).FormulaU = ""
            End With
        Next

        With winObj.Selection(1)
            .Cells(PX).Result("") = dblLeft + (.Cells(PX).Result(""))
            .Cells(PY).Result("") = dblBottom - (.Cells(PY).Result(""))
            .Name = NT
        End With

err:
        Call RecUndo("0")
    End Sub

    Sub DelTab() ' Удаление активной таблицы. Основная процедура
        On Error GoTo errD
        Dim Response As Byte = 0
        ' 6 - Да, 7 - нет, 2 - отмена
        Call CheckSelCells()

        'If Response = 0 Then
        Response = MsgBox("Уверены что хотите удалить эту таблицу?", 67, "Удаление!")
        'End If

        If Response = 6 Then
            winObj = vsoApp.ActiveWindow
            shpsObj = winObj.Page.Shapes
            NT = winObj.Selection(1).Cells(UTN).ResultStr("")
            Call RecUndo("Удалить таблицу")

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
            Call RecUndo(0)
        End If
        Exit Sub
errD:
        MessageBox.Show("DelTab" & vbNewLine & Err.Description)
    End Sub

    Sub InitArrShapeID(strNameShape)  ' Заполнение массива шейпами активной таблицы
        'strNameTable - строковая переменая, значение ячейки "User.TableName" любого шейпа из активной таблицы

        Dim shObj As Visio.Shape
        Dim cMax As Integer = shpsObj.Item(strNameShape).Cells(UTC).Result("")
        Dim rMax As Integer = shpsObj.Item(strNameShape).Cells(UTR).Result("")

        ReDim ArrShapeID(cMax, rMax)

        For Each shObj In shpsObj
            With shObj
                If .CellExistsU(UTN, 0) Then
                    If StrComp(.Cells(UTN).ResultStr(""), strNameShape) = 0 Then
                        'MsgBox(.Cells(UTC).Result("") & " x " & .Cells(UTR).Result(""))
                        ArrShapeID(.Cells(UTC).Result(""), .Cells(UTR).Result("")) = .ID
                    End If
                End If
            End With
        Next
        ArrShapeID(0, 0) = shpsObj.Item(strNameShape).ID
        If ArrShapeID(0, 0) = ArrShapeID(cMax, rMax) Then ArrShapeID(cMax, rMax) = 0
    End Sub

    Sub InsertText(arg) ' Вставить в ячейки текст, дату, время, комментарий, номер столбца, номер строки
        If Not CheckSelCells() Then Exit Sub
        Dim title As String, msgComm As String, txt As String = "", i As Integer
        Dim vsoSel As Visio.Selection = winObj.Selection, arrArg() As String

        title = "Вставить в ячейки"
        Call RecUndo("Вставить в ячейки")

        Dim TextInsert = Sub()
                             If txt <> "" Then
                                 For i = 1 To vsoSel.Count
                                     vsoSel(i).Characters.Text = txt
                                 Next
                             End If
                         End Sub

        Dim NumInsert = Sub()
                            For i = 1 To vsoSel.Count
                                vsoSel(i).Characters.AddCustomFieldU(txt, 0)
                            Next
                        End Sub

        Select Case arg ' Надо поправить RecUndo
            Case 0 : txt = InputBox("Вставить текст", title, "Текст...") : TextInsert()
            Case 1 : txt = InputBox("Вставить дату", title, Today) : TextInsert()
            Case 2 : txt = InputBox("Вставить время", title, TimeString) : TextInsert()
            Case 3
                msgComm = "0 - Восстановить по умолчанию" & vbCrLf & "1 - Текст ячейки в комментарий" & vbCrLf & "2 - Текст комментария в ячейку" & vbCrLf
                txt = InputBox("Комментарии:" & vbCrLf & msgComm, title, "Комментарий...")
                Select Case txt
                    Case "" ' Отмена
                        GoTo err
                    Case 0 ' Восстановить по умолчанию
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = "Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                        Next
                    Case 1 ' Текст ячейки в комментарий
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = Chr(34) & vsoSel(i).Characters.Text & Chr(34)
                        Next
                    Case 2 ' Текст комментария в ячейку
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Characters.Text = vsoSel(i).Cells("Comment").ResultStr("")
                        Next
                    Case Else ' Комментарий пользователя
                        For i = 1 To vsoSel.Count
                            vsoSel(i).Cells("Comment").FormulaForceU = Chr(34) & txt & Chr(34)
                        Next
                End Select
            Case 4
                msgComm = "Формат:" & vbCrLf & "Префикс, Смещение, Постфикс"
                txt = InputBox("Вставить номер столбца" & vbCrLf & msgComm & vbCrLf, title, ",0,")
                If txt = "" Then GoTo err
                arrArg = Split(txt, ",")
                If UBound(arrArg) <> 2 Then GoTo err
                txt = """" & arrArg(0) & """" & "&" & "User.TableCol" & "+" & Int(Val(arrArg(1))) & "&" & """" & arrArg(2) & """" : NumInsert()
            Case 5
                msgComm = "Формат:" & vbCrLf & "Префикс, Смещение, Постфикс"
                txt = InputBox("Вставить номер столбца" & vbCrLf & msgComm & vbCrLf, title, ",0,")
                If txt = "" Then GoTo err
                arrArg = Split(txt, ",")
                If UBound(arrArg) <> 2 Then GoTo err
                txt = """" & arrArg(0) & """" & "&" & "User.TableRow" & "+" & Int(Val(arrArg(1))) & "&" & """" & arrArg(2) & """" : NumInsert()
        End Select

err:
        Call RecUndo("0")
    End Sub

    Sub PropLayers(arg As Byte) ' Включение/выключение видимости и блокировки слоев на время выполнения кода - Titles_Tables и Cells_Tables

        With winObj.Page.Layers
            Select Case arg
                Case 1
                    LayerVisible = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaU
                    '            LayerVisible1 = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0) + 2).CellsC(4).FormulaU
                    LayerLock = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(7).FormulaU
                    '            LayerLock1 = .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0) + 2).CellsC(7).FormulaU
                    If LayerVisible <> 1 Then .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaForceU = "1"
                    If LayerLock <> 0 Then .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).Result("").CellsC(7).FormulaForceU = "0"
                Case 0
                    .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(4).FormulaForceU = LayerVisible
                    .Item(shpsObj.Item(NT).CellsSRC(1, 6, 0).Result("") + 1).CellsC(7).FormulaForceU = LayerLock
            End Select
        End With

    End Sub

    Sub RecUndo(index) ' Сохранение данных для операций Undo, Redo
        If index <> "0" Then
            UndoScopeID = vsoApp.BeginUndoScope(index)
        Else
            vsoApp.EndUndoScope(UndoScopeID, True)
        End If
    End Sub

    Sub SelCell(arg As Byte) ' Выделение(разное) ячеек таблицы
        If Not CheckSelCells() Then Exit Sub

        Dim vsoSel As Visio.Selection, intMaxC As Integer, intMaxR As Integer
        Dim iCount As Integer, UT As String, Shp As Visio.Shape

        vsoSel = winObj.Selection
        Call InitArrShapeID(NT)
        winObj.DeselectAll()

        intMaxC = UBound(ArrShapeID, 1) : intMaxR = UBound(ArrShapeID, 2)

        Select Case arg

            Case 1 ' Выделение таблицы с УЯ
                Call SelectCells(0, intMaxC, 0, intMaxR)

            Case 2 ' Выделение таблицы без УЯ
                Call SelectCells(1, intMaxC, 1, intMaxR)

            Case 3 ' Выделение диапазона ячеек
                If vsoSel.Count < 2 Then
                    MsgBox("Должно быть выделено не меньше двух ячеек.", 48, "Ошибка!")
                    GoTo err
                Else
                    Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer
                    Call ClearControlCells(UTC) : Call ClearControlCells(UTR)
                    If Not GetMinMaxRange(vsoSel, cMin, cMax, rMin, rMax) Then GoTo err
                    If cMin = 0 Then cMin = 1
                    If rMin = 0 Then rMin = 1
                    Call SelectCells(cMin, cMax, rMin, rMax)
                End If

            Case 4 ' Выделение столбца
                UT = UTC : NotDub(vsoSel, UT)
                For iCount = 1 To NoDupes.Count
                    Call SelectCells(NoDupes(iCount), NoDupes(iCount), 1, intMaxR)
                Next
                NoDupes.Clear() : Shp = Nothing

            Case 5 ' Выделение строки
                UT = UTR : NotDub(vsoSel, UT)
                For iCount = 1 To NoDupes.Count
                    Call SelectCells(1, intMaxC, NoDupes(iCount), NoDupes(iCount))
                Next
                NoDupes.Clear() : Shp = Nothing

            Case 6 ' Выделение  УЯ столбцов
                Call SelectCells(1, intMaxC, 0, 0)

            Case 7 ' Выделение  УЯ строк
                Call SelectCells(0, 0, 1, intMaxR)

        End Select

        Exit Sub

err:
    End Sub

    Sub SelInContent(arg) ' Выделение ячеек таблицы по критерию(текст, дата, значение, пустые/не пустые). Основная процедура
        If Not CheckSelCells() Then Exit Sub

        Dim vsoSel As Visio.Selection = winObj.Selection, shpObj As Visio.Shape

        Dim SelShp = Sub() winObj.Select(shpObj, 1)

        If arg = 8 Then
            Call InitArrShapeID(NT)
            Call SelectCells(1, UBound(ArrShapeID, 1), 1, UBound(ArrShapeID, 2))
        End If

        For Each shpObj In vsoSel
            With shpObj
                If .Cells(UTC).Result("") = 0 Or .Cells(UTR).Result("") = 0 Then SelShp()
                Select Case arg
                    Case 1 'Текст
                        If IsNumeric(.Characters.Text) Or _
                        Trim(.Characters.Text) = "" Or _
                        IsDate(.Characters.Text) Then SelShp()
                    Case 2 'Числа
                        If Not IsNumeric(.Characters.Text) Then SelShp()
                    Case 3 'Даты
                        If Not IsDate(.Characters.Text) Or IsNumeric(.Characters.Text) Then SelShp()
                        '        Case 4 'Время
                        '            If Not IsDate(.Characters.Text) Or Not .Characters.Text Like "*:*:*" Then GoSub Sub1
                    Case 5 'Не числа
                        If IsNumeric(.Characters.Text) Or _
                        Trim(.Characters.Text) = "" Then SelShp()
                    Case 6 'Пустые
                        If Trim(.Characters.Text) <> "" Then SelShp()
                    Case 7 'Не пустые
                        If Trim(.Characters.Text) = "" Then SelShp()
                    Case 8 'Инвертировать относительно таблицы
                        SelShp()
                End Select
            End With
        Next

    End Sub

    Sub SelectCells(intStartCol As Integer, intEndCol As Integer, intStartRow As Integer, intEndRow As Integer) 'Различное выделение в таблице

        Dim intColNum As Integer, intRowNum As Integer

        On Error Resume Next
        For intColNum = intStartCol To intEndCol
            For intRowNum = intStartRow To intEndRow
                winObj.Select(winObj.Page.Shapes.ItemFromID(ArrShapeID(intColNum, intRowNum)), 2)
            Next
        Next
        On Error GoTo 0
    End Sub

#End Region

#Region "Private Sub"

    Private Sub AlignOnText(ShNum As Integer, bytColumnOrRow As Byte, Optional bytNothingOrAutoOrLock As Byte = 0)
        ' Выравнивание/автовыравнивание ячеек таблицы по ширине/высоте текста. Основная процедура

        Dim cellName As String = "", txt As String = "", txt1 As String, txt2 As String, lentxt As Integer
        Dim intCount As Integer, iC As Integer, iR As Integer

        Select Case bytColumnOrRow
            Case 4
                cellName = WI : txt = "MAX(TEXTWIDTH(" : txt1 = "!TheText),TEXTWIDTH(" : lentxt = Len(txt)
                For intCount = 1 To UBound(ArrShapeID, 2)
                    With shpsObj.ItemFromID(ArrShapeID(ShNum, intCount))
                        If InStr(1, .Cells(cellName).FormulaU, ",", 1) = 0 _
                            And ArrShapeID(ShNum, intCount) <> 0 Then txt = txt & .Name & txt1
                    End With
                Next
                iC = ShNum : iR = 0
            Case 5
                cellName = HE : txt = "MAX(TEXTHEIGHT(" : txt1 = "!TheText," : txt2 = "!Width),TEXTHEIGHT(" : lentxt = Len(txt)
                For intCount = 1 To UBound(ArrShapeID, 1)
                    With shpsObj.ItemFromID(ArrShapeID(intCount, ShNum))
                        If InStr(1, .Cells(cellName).FormulaU, ",", 1) = 0 _
                            And ArrShapeID(intCount, ShNum) <> 0 Then txt = txt & .Name & txt1 & .Name & txt2
                    End With
                Next
                iC = 0 : iR = ShNum
        End Select

        On Error Resume Next
        With shpsObj.ItemFromID(ArrShapeID(iC, iR))
            .Cells(cellName).FormulaForceU = Left$(txt, Len(txt) - lentxt + 3) & ")"
            If bytNothingOrAutoOrLock = 0 Then .Cells(cellName).Result(64) = .Cells(cellName).Result(64)
            If bytNothingOrAutoOrLock = 2 Then .Cells(cellName).FormulaForceU = GU & Left$(txt, Len(txt) - lentxt + 3) & ")" & ")"
        End With

    End Sub

    Private Sub NotDub(vsoSel, UT) ' Заполнение коллекции значениями без дубликатов
        Dim Shp As Visio.Shape

        On Error Resume Next

        For Each Shp In vsoSel
            NoDupes.Add(Shp.Cells(UT).Result(""), CStr(Shp.Cells(UT).Result("")))
        Next

        On Error GoTo 0

    End Sub

#End Region

#Region "Functions"

    Function GetMinMaxRange(ByVal vsoSel As Visio.Selection, ByRef cMin As Integer, ByRef cMax As Integer, ByRef rMin As Integer, ByRef rMax As Integer) As Boolean
        ' Функция определения минимального и максимального номера столбцов/строк среди выделенного диапазона ячеек
        Dim i As Integer
        rMin = 1000 : cMin = 1000 : rMax = 0 : cMax = 0

        On Error GoTo err

        For i = 1 To vsoSel.Count
            With vsoSel(i)
                If rMin > .Cells(UTR).Result("") Then rMin = .Cells(UTR).Result("")
                If cMin > .Cells(UTC).Result("") Then cMin = .Cells(UTC).Result("")
                If rMax < .Cells(UTR).Result("") Then rMax = .Cells(UTR).Result("")
                If cMax < .Cells(UTC).Result("") Then cMax = .Cells(UTC).Result("")
            End With
        Next

        GetMinMaxRange = True
        Exit Function

err:
        GetMinMaxRange = False
    End Function

    Function CheckSelCells() As Boolean ' Сообщение об отсутствующем/некорректном выделении на листе

        Dim ErrMsg = Sub()
                         MsgBox("На активном листе отсутствуют выделенные ячейки в таблице!" &
                            vbCrLf & "Дальнейшая работа невозможна." & vbCrLf &
                            "Нужно выбрать ячейку в таблице и выполнить операцию еще раз." & vbCrLf, 48, "Внимание")
                         CheckSelCells = False
                     End Sub

        With winObj
            If .Selection.Count = 0 Then ErrMsg()

            Dim shObj As Visio.Shape

            For Each shObj In .Selection
                If Not shObj.CellExistsU(UTN, 0) Then .Select(shObj, 1)
            Next
            If .Selection.Count = 0 Then ErrMsg()

            NT = .Selection(1).Cells(UTN).ResultStr("")

            For Each shObj In .Selection
                If StrComp(shObj.Cells(UTN).ResultStr(""), NT) <> 0 Then _
                .Select(shObj, 1)
            Next
        End With

        Return True
    End Function

#End Region

End Module
