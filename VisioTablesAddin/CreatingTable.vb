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

#Region "Load Sub"

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
        Dim NewTable As VisioTable = New VisioTable(a, b, c, d, e, f, g, h, i, j)
        NewTable.CreatTable()
        NewTable = Nothing
    End Sub

    Sub LoaddlgPicture()
        Dim dlgNew As New dlgPictures
        dlgNew.ShowDialog()
        dlgNew = Nothing
    End Sub

#End Region

#Region "Friend Sub"

    Sub AddColumns(arg As Byte)  ' Вставка нового столбца. Основная процедура
        'arg = 0 вставка столбца перед выделенным, arg = 1 вставка столбца после выделенного

        Call CheckSelCells()
        Call ClearControlCells(UTC)
        If winObj.Selection.Count = 0 Then Exit Sub

        shpsObj = winObj.Page.Shapes

        Dim shObj As Visio.Shape, vs As Visio.Selection, vsoDups As Visio.Selection, i As Integer, j As Integer
        Dim iAll As Integer, nCol As Integer, NTNew As String, strF As String

        shObj = winObj.Selection(1)

        Call InitArrShapeID(NT) : winObj.DeselectAll()

        Call PropLayers(1)

        Call SelectCells(shObj.Cells(UTC).Result(""), shObj.Cells(UTC).Result(""), 0, UBound(ArrShapeID, 2))

        vs = winObj.Selection
        iAll = shpsObj.Item(NT).Cells(UTC).Result("")
        nCol = shObj.Cells(UTC).Result("")

        Call RecUndo("Добавить столбец")
        vs.Duplicate() : vsoDups = winObj.Selection

        For i = 2 To vsoDups.Count
            With vsoDups(i)
                If Not .Characters.IsField Then .Characters.Text = ""
                .Cells("Comment").FormulaForceU = "=Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                If InStr(1, .Cells(WI).FormulaU, "SUM") <> 0 Then
                    .Cells(LD).FormulaForceU = 0
                    .Delete()
                End If
            End With
        Next

        NTNew = vsoDups(1).Name

        If arg = 0 Then   ' Вставка столбца перед выделенным
            With vs(1)
                .Cells(PX).FormulaForceU = GU & NTNew & "!PinX+(" & NTNew & "!Width/2)+(Width/2))"
            End With
            For i = nCol To UBound(ArrShapeID, 1) ' Перенумерация управляющих ячеек
                shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).Result("") + 1 & ")"
            Next

        ElseIf arg = 1 Then  ' Вставка столбца после выделенного
            With vsoDups(1)
                .Cells(PX).FormulaForceU = GU & vs(1).Name & "!PinX+(" & vs(1).Name & "!Width/2)+(Width/2))"
                .Cells(UTC).FormulaForceU = GU & .Cells(UTC).Result("") + 1 & ")"
            End With
            If nCol <> shpsObj.Item(NT).Cells(UTC).Result("") Then
                shpsObj.ItemFromID(ArrShapeID(vs(1).Cells(UTC).Result("") + 1, 0)).Cells(PX).FormulaForceU = GU & vsoDups(1).Name & "!PinX+(" & vsoDups(1).Name & "!Width/2)+(Width/2))"
                For i = nCol + 1 To UBound(ArrShapeID, 1) ' Перенумерация управляющих ячеек
                    shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(i, 0)).Cells(UTC).Result("") + 1 & ")"
                Next
            End If
        End If

        With shpsObj
            .Item(NT).Cells(UTC).FormulaForceU = "GUARD(" & iAll + 1 & ")"
            .Item(NTNew).Cells("Controls.ControlWidth").FormulaForceU = "Width*1"
            .Item(NTNew).SendToBack()

            If vsoDups.Count <> iAll + 1 Then ' Определение объединенных ячеек и их обработка
                For j = 1 To UBound(ArrShapeID, 1)
                    For i = 1 To UBound(ArrShapeID, 2)
                        If ArrShapeID(j, i) <> 0 Then
                            With .ItemFromID(ArrShapeID(j, i))
                                If InStr(1, .Cells(WI).FormulaU, "SUM", 1) <> 0 Then
                                    If InStr(1, .Cells(WI).FormulaU, vs(1).Name & "!", 1) <> 0 Then
                                        If arg = 0 Then
                                            .Cells(WI).FormulaForceU = Replace$(.Cells(WI).FormulaU, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            strF = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!PinX-", NTNew & "!PinX-", 1)
                                            strF = Replace$(strF, vs(1).Name & "!Width/2", NTNew & "!Width/2", 1)
                                            .Cells(PX).FormulaForceU = Replace$(strF, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            .Cells(UTC).FormulaForceU = Replace$(.Cells(UTC).FormulaU, vs(1).Name & "!", NTNew & "!", 1)
                                        End If
                                        If arg = 1 Then
                                            .Cells(WI).FormulaForceU = Replace$(.Cells(WI).FormulaU, vs(1).Name & "!Width", NTNew & "!Width" & "," & vs(1).Name & "!Width", 1)
                                            .Cells(PX).FormulaForceU = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!Width,", vs(1).Name & "!Width" & "," & NTNew & "!Width,", 1)
                                            .Cells(PX).FormulaForceU = Replace$(.Cells(PX).FormulaU, vs(1).Name & "!Width)", vs(1).Name & "!Width" & "," & NTNew & "!Width)", 1)
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    Next
                Next
            End If
        End With

        Call RecUndo("0")
        On Error Resume Next
        winObj.Selection = vsoDups
        Call PropLayers(0)

    End Sub

    Sub AddRows(arg As Byte) ' Вставка новой строки. Основная процедура
        'arg = 0 вставка строки перед выделенной, arg = 1 вставка строки после выделенной

        Call CheckSelCells()
        Call ClearControlCells(UTR)
        If winObj.Selection.Count = 0 Then Exit Sub

        shpsObj = winObj.Page.Shapes

        Dim shObj As Visio.Shape, vs As Visio.Selection, vsoDups As Visio.Selection, i As Integer, j As Integer
        Dim iAll As Integer, nRow As Integer, NTNew As String, strF As String

        shObj = winObj.Selection(1)

        Call InitArrShapeID(NT) : winObj.DeselectAll()

        Call PropLayers(1)

        Call SelectCells(0, UBound(ArrShapeID, 1), shObj.Cells(UTR).Result(""), shObj.Cells(UTR).Result(""))

        vs = winObj.Selection
        iAll = shpsObj.Item(NT).Cells(UTR).Result("")
        nRow = shObj.Cells(UTR).Result("")

        Call RecUndo("Добавить строку")
        vs.Duplicate() : vsoDups = winObj.Selection

        For i = 2 To vsoDups.Count
            With vsoDups(i)
                If Not .Characters.IsField Then .Characters.Text = ""
                .Cells("Comment").FormulaForceU = "=Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"
                If InStr(1, .Cells(HE).FormulaU, "SUM") <> 0 Then
                    .Cells(LD).FormulaForceU = 0
                    .Delete()
                End If
            End With
        Next

        NTNew = vsoDups(1).Name

        If arg = 0 Then ' Вставка строки перед выделенной
            With vs(1)
                .Cells(PY).FormulaForceU = GU & NTNew & "!PinY-(" & NTNew & "!Height/2)-(Height/2))"
            End With
            For i = nRow To UBound(ArrShapeID, 2) ' Перенумерация управляющих ячеек
                shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).Result("") + 1 & ")"
            Next

        ElseIf arg = 1 Then ' Вставка строки после выделенной
            With vsoDups(1)
                .Cells(PY).FormulaForceU = GU & vs(1).Name & "!PinY-(" & vs(1).Name & "!Height/2)-(Height/2))"
                .Cells(UTR).FormulaForceU = GU & .Cells(UTR).Result("") + 1 & ")"
            End With
            If nRow <> shpsObj.Item(NT).Cells(UTR).Result("") Then
                shpsObj.ItemFromID(ArrShapeID(0, vs(1).Cells(UTR).Result("") + 1)).Cells(PY).FormulaForceU = GU & vsoDups(1).Name & "!PinY-(" & vsoDups(1).Name & "!Height/2)-(Height/2))"
                For i = nRow + 1 To UBound(ArrShapeID, 2) ' Перенумерация управляющих ячеек
                    shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(UTR).Result("") + 1 & ")"
                Next
            End If
        End If

        With shpsObj
            .Item(NT).Cells(UTR).FormulaForceU = "GUARD(" & iAll + 1 & ")"
            .Item(NTNew).Cells("Controls.ControlHeight").FormulaForceU = "Guard(Height*0)"
            .Item(NTNew).SendToBack()

            If vsoDups.Count <> iAll + 1 Then ' Определение объединенных ячеек и их обработка
                For j = 1 To UBound(ArrShapeID, 1)
                    For i = 1 To UBound(ArrShapeID, 2)
                        If ArrShapeID(j, i) <> 0 Then
                            With .ItemFromID(ArrShapeID(j, i))
                                If InStr(1, .Cells(HE).FormulaU, "SUM", 1) <> 0 Then
                                    If InStr(1, .Cells(HE).FormulaU, vs(1).Name & "!", 1) <> 0 Then
                                        If arg = 0 Then
                                            .Cells(HE).FormulaForceU = Replace$(.Cells(HE).FormulaU, vs(1).Name & "!Height", NTNew & "!Height" & "," & vs(1).Name & "!Height", 1)
                                            strF = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!PinY+", NTNew & "!PinY+", 1)
                                            strF = Replace$(strF, vs(1).Name & "!Height/2", NTNew & "!Height/2", 1)
                                            .Cells(PY).FormulaForceU = Replace$(strF, vs(1).Name & "!Height", NTNew & "!Height" & "," & vs(1).Name & "!Height", 1)
                                            .Cells(UTR).FormulaForceU = Replace$(.Cells(UTR).FormulaU, vs(1).Name & "!", NTNew & "!", 1)
                                        End If
                                        If arg = 1 Then
                                            .Cells(HE).FormulaForceU = Replace$(.Cells(HE).FormulaU, vs(1).Name & "!Height", vs(1).Name & "!Height" & "," & NTNew & "!Height", 1)
                                            .Cells(PY).FormulaForceU = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!Height,", vs(1).Name & "!Height" & "," & NTNew & "!Height,", 1)
                                            .Cells(PY).FormulaForceU = Replace$(.Cells(PY).FormulaU, vs(1).Name & "!Height)", vs(1).Name & "!Height" & "," & NTNew & "!Height)", 1)
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    Next
                Next
            End If
        End With

        Call RecUndo("0")
        On Error Resume Next
        winObj.Selection = vsoDups
        Call PropLayers(0)

    End Sub

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

    Sub CallHelp() ' Вызов справочного файла - Таблицы в Visio.chm
        Dim RetVal, strPath As String
        strPath = "C:\Windows\hh.exe " & vsoApp.MyShapesPath & "\" & "Таблицы в Visio.chm"
        RetVal = Shell(strPath, 1)
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

    Sub CopyT() ' Копирование содержимого выделенных ячеек таблицы
        If Not CheckSelCells() Then Exit Sub

        Dim txt As String = ""
        My.Computer.Clipboard.SetText(fArrT(txt))
    End Sub

    Sub DelColRows(bytColsOrRows As Byte) ' Удаление столбцов/строк из активной таблицы из шейпа. Предварительная процедура
        If Not CheckSelCells() Then Exit Sub

        shpsObj = winObj.Page.Shapes
        Dim vsoSel As Visio.Selection = winObj.Selection, shObj As Visio.Shape

    For Each shObj In vsoSel
            If InStr(1, shObj.Name, "Sheet", 1) = 0 Then
                Select Case bytColsOrRows
                    Case 0
                        If shpsObj.Item(NT).Cells(UTC).Result("") = 1 Then GoTo err
                        Call InitArrShapeID(NT) : Call DeleteColumn(shObj)
                    Case 1
                        If shpsObj.Item(NT).Cells(UTR).Result("") = 1 Then GoTo err
                        Call InitArrShapeID(NT) : Call DeleteRow(shObj)
                End Select
            End If
        Next
err:
        vsoSel = Nothing
    End Sub

    Sub DelTab() ' Удаление активной таблицы. Основная процедура
        On Error GoTo errD
        Dim Response As Byte = 0
        ' 6 - Да, 7 - нет, 2 - отмена
        If Not CheckSelCells() Then Exit Sub

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

    Sub GutT() ' Вырезание содержимого из выделенных ячеек таблицы
        If Not CheckSelCells() Then Exit Sub

        Dim txt As String = ""
        My.Computer.Clipboard.SetText(fArrT(txt))

        Call RecUndo("Вырезать текст из ячеек")

        Dim vsoSelection As Visio.Selection = winObj.Selection
        For Each shpObj In vsoSelection
            shpObj.Characters.Text = ""
        Next

        Call RecUndo("0")
    End Sub

    Sub IntDeIntCells() ' Объединение/Разъединение ячеек из шейпа. Предварительная процедура
        If Not CheckSelCells() Then Exit Sub
        Call ClearControlCells(UTC) : Call ClearControlCells(UTR)

        If Not CheckSelCells() Then Exit Sub

        Dim shObj As Visio.Shape, vsoSel As Visio.Selection = winObj.Selection
        shpsObj = winObj.Page.Shapes
        shObj = vsoSel(1)
        Call InitArrShapeID(NT)

        If InStr(1, shObj.Cells("Width").FormulaU, "SUM", 1) <> 0 Or InStr(1, shObj.Cells("Height").FormulaU, "Sum", 1) <> 0 Then
            Call RecUndo("Разъединить ячейки")
            Call DeIntegrateCells(shObj)
        Else
            Call RecUndo("Объединить ячейки")
            Call IntegrateCells()
        End If

        Call RecUndo("0")
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

    Sub LockPicture(hAL, Val, shN, lF) ' Закрепление изображений в ячейках таблицы
        ' hAL - выравнивание по горизонтали(1-3), vAL - выравнивание по вертикали(1-3)
        ' shN - помещать названия(0,1),lF - блокировать формулы(True,False)
        Dim vsoSel As Visio.Selection = winObj.Selection
        Dim shpObj As Visio.Shape, shpObj1 As Visio.Shape
        Dim strH As String = "", strV As String = "", strL As String = "", strL1 As String = ""
        Dim cnt As Integer, intDot As Integer
        Dim resvar As Byte

        shpsObj = winObj.Page.Shapes

        Select Case hAL
            Case 1
                strH = "Width*0"
            Case 2
                strH = "Width*0.5"
            Case 3
                strH = "Width*1"
        End Select

        Select Case Val
            Case 1
                strV = "Height*1"
            Case 2
                strV = "Height*0.5"
            Case 3
                strV = "Height*0"
        End Select

        Select Case lF
            Case True
                strL = "Guard(" : strL1 = ")"
            Case False
                strL = "" : strL1 = ""
        End Select

        Call RecUndo("Закрепить изображения")

        For Each shpObj In vsoSel
            If shpObj.CellExistsU(UTN, 0) Then
                For Each shpObj1 In shpsObj
                    If Not shpObj1.CellExistsU(UTN, 0) Then
                        resvar = shpObj1.SpatialRelation(shpObj, 0, 10)
                        If resvar = 4 Then
                            shpObj1.Cells("LocPinX").FormulaForceU = strH
                            shpObj1.Cells("LocPinY").FormulaForceU = strV

                            Select Case hAL
                                Case 1 'X слева
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX+" & shpObj.Name & "!LeftMargin-" & shpObj.Name & "!Width/2" & strL1
                                Case 2 'X по центру
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX" & strL1
                                Case 3 'X справа
                                    shpObj1.Cells(PX).FormulaForceU = strL & shpObj.Name & "!PinX-" & shpObj.Name & "!RightMargin+" & shpObj.Name & "!Width*0.5" & strL1
                            End Select

                            Select Case Val
                                Case 1 'Y сверху
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY-" & shpObj.Name & "!TopMargin+" & shpObj.Name & "!Height/2" & strL1
                                Case 2 'Y по центру
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY" & strL1
                                Case 3 'Y снизу
                                    shpObj1.Cells(PY).FormulaForceU = strL & shpObj.Name & "!PinY+" & shpObj.Name & "!BottomMargin-" & shpObj.Name & "!Height*0.5" & strL1
                            End Select

                            If shN <> 0 Then
                                intDot = InStr(1, shpObj1.Name, ".")
                                If intDot <> 0 Then
                                    shpObj.Characters.Text = Left$(shpObj1.Name, intDot - 1)
                                Else
                                    shpObj.Characters.Text = shpObj1.Name
                                End If
                            End If
                            cnt = cnt + 1
                            Exit For
                        End If
                    End If
                Next
            End If
        Next

        Call RecUndo("0")

        ' Результаты метода SpatialRelation (resvar)
        ' 4 - фигура внутри фигуры
        ' 1 - фигуры перекрываются
        ' 8 - фигуры соприкасаются
        ' 0 - фигуры не имеют равных точек
        MsgBox("Готово." & vbCrLf & "Закреплено " & cnt & " фигур в ячейках.")

    End Sub

    Sub PasteT() ' Вставка содержимого буфера обмена в ячейки таблицы
        If Not CheckSelCells() Then Exit Sub

        shpsObj = winObj.Page.Shapes

        Call InitArrShapeID(NT)

        Dim ShapeObj As Visio.Shape
        Dim arrId(,) As String, arrTMP() As String, arrTMP1() As String, txt As String
        Dim i As Integer, j As Integer

        On Error GoTo err

        txt = My.Computer.Clipboard.GetText
        arrTMP = Split(txt, vbCrLf)
        arrTMP1 = Split(arrTMP(0), vbTab)

        ReDim arrId(UBound(arrTMP, 1) - 1, UBound(arrTMP1, 1))
        For i = LBound(arrId, 1) To UBound(arrId, 1)
            arrTMP1 = Split(arrTMP(i), vbTab)
            For j = LBound(arrTMP1, 1) To UBound(arrTMP1, 1)
                arrId(i, j) = arrTMP1(j)
            Next
        Next

        ShapeObj = winObj.Selection(1) : shpsObj = winObj.Page.Shapes

        On Error Resume Next

        Call RecUndo("Вставить текст в ячейки")

        For i = LBound(arrId, 1) To UBound(arrId, 1)
            For j = LBound(arrId, 2) To UBound(arrId, 2)
                With shpsObj.ItemFromID(ArrShapeID(j + ShapeObj.Cells(UTC).Result(""), i + ShapeObj.Cells(UTR).Result("")))
                    .Characters.Text = arrId(i, j)
                End With
            Next
        Next

err:
        Call RecUndo("0")
        Erase arrId : Erase arrTMP : Erase arrTMP1
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

    Private Sub ClearControlCells(arg)   ' Deselect УЯ столбцов или строк
        Dim shObj As Visio.Shape

        With winObj
            For Each shObj In .Selection
                If shObj.Cells(arg).Result("") = 0 Then .Select(shObj, 1)
                If shObj.Name = NT Then .Select(shObj, 1)
            Next
        End With

    End Sub

    Private Sub IntegrateCells() ' Объединение выделенных ячеек в одну с сохранением содержимого. Основная процедура

        Dim vsoSel As Visio.Selection = winObj.Selection
        Dim shObj As Visio.Shape, flagCheck As Boolean
        flagCheck = True

        If vsoSel.Count < 2 Then
            MsgBox("Должно быть выделено не меньше двух ячеек:" & vbCrLf & "Первая и последняя в предполагаемом диапазоне объединения. Или все объединяемые", 48, "Ошибка!")
            Exit Sub
        End If
        '------------------------------- START --------------------------------------------------------
        Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer, NText As String
        Dim Matr As Integer, i As Integer, fWn As String, fHn As String, fXn As String, fYn As String, shObj1 As Visio.Shape

        If Not GetMinMaxRange(vsoSel, cMin, cMax, rMin, rMax) Then GoTo err
        shObj1 = shpsObj.ItemFromID(GetShapeId(cMin, rMin))
        NText = shObj1.Characters.Text

        winObj.DeselectAll()
        Call SelectCells(cMin, cMax, rMin, rMax)
        vsoSel = winObj.Selection

        ' Start Проверка на вшивость------------------------------------------------
        Matr = (cMax - cMin + 1) * (rMax - rMin + 1)
        If vsoSel.Count <> Matr Then flagCheck = False
        For i = 1 To vsoSel.Count
            If InStr(1, winObj.Selection(i).Cells(WI).FormulaU, "SUM", 1) <> 0 Or InStr(1, winObj.Selection(i).Cells(HE).FormulaU, "SUM", 1) <> 0 Then flagCheck = False
        Next
        If flagCheck = False Then GoTo err
        ' End Проверка на вшивость -------------------------------------------------

        winObj.DeselectAll()

        'Start Генерация  и переопределение формул для объединенной ячейки: PinX, PinY, Width, Height
        If cMax - cMin <> 0 Then
            fWn = "=GUARD(SUM("
            fXn = "=GUARD(Sheet." & ArrShapeID(cMin, 0) & "!PinX-(Sheet." & ArrShapeID(cMin, 0) & "!Width/2)+SUM("
            For i = cMin To cMax
                shObj = shpsObj.ItemFromID(ArrShapeID(i, 0))
                fWn = fWn & shObj.Name & "!Width,"
                fXn = fXn & shObj.Name & "!Width,"
            Next
            fWn = Left$(fWn, Len(fWn) - 1) & "))"
            fXn = Left$(fXn, Len(fXn) - 1) & ")/2)"
        Else
            fWn = GU & shpsObj.ItemFromID(ArrShapeID(cMin, 0)).Name & "!Width)"
            fXn = GU & shpsObj.ItemFromID(ArrShapeID(cMin, 0)).Name & "!PinX)"
        End If

        '---------------------------------------------------------------------------
        If rMax - rMin <> 0 Then
            fHn = "=GUARD(SUM("
            fYn = "=GUARD(Sheet." & ArrShapeID(0, rMin) & "!PinY+(Sheet." & ArrShapeID(0, rMin) & "!Height/2)-SUM("
            For i = rMin To rMax
                shObj = shpsObj.ItemFromID(ArrShapeID(0, i))
                fHn = fHn & shObj.Name & "!Height,"
                fYn = fYn & shObj.Name & "!Height,"
            Next
            fHn = Left$(fHn, Len(fHn) - 1) & "))"
            fYn = Left$(fYn, Len(fYn) - 1) & ")/2)"
        Else
            fHn = GU & shpsObj.ItemFromID(ArrShapeID(0, rMin)).Name & "!Height)"
            fYn = GU & shpsObj.ItemFromID(ArrShapeID(0, rMin)).Name & "!PinY)"
        End If

        '---------------------------------------------------------------------------
        With shObj1 ' переопределение формул ячейки
            .Cells(PX).FormulaForceU = fXn
            .Cells(PY).FormulaForceU = fYn
            .Cells(WI).FormulaForceU = fWn
            .Cells(HE).FormulaForceU = fHn
            .BringToFront()
            .Characters.Text = NText
        End With
        'End переопределение формул =============================================

        For i = 2 To vsoSel.Count ' Удаление мусорных ячеек
            vsoSel(i).Cells(LD).FormulaForceU = 0
            vsoSel(i).Delete()
        Next

        winObj.Select(shObj1, 2)
        Exit Sub

err:
        Dim msg As String
        msg = "Возможные причины ошибки:" & vbCrLf
        msg = msg & "Выделена уже объединенная ячейка." & vbCrLf
        msg = msg & "Что-то пошло не так." & vbCrLf
        MsgBox(msg, 48, "Ошибка!")

    End Sub

    Private Sub DeIntegrateCells(shObj As Visio.Shape) ' Разъединение выделенной ячейки с сохранением содержимого. Основная процедура

        Dim flagCheck As Boolean, flagTxt As Boolean
        flagCheck = True

        If winObj.Selection.Count <> 1 Then
            MsgBox("Должна быть выделена одна ячейка:", 48, "Ошибка!")
            Exit Sub
        End If

        If InStr(1, shObj.Cells(WI).FormulaU, "SUM", 1) = 0 And InStr(1, shObj.Cells(HE).FormulaU, "SUM", 1) = 0 Then flagCheck = False
        If Not flagCheck Then GoTo err

        '------------------------------- START --------------------------------------------------------
        Dim vsoDup As Visio.Shape
        Dim fx As String, fy As String, arrX() As String, arrY() As String, NText As String
        Dim j As Integer, i As Integer

        With shObj
            fx = .Cells(PX).FormulaU : fy = .Cells(PY).FormulaU
            NText = .Characters.Text
        End With

        '---------------------------------------------------------------
        If InStr(1, fx, "SUM", 1) <> 0 Then
            fx = Left$(fx, Len(fx) - 4)
            fx = Right$(fx, Len(fx) - InStr(1, fx, "+") - 4)
            fx = Replace$(fx, WI, "", 1)
            arrX = Split(fx, ",")
        Else
            fx = Replace$(fx, "GUARD(", "", 1)
            ReDim arrX(0)
            arrX(0) = Replace$(fx, "PinX)", "", 1)
        End If

        '---------------------------------------------------------------
        If InStr(1, fy, "SUM", 1) <> 0 Then
            fy = Left$(fy, Len(fy) - 4)
            fy = Right$(fy, Len(fy) - InStr(1, fy, "-") - 4)
            fy = Replace$(fy, HE, "", 1)
            arrY = Split(fy, ",")
        Else
            fy = Replace$(fy, "GUARD(", "", 1)
            ReDim arrY(0)
            arrY(0) = Replace$(fy, "PinY)", "", 1)
        End If

        '---------------------------------------------------------------
        flagTxt = True

        shObj.Characters.Text = NText

        For j = 0 To UBound(arrY)
            For i = 0 To UBound(arrX)
                vsoDup = shObj.Duplicate
                With vsoDup ' переопределение формул ячейки
                    .Cells(PX).FormulaForceU = GU & arrX(i) & "PinX)"
                    .Cells(PY).FormulaForceU = GU & arrY(j) & "PinY)"
                    .Cells(WI).FormulaForceU = GU & arrX(i) & "Width)"
                    .Cells(HE).FormulaForceU = GU & arrY(j) & "Height)"
                    .Cells(UTN).FormulaForceU = shObj.Cells(UTN).FormulaU
                    .Cells(UTC).FormulaForceU = GU & arrX(i) & "User.TableCol)"
                    .Cells(UTR).FormulaForceU = GU & arrY(j) & "User.TableRow)"
                    If j <> 0 Or i <> 0 Then .Cells("Comment").FormulaForceU = "Guard(IF(" & NT & "!Actions.Comments.Checked=1," & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"

                    If flagTxt Then
                        .Characters.Text = NText
                        flagTxt = False
                    Else
                        .Characters.Text = ""
                    End If
                End With

            Next
        Next
        shObj.Cells(LD).FormulaForceU = 0
        shObj.Delete()

        Exit Sub

err:
        Dim msg As String
        msg = "Возможные причины ошибки:" & vbCrLf & vbCrLf
        msg = msg & "1. Выделено больше одной ячейки" & vbCrLf
        msg = msg & "2. Выделена не объединенная ячейка" & vbCrLf
        MsgBox(msg, 48, "Ошибка!")

    End Sub

    Private Sub DeleteColumn(shObj) ' Удаление столбца. Основная процедура
        If shObj.Cells(UTC).Result("") = 0 Or shObj.Cells(UTN).FormulaU = "GUARD(NAME(0))" Then Exit Sub
        Call RecUndo("Удалить столбец")

        Dim iAll As Integer, iDel As Integer, i As Integer, j As Integer
        Dim NTDel As String, strF As String, tmpName As String = "", PropC(1) As String

        iDel = shObj.Cells(UTC).Result("") : iAll = shpsObj.Item(NT).Cells(UTC).Result("") : NTDel = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Name

        If iDel < iAll Then ' Сохранение свойств удаляемой упр. ячейки
            PropC(0) = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Cells(PX).FormulaU
            PropC(1) = shpsObj.ItemFromID(ArrShapeID(iDel, 0)).Cells(PY).FormulaU
        End If

        Call PropLayers(1)

        If iDel <> iAll Then tmpName = shpsObj.ItemFromID(ArrShapeID(iDel + 1, 0)).Name

        With shpsObj ' Определение объединенных ячеек и их обработка
            For j = 1 To UBound(ArrShapeID, 1)
                For i = 1 To UBound(ArrShapeID, 2)
                    With .ItemFromID(ArrShapeID(j, i))
                        If InStr(1, .Cells(WI).FormulaU, "SUM") <> 0 Then
                            If InStr(1, .Cells(WI).FormulaU, NTDel) <> 0 And InStr(1, .Cells(WI).FormulaU, ",") <> 0 Then
                                strF = Replace$(.Cells(WI).FormulaU, NTDel & "!Width", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(WI).FormulaForceU = Replace$(strF, ",,", ",", 1)
                                '----------------------------------------------------------------------------------------------------------------
                                strF = .Cells(PX).FormulaU
                                If iDel <> iAll Then
                                    strF = Replace$(.Cells(PX).FormulaU, NTDel & "!PinX", tmpName & "!PinX", 1)
                                    strF = Replace$(strF, NTDel & "!Width/2", tmpName & "!Width/2", 1)
                                    .Cells(UTC).FormulaForceU = Replace$(.Cells(UTC).FormulaU, NTDel & "!", tmpName & "!", 1)
                                End If
                                strF = Replace$(strF, NTDel & "!Width", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(PX).FormulaForceU = Replace$(strF, ",,", ",", 1)
                            End If
                            If .Cells(WI).Result(64) = shpsObj.ItemFromID(ArrShapeID(j, 0)).Cells(WI).Result(64) Then
                                .Cells(WI).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j, 0)).Name & "!Width)"
                                .Cells(PX).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(j, 0)).Name & "!PinX)"
                            End If
                        End If
                    End With
                Next
            Next

            For i = LBound(ArrShapeID, 2) To UBound(ArrShapeID, 2) 'Удаление выделенных ячеек по критерию
                With .ItemFromID(ArrShapeID(iDel, i))
                    If ArrShapeID(iDel, i) <> 0 Then
                        If InStr(1, .Cells(WI).FormulaU, "SUM", 1) = 0 Then
                            .Cells(LD).FormulaForceU = 0
                            .Delete()
                            ArrShapeID(iDel, i) = 0
                        End If
                    End If
                End With
            Next
        End With

        shpsObj.Item(NT).Cells(UTC).FormulaForceU = "GUARD(" & iAll - 1 & ")"

        If iDel < iAll Then
            With shpsObj.ItemFromID(ArrShapeID(iDel + 1, 0))
                .Cells(PX).FormulaForceU = PropC(0)
                .Cells(PY).FormulaForceU = PropC(1)
            End With

            With shpsObj ' Перенумерование столбцов
                j = 0
                For i = 1 To UBound(ArrShapeID, 1)
                    If ArrShapeID(i, 0) <> 0 Then
                        j = j + 1
                        .ItemFromID(ArrShapeID(i, 0)).Cells(UTC).FormulaForceU = GU & j & ")"
                    End If
                Next
            End With
        End If

        Erase PropC
        Call PropLayers(0)
        winObj.Select(shpsObj.ItemFromID(ArrShapeID(0, 0)), 2)
        Call RecUndo("0")

    End Sub

    Private Sub DeleteRow(shObj) ' Удаление строки. Основная процедура
        If shObj.Cells(UTR).Result("") = 0 Or shObj.Cells(UTN).FormulaU = "GUARD(NAME(0))" Then Exit Sub
        Call RecUndo("Удалить строку")

        Dim iAll As Integer, iDel As Integer, i As Integer, j As Integer
        Dim NTDel As String, strF As String, tmpName As String = "", PropC(1) As String

        iDel = shObj.Cells(UTR).Result("") : iAll = shpsObj.Item(NT).Cells(UTR).Result("") : NTDel = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Name

        If iDel < iAll Then ' Сохранение свойств удаляемой упр. ячейки
            PropC(0) = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Cells(PX).FormulaU
            PropC(1) = shpsObj.ItemFromID(ArrShapeID(0, iDel)).Cells(PY).FormulaU
        End If

        Call PropLayers(1)

        If iDel <> iAll Then tmpName = shpsObj.ItemFromID(ArrShapeID(0, iDel + 1)).Name

        With shpsObj ' Определение объединенных ячеек и их обработка
            For j = 1 To UBound(ArrShapeID, 1)
                For i = 1 To UBound(ArrShapeID, 2)
                    With .ItemFromID(ArrShapeID(j, i))
                        If InStr(1, .Cells(HE).FormulaU, "SUM") <> 0 Then
                            If InStr(1, .Cells(HE).FormulaU, NTDel) <> 0 And InStr(1, .Cells(HE).FormulaU, ",") <> 0 Then
                                strF = Replace$(.Cells(HE).FormulaU, NTDel & "!Height", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(HE).FormulaForceU = Replace$(strF, ",,", ",", 1)
                                '-------------------------------------------------------------------------------------------------------------
                                strF = .Cells(PY).FormulaU
                                If iDel <> iAll Then
                                    strF = Replace$(.Cells(PY).FormulaU, NTDel & "!PinY", tmpName & "!PinY", 1)
                                    strF = Replace$(strF, NTDel & "!Height/2", tmpName & "!Height/2", 1)
                                    .Cells(UTR).FormulaForceU = Replace$(.Cells(UTR).FormulaU, NTDel & "!", tmpName & "!", 1)
                                End If
                                strF = Replace$(strF, NTDel & "!Height", "", 1)
                                strF = Replace$(strF, "(,", "(", 1) : strF = Replace$(strF, ",)", ")", 1)
                                .Cells(PY).FormulaForceU = Replace$(strF, ",,", ",", 1)
                            End If
                            If .Cells(HE).Result(64) = shpsObj.ItemFromID(ArrShapeID(0, i)).Cells(HE).Result(64) Then
                                .Cells(HE).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Name & "!Height)"
                                .Cells(PY).FormulaForceU = GU & shpsObj.ItemFromID(ArrShapeID(0, i)).Name & "!PinY)"
                            End If
                        End If
                    End With
                Next
            Next

            For i = LBound(ArrShapeID, 1) To UBound(ArrShapeID, 1) 'Удаление выделенных ячеек по критерию
                With .ItemFromID(ArrShapeID(i, iDel))
                    If ArrShapeID(i, iDel) <> 0 Then
                        If InStr(1, .Cells(HE).FormulaU, "SUM", 1) = 0 Then
                            .Cells(LD).FormulaForceU = 0
                            .Delete()
                            ArrShapeID(i, iDel) = 0
                        End If
                    End If
                End With
            Next
        End With

        shpsObj.Item(NT).Cells(UTR).FormulaForceU = "GUARD(" & iAll - 1 & ")"

        If iDel < iAll Then
            With shpsObj.ItemFromID(ArrShapeID(0, iDel + 1))
                .Cells(PX).FormulaForceU = PropC(0)
                .Cells(PY).FormulaForceU = PropC(1)
            End With

            With shpsObj ' Перенумерование строк
                j = 0
                For i = 1 To UBound(ArrShapeID, 2)
                    If ArrShapeID(0, i) <> 0 Then
                        j = j + 1
                        .ItemFromID(ArrShapeID(0, i)).Cells(UTR).FormulaForceU = GU & j & ")"
                    End If
                Next
            End With
        End If

        Erase PropC
        Call PropLayers(0)
        winObj.Select(shpsObj.ItemFromID(ArrShapeID(0, 0)), 2)
        Call RecUndo("0")

    End Sub

    Private Sub NotDub(vsoSel, UT) ' Заполнение коллекции значениями без дубликатов
        Dim Shp As Visio.Shape

        On Error Resume Next

        For Each Shp In vsoSel
            NoDupes.Add(Shp.Cells(UT).Result(""), CStr(Shp.Cells(UT).Result("")))
        Next

        On Error GoTo 0

    End Sub

    Private Sub PropLayers(arg As Byte) ' Включение/выключение видимости и блокировки слоев на время выполнения кода - Titles_Tables и Cells_Tables

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

    Private Sub SelectCells(intStartCol As Integer, intEndCol As Integer, intStartRow As Integer, intEndRow As Integer) 'Различное выделение в таблице

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

#Region "Functions"

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

    Private Function fArrT(txt) ' Заполнение массива данными из ячеек таблицы
        Dim i As Integer, j As Integer, arrId(,) As String, Response As Boolean
        Dim cMin As Integer, rMin As Integer, cMax As Integer, rMax As Integer

        Call ClearControlCells(UTC) : Call ClearControlCells(UTR)

        Dim vsoSelection As Visio.Selection = winObj.Selection

        Response = GetMinMaxRange(vsoSelection, cMin, cMax, rMin, rMax)

        ReDim arrId(rMax + 1, cMax + 1)

        For i = 1 To vsoSelection.Count
            With vsoSelection(i)
                arrId(.Cells(UTR).Result(""), .Cells(UTC).Result("")) = .Characters.Text
            End With
        Next

        For j = rMin To rMax
            For i = cMin To cMax
                txt = IIf(i = cMax, txt & arrId(j, i) & vbCrLf, txt & arrId(j, i) & vbTab)
            Next
        Next

        Erase arrId
        fArrT = txt
    End Function

    Private Function GetMinMaxRange(ByVal vsoSel As Visio.Selection, ByRef cMin As Integer, ByRef cMax As Integer, ByRef rMin As Integer, ByRef rMax As Integer) As Boolean
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

    Private Function GetShapeId(ByVal intColNum As Integer, ByVal intRowNum As Integer) As Integer
        ' Получение ID ячейки таблицы по номеру столбца и строки

        On Error GoTo err

        If ArrShapeID(intColNum, intRowNum) <> 0 Then
            GetShapeId = ArrShapeID(intColNum, intRowNum)
        Else
            Dim i As Integer, j As Integer, cN As String, rN As String
            With winObj.Page.Shapes
                cN = .ItemFromID(ArrShapeID(intColNum, 0)).Name : rN = .ItemFromID(ArrShapeID(0, intRowNum)).Name
                For i = 1 To intColNum
                    For j = 1 To intRowNum
                        If ArrShapeID(i, j) <> 0 Then
                            If InStr(1, .ItemFromID(ArrShapeID(i, j)).Cells(PX).FormulaU, cN) <> 0 And _
                               InStr(1, .ItemFromID(ArrShapeID(i, j)).Cells(PY).FormulaU, rN) <> 0 Then
                                GetShapeId = ArrShapeID(i, j)
                                Exit Function
                            End If
                        End If
                    Next
                Next
            End With
        End If
        Exit Function

err:
        GetShapeId = 0
    End Function

#End Region

End Module
