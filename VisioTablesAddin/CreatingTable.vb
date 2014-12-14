Imports System.Drawing
Imports System.Windows.Forms

Module CreatingTable

#Region "LIST OF VARIABLES"

    Dim frmNewTable As System.Windows.Forms.Form = New dlgNewTable
    Public booOpenForm As Boolean = False
    Public strNameTable As String
    Public vsoApp As Visio.Application = Globals.ThisAddIn.Application
    Dim docObj As Visio.Document = vsoApp.ActiveDocument
    Dim winObj As Visio.Window = vsoApp.ActiveWindow
    Dim pagObj As Visio.Page = vsoApp.ActivePage
    Dim shpsObj As Visio.Shapes = pagObj.Shapes
    Dim selObj As Visio.Selection
    Dim shapeTb, shapeTh, shapeTv, shapeCl As Visio.Shape
    Dim shpObj As Visio.Shape
    Dim MemSel As Visio.Shape
    Dim vsoSelection As Visio.Selection
    Dim shape_TbL As Visio.Shape, shape_ThC As Visio.Shape, shape_TvR As Visio.Shape, shape_ClW As Visio.Shape
    Dim CountID As Integer
    Const UTN = "User.TableName", UTR = "User.TableRow", UTC = "User.TableCol", PX = "PinX", PY = "PinY"
    Const WI = "Width", HE = "Height", CA = "Angle", LD = "LockDelete", GU = "=GUARD("
    Dim NT As String = "", UndoScopeID As Long = 0
    Dim arrShapeID() As Integer
    Public FlagPage As Byte
    Dim LayerVisible As String, LayerLock As String
    Const strATC = "!Actions.Titles.Checked=1,"
    Const strACC = "!Actions.Comments.Checked=1,"
    Const strThGu000 = "THEMEGUARD(MSOTINT(RGB(0,0,0),50))"
    Const strThGu255 = "THEMEGUARD(RGB(255,255,255))"
    Const strThGu191 = "THEMEGUARD(RGB(191,191,191))"
    Const GS = "=GUARD(Sheet."
    Const GI = "=Guard(IF("
    Const sh = "Sheet."
    Const frm = "###0.0###"
    Const GU5 = "=GUARD(10 mm)" ' Переделать на DrawUn
    Const P50 = "50%"
    Const GT = "GUARD(TRUE)"
    Const G1 = "Guard(1)"
    Dim iGT As Integer = 0, jGT As Integer = 0

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
        'MessageBox.Show("Таблица: " & nC & " x " & nR)
    End Sub

    Sub Load_dlgNewTable()
        If Not booOpenForm Then
            frmNewTable = New dlgNewTable
            frmNewTable.Show()
        End If
    End Sub

#Region "CREATTABLE"

    Public Sub CreatTable(strNameTable, bytInsertType, intColumnsCount, intRowCount, sngWidthCells, sngHeightCells, _
    sngWidthTable, sngHeightTable, booDeleteTargetShape, booVisibleProgressBar)
        On Error GoTo errD
        'Создание новой таблицы на активном листе
        '-------------------------------------------------------------------------------------------
        'Входные аргументы

        'strNameTable As String             ' - Имя таблицы (Любая строка или "")
        'bytInsertType As Byte              ' - Параметры вставки таблицы (1-4)
        'intColumnsCount As Integer         ' - Количество столбцов
        'intRowCount As Integer             ' - Количество строк
        'sngWidthCells As Single            ' - Ширина ячеек
        'sngHeightCells As Single           ' - Высота ячеек
        'sngWidthTable As Single            ' - Ширина таблицы
        'sngHeightTable As Single           ' - Высота таблицы
        'booDeleteTargetShape As Boolean    ' - Удалить целевой шейп (True, False)
        'booVisibleProgressBar As Boolean   ' - Отключение отображения полосы прогресса (True, False)
        ' -------------------------------------------------------------------------------------------
        'Проверка аргументов
        If Trim(strNameTable) = "" Then strNameTable = "TbL"
        If bytInsertType < 0 And bytInsertType > 4 Then bytInsertType = 1
        If intColumnsCount < 0 Or intColumnsCount > 1000 Then intColumnsCount = 5
        If intRowCount < 0 Or intRowCount > 1000 Then intRowCount = 5
        If sngWidthCells = 0 Then sngWidthCells = 20
        If sngHeightCells = 0 Then sngHeightCells = 10
        If sngWidthTable = 0 Then sngWidthTable = 200
        If sngHeightTable = 0 Then sngHeightTable = 100
        ' -------------------------------------------------------------------------------------------

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

        Dim vsoLayerTitles As Visio.Layer, vsoLayerCells As Visio.Layer, MemSHID As Integer
        Dim TypeCell As String, VarCell As Byte, sngTW As Double, sngTH As Double
        Dim sngTX As Double, sngTY As Double
        'Dim i As Integer
        'Dim vsoShape As Visio.Shape
        Dim ThemeC As String, ThemeE As String

        winObj = vsoApp.ActiveWindow
        pagObj = vsoApp.ActivePage

        ThemeC = pagObj.ThemeColors
        If ThemeC <> "visThemeNone" Then pagObj.ThemeColors = "visThemeNone"

        ThemeE = pagObj.ThemeEffects
        If ThemeE <> "visThemeEffectsNone" Then pagObj.ThemeEffects = "visThemeEffectsNone"

        vsoLayerTitles = pagObj.Layers.Add("Titles_Tables")
        vsoLayerCells = pagObj.Layers.Add("Cells_Tables")
        vsoLayerTitles.CellsC(5).FormulaU = "0"

        If bytInsertType = 2 Then
            Dim sngPW As Single, sngPH As Single, sngPLM As Single, sngPRM As Single, sngPTM As Single, sngPBM As Single
            With pagObj
                sngPW = .PageSheet.Cells("PageWidth").Result(64)
                sngPH = .PageSheet.Cells("PageHeight").Result(64)
                sngPLM = .PageSheet.Cells("PageLeftMargin").Result(64)
                sngPRM = .PageSheet.Cells("PageRightMargin").Result(64)
                sngPTM = .PageSheet.Cells("PageTopMargin").Result(64)
                sngPBM = .PageSheet.Cells("PageBottomMargin").Result(64)
                sngTW = (sngPW - sngPLM - sngPRM) / intColumnsCount
                sngTH = (sngPH - sngPTM - sngPBM) / intRowCount
            End With
        End If

        If bytInsertType = 4 Then ' Использовать вместо PinX и PinY - BoundingBox left and top
            If winObj.Selection.Count = 0 Then
                MsgBox("Вы должны выбрать одну фигуру")
                Exit Sub
            Else
                With winObj.Selection(1)
                    .BoundingBox(3, sngTX, sngTY, sngTW, sngTH) ' L, B, R, T
                    MemSHID = .ID
                    sngTX = vsoApp.FormatResult(sngTX, "", 64, "0.0000")  '.Cells(PX).Result(64) - .Cells("LocPinX").Result(64)
                    sngTY = vsoApp.FormatResult(sngTH, "", 64, "0.0000")  '.Cells(PY).Result(64) + .Cells("LocPinY").Result(64)
                    sngTW = .Cells(WI).Result(64) / intColumnsCount
                    sngTH = .Cells(HE).Result(64) / intRowCount
                End With
            End If
        End If

        ReDim arrShapeID((intRowCount * intColumnsCount) + (intRowCount + intColumnsCount))
        CountID = -1

        'If booVisibleProgressBar Then
        '    Load frmWait : frmWait.Caption = "Создание таблицы..." : frmWait.Show() : frmWait.Repaint()
        'End If

        'Call RecUndo("Создание таблицы...")

        vsoApp.ShowChanges = False

        TypeCell = strNameTable : VarCell = 3 'Вставка 1 ячейки
        NewShape(TypeCell)
        shpObj = shape_TbL
        DrawOfCells(VarCell, bytInsertType, intColumnsCount, intRowCount, sngTW, sngTH, sngWidthTable, _
                            sngHeightTable, sngWidthCells, sngHeightCells, sngTX, sngTY, iGT)
        vsoLayerTitles.Add(shpObj, 1)

        TypeCell = "ThC" : VarCell = 2 'Вставка 1 ряда таблицы
        For iGT = 1 To intColumnsCount
            If iGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_ThC
            Else
                shpObj = shape_ThC.Duplicate
            End If
            DrawOfCells(VarCell, bytInsertType, intColumnsCount, intRowCount, sngTW, sngTH, sngWidthTable, _
                            sngHeightTable, sngWidthCells, sngHeightCells, sngTX, sngTY, iGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "TvR" : VarCell = 1 'Вставка 1 столбца таблицы
        For jGT = 1 To intRowCount
            If jGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_TvR
            Else
                shpObj = shape_TvR.Duplicate
            End If
            DrawOfCells(VarCell, bytInsertType, intColumnsCount, intRowCount, sngTW, sngTH, sngWidthTable, _
                            sngHeightTable, sngWidthCells, sngHeightCells, sngTX, sngTY, iGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "ClW" : VarCell = 0 'Вставка рабочих ячеек
        For jGT = 1 To intRowCount
            'If booVisibleProgressBar Then
            '    frmWait.lblProgress.Width = (300 / intRowCount) * jGT : frmWait.Repaint() : DoEvents()
            'End If
            For iGT = 1 To intColumnsCount
                If jGT = 1 And iGT = 1 Then
                    NewShape(TypeCell)
                    shpObj = shape_ClW
                Else
                    shpObj = shape_ClW.Duplicate
                End If
                DrawOfCells(VarCell, bytInsertType, intColumnsCount, intRowCount, sngTW, sngTH, sngWidthTable, _
                            sngHeightTable, sngWidthCells, sngHeightCells, sngTX, sngTY, iGT)
                vsoLayerCells.Add(shpObj, 0)
            Next
        Next

        shpObj = pagObj.Shapes.ItemFromID(arrShapeID(0))
        shpObj.Cells(UTC).FormulaU = "GUARD(" & intColumnsCount & ")"
        shpObj.Cells(UTR).FormulaU = "GUARD(" & intRowCount & ")"

        For iGT = 0 To intColumnsCount + intRowCount - 1
            winObj.Page.Shapes.ItemFromID(arrShapeID(iGT)).Cells("LockTextEdit").FormulaU = "Guard(1)"
        Next

        '        Call RecUndo(0)

        '        If booVisibleProgressBar Then Unload frmWait

        If booDeleteTargetShape Then
            If pagObj.Shapes.ItemFromID(MemSHID).Cells(LD).Result("") = 0 Then pagObj.Shapes.ItemFromID(MemSHID).DeleteEx(0)
        End If

        vsoApp.ShowChanges = True
        'winObj.DeselectAll()
        winObj.Select(shpObj, 258)
        'winObj = Nothing : pagObj = Nothing : shpObj = Nothing
        'shape_TbL = Nothing : shape_ThC = Nothing : shape_TvR = Nothing : shape_ClW = Nothing
        'vsoLayerTitles = Nothing : vsoLayerCells = Nothing
        Erase arrShapeID

        If ThemeC <> pagObj.ThemeColors Then pagObj.ThemeColors = ThemeC
        If ThemeE <> pagObj.ThemeEffects Then pagObj.ThemeEffects = ThemeE
        Exit Sub
errD:
        MessageBox.Show("CreatTable" & vbNewLine & Err.Description)
    End Sub

    Private Sub DrawOfCells(VarCell, bytInsertType, intColumnsCount, intRowCount, sngTW, sngTH, sngWidthTable, _
                            sngHeightTable, sngWidthCells, sngHeightCells, sngTX, sngTY, iGT)
        On Error GoTo errD
        With pagObj
            CountID = CountID + 1
            arrShapeID(CountID) = shpObj.ID
            With shpObj
                Select Case VarCell

                    Case 0 'Вставка рабочих ячеек
                        .Cells(PX).FormulaForceU = GS & arrShapeID(iGT) & "!PinX)"
                        .Cells(PY).FormulaForceU = GS & arrShapeID(intColumnsCount + jGT) & "!PinY)"
                        .Cells(WI).FormulaForceU = GS & arrShapeID(iGT) & "!Width)"
                        .Cells(HE).FormulaForceU = GS & arrShapeID(intColumnsCount + jGT) & "!Height)"
                        .Cells(UTN).FormulaForceU = GS & arrShapeID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GS & arrShapeID(iGT) & "!User.TableCol)"
                        .Cells(UTR).FormulaForceU = GS & arrShapeID(intColumnsCount + jGT) & "!User.TableRow)"
                        .Cells("Comment").FormulaForceU = GI & sh & arrShapeID(0) & strACC & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"

                    Case 2 ' упр строка
                        .Cells(PX).FormulaForceU = GS & arrShapeID(CountID - 1) & "!PinX+(Sheet." & arrShapeID(CountID - 1) & "!Width/2)+(Width/2))"
                        .Cells(PY).FormulaForceU = GS & arrShapeID(0) & "!PinY)"
                        .Cells(HE).FormulaForceU = GS & arrShapeID(0) & "!Height)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(WI).Result(64) = sngTW
                        ElseIf bytInsertType = 3 Then
                            .Cells(WI).Result(64) = sngWidthTable / intColumnsCount
                        Else
                            .Cells(WI).Result(64) = sngWidthCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrShapeID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GU & iGT & ")"
                        .Cells(UTR).FormulaForceU = GU & 0 & ")"
                        .Characters.AddCustomFieldU(UTC, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTC & ")"
                        .Cells("Char.Color").FormulaForceU = strThGu255
                        .Cells("FillForegnd").FormulaForceU = GI & sh & arrShapeID(0) & strATC & strThGu000 & "," & strThGu255 & "))"
                        .Cells("FillForegndTrans").FormulaForceU = GI & sh & arrShapeID(0) & strATC & "0%" & "," & "50%" & "))"
                        .Cells("LineColor").FormulaForceU = GI & sh & arrShapeID(0) & strATC & strThGu191 & "," & strThGu255 & "))"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrShapeID(0) & "!Actions.Titles.Checked))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrShapeID(0) & strACC & """Управляющая ячейка столбца""" & "," & """""" & "))"

                    Case 1 ' упр столбец
                        .Cells(PX).FormulaU = GS & arrShapeID(0) & "!PinX)"
                        If jGT = 1 Then
                            .Cells(PY).FormulaU = GS & arrShapeID(CountID - intColumnsCount - 1) & "!PinY-(Sheet." & arrShapeID(CountID - intColumnsCount - 1) & "!Height/2)-(Height/2))"
                        Else
                            .Cells(PY).FormulaForceU = GS & arrShapeID(CountID - 1) & "!PinY-(Sheet." & arrShapeID(CountID - 1) & "!Height/2)-(Height/2))"
                        End If
                        .Cells(WI).FormulaU = GS & arrShapeID(0) & "!Width)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(HE).Result(64) = sngTH
                        ElseIf bytInsertType = 3 Then
                            .Cells(HE).Result(64) = sngHeightTable / intRowCount
                        Else
                            .Cells(HE).Result(64) = sngHeightCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrShapeID(0) & "!Name(0))"
                        .Cells(UTC).FormulaU = GU & 0 & ")"
                        .Cells(UTR).FormulaForceU = GU & jGT & ")"
                        .Characters.AddCustomFieldU(UTR, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTR & ")"
                        .Cells("Char.Color").FormulaForceU = strThGu255
                        .Cells("FillForegnd").FormulaForceU = GI & sh & arrShapeID(0) & strATC & strThGu000 & "," & strThGu255 & "))"
                        .Cells("FillForegndTrans").FormulaForceU = GI & sh & arrShapeID(0) & strATC & "0%" & "," & "50%" & "))"
                        .Cells("LineColor").FormulaForceU = GI & sh & arrShapeID(0) & strATC & strThGu191 & "," & strThGu255 & "))"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrShapeID(0) & "!Actions.Titles.Checked))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrShapeID(0) & strACC & """Управляющая ячейка строки""" & "," & """""" & "))"

                    Case 3 ' 1 ГлавУпр
                        .Cells(WI).FormulaForceU = GU & Str(vsoApp.FormatResult(10, "mm", "", frm)) & ")"
                        .Cells(HE).FormulaForceU = GU & Str(vsoApp.FormatResult(10, "mm", "", frm)) & ")"
                        If bytInsertType = 4 Then
                            .Cells(PX).Result(64) = sngTX + (.Cells(WI).Result(64) / 2) - .Cells(WI).Result(64)
                            .Cells(PY).Result(64) = sngTY - (.Cells(HE).Result(64) / 2) + .Cells(HE).Result(64)
                        Else
                            .Cells(PX).FormulaU = "=ThePage!PageLeftMargin-5 mm"
                            .Cells(PY).FormulaU = "=ThePage!PageHeight-ThePage!PageTopMargin+5 mm"
                        End If
                        .UpdateAlignmentBox()
                        .Cells(UTN).FormulaU = "=GUARD(Name(0))"
                        .Cells("Char.Color").FormulaForceU = strThGu255
                End Select

                'Для всех ячеек----------------------------------------------------------------
                .Cells(CA).FormulaU = GU & "0 deg)"
            End With
        End With
        Exit Sub
errD:
        MessageBox.Show("DrawOfCells" & vbNewLine & Err.Description)
    End Sub

    Private Sub NewShape(TypeCell)
        On Error GoTo errD
        ' Подпроцедура создания шейпов таблицы и настройка их
        Dim vsoShape As Visio.Shape
        Dim AddSectionNum As Integer, intArrNum() As Integer, arrRowData
        vsoShape = winObj.Page.DrawRectangle(0, 0, 1, 1)

        With vsoShape
            .Name = TypeCell

            ' Добавить User секцию для всех ячеек
            AddSectionNum = 242
            intArrNum = {0, 1}
            'arrRowData =Array("TableName", "Name(0)", """Таблица"""), _
            '    Array("TableCol", """""", """Столбец"""), _
            '    Array("TableRow", """""", """Строка"""))
            arrRowData = {{"TableName", "Name(0)", """Таблица"""}, _
                          {"TableCol", """""", """Столбец"""}, _
                          {"TableRow", """""", """Строка"""}}
            AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)
            '------------------------------------------------------------------

            .Cells("LocPinX").FormulaU = "Guard(Width*0.5)"
            .Cells("LocPinY").FormulaU = "Guard(Height*0.5)"
            .Cells("LinePattern").FormulaU = "1"
            .Cells("LineWeight").FormulaU = "0.1 pt"
            .Cells("Rounding").FormulaU = "Guard(0 mm)"
            .Cells("UpdateAlignBox").FormulaForceU = GT
            .Cells("LockDelete").FormulaU = G1
            .Cells("LockRotate").FormulaU = G1
            .CellsSRC(1, 16, 1).FormulaU = "char(169)&char(32)&char(82)&char(79)&char(77)&char(65)&char(78)&char(79)&char(86)&char(32)&char(86)&char(54)&char(46)&char(48)"
            Select Case TypeCell
                Case strNameTable, "ThC", "TvR"
                    ' Настройка форматов
                    .Cells("LockFormat").FormulaU = G1
                    .Cells("LockFromGroupFormat").FormulaU = G1
                    .Cells("LockThemeColors").FormulaU = G1
                    .Cells("LockThemeEffects").FormulaU = G1
                    ' Настройка Miscellaneous
                    .Cells("NoObjHandles").FormulaForceU = GT
                    .Cells("NonPrinting").FormulaForceU = GT
            End Select

            Select Case TypeCell
                Case "ClW" ' Рабочая ячейка
                    AddSectionNum = 240 ' Добавить Action секцию
                    intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    arrRowData = {{"SelectWindow", "RUNMACRO(""MainModule.LoadSelectfrmWorks"",""Таблица_V"")", """Выде&лить...""", 2035, 1, 0, "FALSE", "FALSE"}, _
                        {"InsertWindow", "RUNMACRO(""MainModule.LoadInsertfrmWorks"",""Таблица_V"")", """Вставить...""", 22, 2, 0, "FALSE", "FALSE"}, _
                        {"SizeWindow", "RUNMACRO(""MainModule.LoadSizefrmWorks"",""Таблица_V"")", """Разм&еры...""", 1605, 3, 0, "FALSE", "TRUE"}, _
                        {"AutoSizeWindow", "RUNMACRO(""MainModule.LoadAutoSizefrmWorks"",""Таблица_V"")", """Авторазмер...""", 583, 4, 0, "FALSE", "FALSE"}, _
                        {"MoreWindow", "RUNMACRO(""MainModule.LoadMorefrmWorks"",""Таблица_V"")", """Дополнительно...""", 586, 5, 0, "FALSE", "FALSE"}, _
                        {"GutT", "RUNMACRO(""MainModule.GutT"",""Таблица_V"")", """Вырезать текст из ячеек""", 2046, 6, 0, "FALSE", "TRUE"}, _
                        {"CopyT", "RUNMACRO(""MainModule.CopyT"",""Таблица_V"")", """Копировать текст из ячеек""", 2045, 7, 0, "FALSE", "FALSE"}, _
                        {"PasteT", "RUNMACRO(""MainModule.PasteT"",""Таблица_V"")", """Вставить текст в ячейки""", 22, 8, 0, "FALSE", "FALSE"}, _
                        {"IntCells", "RUNMACRO(""MainModule.IntDeIntCells"",""Таблица_V"")", """Объединить ячейки""", 402, 9, 0, "FALSE", "TRUE"}, _
                        {"DeIntCells", "RUNMACRO(""MainModule.IntDeIntCells"",""Таблица_V"")", """Отменить объединение""", 402, 9, 0, "TRUE", "TRUE"}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)
                    shape_ClW = vsoShape

                Case "TvR" ' УЯ строки
                    AddSectionNum = 240 ' Добавить Action секцию
                    intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    arrRowData = {{"AddRowTop", "RUNMACRO(""MainModule.AddRowBefore"",""Таблица_V"")", """&Добавить строку выше""", 296, 1, 0, "FALSE", "FALSE"}, _
                        {"AddRowBottom", "RUNMACRO(""MainModule.AddRowAfter"",""Таблица_V"")", """&Добавить строку ниже""", 296, 1, 0, "FALSE", "FALSE"}, _
                        {"SelectRow", "RUNMACRO(""MainModule.SelRows"",""Таблица_V"")", """Выде&лить строку""", 801, 2, 0, "FALSE", "FALSE"}, _
                        {"AlignOnTextHeight", "RUNMACRO(""MainModule.OnHeight"",""Таблица_V"")", """По высоте &текста""", 541, 3, 0, "FALSE", "FALSE"}, _
                        {"DeleteROw", "RUNMACRO(""MainModule.DelRows"",""Таблица_V"")", """&Удалить строку""", 292, 4, 0, "FALSE", "True"}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    AddSectionNum = 9 ' Добавить Control секцию
                    intArrNum = {0, 1, 2, 3, 6, 8} ' Сделать не меньше нуля
                    arrRowData = {{"ControlHeight", "GUARD(Width*0)", "Height*0", "GUARD(Controls.ControlHeight)", "GUARD(Controls.ControlHeight.Y)", "False", """Изменение высоты ячейки"""}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .CellsSRC(10, 1, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 2, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 5, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    shape_TvR = vsoShape

                Case "ThC" ' УЯ столбца
                    AddSectionNum = 240 ' Добавить Action секцию
                    intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    arrRowData = {{"AddColumnLeft", "RUNMACRO(""MainModule.AddColBefore"",""Таблица_V"")", """&Добавить столбец слева""", 297, 1, 0, "FALSE", "FALSE"}, _
                        {"AddColumnRight", "RUNMACRO(""MainModule.AddColAfter"",""Таблица_V"")", """&Добавить столбец справа""", 297, 1, 0, "FALSE", "FALSE"}, _
                        {"SelectColumn", "RUNMACRO(""MainModule.SelCols"",""Таблица_V"")", """Выде&лить столбец""", 802, 2, 0, "FALSE", "FALSE"}, _
                        {"AlignOnTextWidth", "RUNMACRO(""MainModule.OnWidth"",""Таблица_V"")", """По ширине &текста""", 542, 3, 0, "FALSE", "FALSE"}, _
                        {"DeleteColumn", "RUNMACRO(""MainModule.DelCols"",""Таблица_V"")", """&Удалить столбец""", 294, 4, 0, "FALSE", "True"}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    AddSectionNum = 9 ' Добавить Control секцию
                    intArrNum = {0, 1, 2, 3, 6, 8} ' Сделать не меньше нуля
                    arrRowData = {{"ControlWidth", "Width*1", "GUARD(Height)", "GUARD(Controls.ControlWidth)", "GUARD(Controls.ControlWidth.Y)", "False", """Изменение ширины ячейки"""}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .CellsSRC(10, 2, 0).FormulaU = "GUARD(Controls.ControlWidth)"
                    .CellsSRC(10, 3, 0).FormulaU = "GUARD(Controls.ControlWidth)"
                    shape_ThC = vsoShape

                Case strNameTable ' Главная УЯ
                    AddSectionNum = 240 ' Добавить Action секцию
                    intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    arrRowData = {{"NewTable", "RUNMACRO(""MainModule.LoadfrmAddTable"",""Таблица_V"")", """Соз&дать таблицу""", 333, 1, 0, "FALSE", "FALSE"}, _
                        {"SelectTable", "RUNMACRO(""MainModule.SelTable"",""Таблица_V"")", """Выде&лить таблицу""", 803, 2, 0, "FALSE", "FALSE"}, _
                        {"LockPicture", "RUNMACRO(""MainModule.LoadfrmPicture"",""Таблица_V"")", """Закрепить &фигуры""", 954, 3, 0, "FALSE", "FALSE"}, _
                        {"LinkData", "RUNMACRO(""MainModule.LoadfrmLinkData"",""Таблица_V"")", """Вне&шние данные""", 4143, 4, 0, "FALSE", "FALSE"}, _
                        {"Titles", "SETF(GetRef(Actions.Titles.Checked),NOT(Actions.Titles.Checked))", """П&оказывать заголовки""", """""", 5, 1, "FALSE", "TRUE"}, _
                        {"Comments", "SETF(GetRef(Actions.Comments.Checked),NOT(Actions.Comments.Checked))", """Показывать коммента&рии""", """""", 6, 1, "FALSE", "FALSE"}, _
                        {"FixingTable", "SETF(GetRef(Actions.FixingTable.Checked),NOT(Actions.FixingTable.Checked))", """Заф&иксировать таблицу""", """""", 7, 0, "FALSE", "FALSE"}, _
                        {"DeleteTable", "RUNMACRO(""MainModule.DelTab"",""Таблица_V"")", """&Удалить таблицу""", 2487, 8, 0, "FALSE", "TRUE"}, _
                        {"Help", "RUNMACRO(""MainModule.CallHelp"",""Таблица_V"")", """Справка""", 3998, 8, 0, "FALSE", "FALSE"}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .Cells("LineColor").FormulaU = "GUARD(IF(Actions.Titles.Checked=1,THEMEGUARD(RGB(191,191,191)),THEMEGUARD(RGB(255,255,255))))"
                    .Cells("FillForegnd").FormulaU = "GUARD(IF(Actions.Titles.Checked=1,THEMEGUARD(MSOTINT(RGB(0,0,0),50)),THEMEGUARD(RGB(255,255,255))))"
                    .Cells("FillForegndTrans").FormulaU = "GUARD(IF(Actions.Titles.Checked=1,0%,50%))"
                    .Cells("LockMoveX").FormulaU = "GUARD(Actions.FixingTable.Checked)"
                    .Cells("LockMoveY").FormulaU = "GUARD(Actions.FixingTable.Checked)"
                    .Cells("Width").FormulaU = GU5
                    .Cells("Height").FormulaU = GU5
                    .Cells("Comment").FormulaU = "GUARD(IF(Actions.Comments.Checked=1," & "User.TableName&CHAR(10)&" & """Основная управляющая ячейка"", """"))"
                    shape_TbL = vsoShape
            End Select
        End With
        vsoShape = Nothing
        Erase intArrNum : Erase arrRowData
        Exit Sub
errD:
        MessageBox.Show("NewShape" & vbNewLine & Err.Description)
    End Sub

    Private Sub AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)
        On Error GoTo errD
        ' Подпроцесс добавления заданной Section, требуемого кол-ва строк в Section и настройка этих строк
        Dim intI As Byte, intJ As Byte

        With vsoShape
            'If .SectionExists(AddSectionNum, 0) Then .DeleteSection(AddSectionNum)
            .AddSection(AddSectionNum)
            For intI = 0 To UBound(arrRowData)
                .AddRow(AddSectionNum, -2, 0)
                .CellsSRC(AddSectionNum, intI, 0).RowNameU = arrRowData(intI, 0)
                For intJ = 0 To UBound(intArrNum)
                    'If AddSectionNum = 9 Then MessageBox.Show(arrRowData(intI, intJ + 1)) ' Удалить потом
                    .CellsSRC(AddSectionNum, intI, intArrNum(intJ)).FormulaU = arrRowData(intI, intJ + 1)
                Next
            Next
        End With
        Exit Sub
errD:
        MessageBox.Show("AddSections" & vbNewLine & Err.Description)
    End Sub

#End Region

End Module
