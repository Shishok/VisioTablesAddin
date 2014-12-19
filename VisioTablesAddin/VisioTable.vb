Imports System.Drawing
Imports System.Windows.Forms

Public Class VisioTable

#Region "List Of Fields"
    Private strNameTable As String
    Private bytInsertType As Byte
    Private intColumnsCount As Integer
    Private intRowsCount As Integer
    Private sngWidthCells As Single
    Private sngHeightCells As Single
    Private sngWidthTable As Single
    Private sngHeightTable As Single
    Private booDeleteTargetShape As Boolean
    Private booVisibleProgressBar As Boolean
#End Region

#Region "List Of Variables"
    Private vsoApp As Visio.Application = Globals.ThisAddIn.Application
    Private winObj As Visio.Window = vsoApp.ActiveWindow
    Private pagObj As Visio.Page = vsoApp.ActivePage
    Private shpObj As Visio.Shape

    Private shape_TbL As Visio.Shape
    Private shape_ThC As Visio.Shape
    Private shape_TvR As Visio.Shape
    Private shape_ClW As Visio.Shape

    Private arrNewID() As Integer
    Private CountID As Integer = 0

    Private Const UTN = "User.TableName"
    Private Const UTR = "User.TableRow"
    Private Const UTC = "User.TableCol"
    Private Const PX = "PinX"
    Private Const PY = "PinY"
    Private Const WI = "Width"
    Private Const HE = "Height"
    Private Const CA = "Angle"
    Private Const LD = "LockDelete"
    Private Const GU = "=GUARD("
    Private Const strATC = "!Actions.Titles.Checked=1,"
    Private Const strACC = "!Actions.Comments.Checked=1,"
    Private Const strThGu000 = "THEMEGUARD(MSOTINT(RGB(0,0,0),50))"
    Private Const strThGu255 = "THEMEGUARD(RGB(255,255,255))"
    Private Const strThGu191 = "THEMEGUARD(RGB(191,191,191))"
    Private Const GS = "=GUARD(Sheet."
    Private Const GI = "=Guard(IF("
    Private Const sh = "Sheet."
    Private Const frm = "###0.0###"
    Private Const GU5 = "=GUARD(10 mm)" ' Переделать на DrawUn
    Private Const P50 = "50%"
    Private Const GT = "GUARD(TRUE)"
    Private Const G1 = "Guard(1)"
#End Region

    Public Sub New(ByVal a As String, ByVal b As Byte, ByVal c As Integer, ByVal d As Integer, ByVal e As Single, _
                   ByVal f As Single, ByVal g As Single, ByVal h As Single, ByVal i As Boolean, ByVal j As Boolean)

        strNameTable = IIf(Trim(a) = "", "TbL", a)
        bytInsertType = IIf(b > 0 Or b < 5, b, 1)
        intColumnsCount = IIf(c > 0 Or c < 1001, c, 5)
        intRowsCount = IIf(c > 0 Or c < 1001, c, 5)
        sngWidthCells = IIf(e < 1, 20, e)
        sngHeightCells = IIf(f < 1, 10, f)
        sngWidthTable = IIf(g < 1, 200, g)
        sngHeightTable = IIf(h < 1, 100, h)
        booDeleteTargetShape = i
        booVisibleProgressBar = j

    End Sub

    Public Sub CreatTable()
        On Error GoTo errD

        Dim vsoLayerTitles As Visio.Layer, vsoLayerCells As Visio.Layer, MemSHID As Integer
        Dim TypeCell As String, VarCell As Byte, sngTW As Double, sngTH As Double
        Dim sngTX As Double, sngTY As Double
        Dim ThemeC As String, ThemeE As String
        Dim jGT As Integer = 0
        Dim iGT As Integer = 0

        winObj = vsoApp.ActiveWindow
        pagObj = vsoApp.ActivePage

        ThemeC = pagObj.ThemeColors
        If ThemeC <> "visThemeNone" Then pagObj.ThemeColors = "visThemeNone"
        ThemeE = pagObj.ThemeEffects
        If ThemeE <> "visThemeEffectsNone" Then pagObj.ThemeEffects = "visThemeEffectsNone"

        ' Добавление и изменение свойств слоев
        vsoLayerTitles = pagObj.Layers.Add("Titles_Tables")
        vsoLayerCells = pagObj.Layers.Add("Cells_Tables")
        If vsoLayerTitles.CellsC(4).Result("") = 0 Then vsoLayerTitles.CellsC(4).FormulaForceU = 1 ' Сделать если надо слой видимым
        If vsoLayerTitles.CellsC(7).Result("") = 1 Then vsoLayerTitles.CellsC(7).FormulaForceU = 0 ' Разблокировать если надо слой
        If vsoLayerCells.CellsC(4).Result("") = 0 Then vsoLayerCells.CellsC(4).FormulaForceU = 1 ' Сделать если надо слой видимым
        If vsoLayerCells.CellsC(7).Result("") = 1 Then vsoLayerCells.CellsC(7).FormulaForceU = 0 ' Разблокировать если надо слой
        vsoLayerTitles.CellsC(5).FormulaU = "GUARD(0)" ' Слой всегда не печатаемый

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
                sngTH = (sngPH - sngPTM - sngPBM) / intRowsCount
            End With
        End If

        If bytInsertType = 4 Then
            If winObj.Selection.Count = 0 Then
                MsgBox("Вы должны выбрать одну фигуру")
                Exit Sub
            Else
                With winObj.Selection(1)
                    .BoundingBox(3, sngTX, sngTY, sngTW, sngTH) ' L, B, R, T
                    MemSHID = .ID
                    sngTX = vsoApp.FormatResult(sngTX, "", 64, "0.0000")
                    sngTY = vsoApp.FormatResult(sngTH, "", 64, "0.0000")
                    sngTW = .Cells(WI).Result(64) / intColumnsCount
                    sngTH = .Cells(HE).Result(64) / intRowsCount
                End With
            End If
        End If

        ReDim arrNewID((intRowsCount * intColumnsCount) + (intRowsCount + intColumnsCount))
        CountID = -1

        'If booVisibleProgressBar Then
        Dim frm As New dlgWait
        frm.Label1.Text = " " & vbCrLf & "Создание  новой таблицы"
        frm.Show() : frm.Refresh()
        'End If

        Call RecUndo("Создание таблицы...")

        vsoApp.ShowChanges = False

        TypeCell = strNameTable : VarCell = 3 'Вставка 1 ячейки
        NewShape(TypeCell)
        shpObj = shape_TbL
        DrawOfCells(VarCell, sngTW, sngTH, sngTX, sngTY, iGT, jGT)
        vsoLayerTitles.Add(shpObj, 1)

        TypeCell = "ThC" : VarCell = 2 'Вставка 1 ряда таблицы
        For iGT = 1 To intColumnsCount
            If iGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_ThC
            Else
                shpObj = shape_ThC.Duplicate
            End If
            DrawOfCells(VarCell, sngTW, sngTH, sngTX, sngTY, iGT, jGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "TvR" : VarCell = 1 'Вставка 1 столбца таблицы
        For jGT = 1 To intRowsCount
            If jGT = 1 Then
                NewShape(TypeCell)
                shpObj = shape_TvR
            Else
                shpObj = shape_TvR.Duplicate
            End If
            DrawOfCells(VarCell, sngTW, sngTH, sngTX, sngTY, iGT, jGT)
            vsoLayerTitles.Add(shpObj, 1)
        Next

        TypeCell = "ClW" : VarCell = 0 'Вставка рабочих ячеек
        For jGT = 1 To intRowsCount
            'If booVisibleProgressBar Then
            frm.lblProgressBar.Width = (300 / intRowsCount) * jGT : frm.lblProgressBar.Refresh() : Application.DoEvents()
            'End If
            For iGT = 1 To intColumnsCount
                If jGT = 1 And iGT = 1 Then
                    NewShape(TypeCell)
                    shpObj = shape_ClW
                Else
                    shpObj = shape_ClW.Duplicate
                End If
                DrawOfCells(VarCell, sngTW, sngTH, sngTX, sngTY, iGT, jGT)
                vsoLayerCells.Add(shpObj, 0)
            Next
        Next

        shpObj = pagObj.Shapes.ItemFromID(arrNewID(0))
        shpObj.Cells(UTC).FormulaU = "GUARD(" & intColumnsCount & ")"
        shpObj.Cells(UTR).FormulaU = "GUARD(" & intRowsCount & ")"

        For iGT = 0 To intColumnsCount + intRowsCount - 1
            winObj.Page.Shapes.ItemFromID(arrNewID(iGT)).Cells("LockTextEdit").FormulaU = "Guard(1)"
        Next

        Call RecUndo("0")
        frm.Close()
        'If booVisibleProgressBar Then Unload frmWait

        If booDeleteTargetShape Then
            If pagObj.Shapes.ItemFromID(MemSHID).Cells(LD).Result("") = 0 Then pagObj.Shapes.ItemFromID(MemSHID).DeleteEx(0)
        End If

        vsoApp.ShowChanges = True
        winObj.Select(shpObj, 258)

        If ThemeC <> pagObj.ThemeColors Then pagObj.ThemeColors = ThemeC
        If ThemeE <> pagObj.ThemeEffects Then pagObj.ThemeEffects = ThemeE

        Exit Sub
errD:
        MsgBox("CreatTable-Class" & vbNewLine & Err.Description)
    End Sub

    Private Sub DrawOfCells(VarCell, sngTW, sngTH, sngTX, sngTY, ByVal iGT, ByVal jGT)
        On Error GoTo errD
        With pagObj
            CountID = CountID + 1
            arrNewID(CountID) = shpObj.ID
            With shpObj
                Select Case VarCell

                    Case 0 'Вставка рабочих ячеек
                        .Cells(PX).FormulaForceU = GS & arrNewID(iGT) & "!PinX)"
                        .Cells(PY).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!PinY)"
                        .Cells(WI).FormulaForceU = GS & arrNewID(iGT) & "!Width)"
                        .Cells(HE).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!Height)"
                        .Cells(UTN).FormulaForceU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GS & arrNewID(iGT) & "!User.TableCol)"
                        .Cells(UTR).FormulaForceU = GS & arrNewID(intColumnsCount + jGT) & "!User.TableRow)"
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & "User.TableCol.Prompt&"" ""&User.TableCol&CHAR(10)&User.TableRow.Prompt&"" ""&User.TableRow" & "," & """""" & "))"

                    Case 2 ' упр строка
                        .Cells(PX).FormulaForceU = GS & arrNewID(CountID - 1) & "!PinX+(Sheet." & arrNewID(CountID - 1) & "!Width/2)+(Width/2))"
                        .Cells(PY).FormulaForceU = GS & arrNewID(0) & "!PinY)"
                        .Cells(HE).FormulaForceU = GS & arrNewID(0) & "!Height)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(WI).Result(64) = sngTW
                        ElseIf bytInsertType = 3 Then
                            .Cells(WI).Result(64) = sngWidthTable / intColumnsCount
                        Else
                            .Cells(WI).Result(64) = sngWidthCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaForceU = GU & iGT & ")"
                        .Cells(UTR).FormulaForceU = GU & 0 & ")"
                        .Characters.AddCustomFieldU(UTC, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTC & ")"
                        .Cells("Char.Color").FormulaForceU = strThGu255
                        .Cells("FillForegnd").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu000 & "," & strThGu255 & "))"
                        .Cells("FillForegndTrans").FormulaForceU = GI & sh & arrNewID(0) & strATC & "0%" & "," & "50%" & "))"
                        .Cells("LineColor").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu191 & "," & strThGu255 & "))"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrNewID(0) & "!Actions.Titles.Checked))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Управляющая ячейка столбца""" & "," & """""" & "))"

                    Case 1 ' упр столбец
                        .Cells(PX).FormulaU = GS & arrNewID(0) & "!PinX)"
                        If jGT = 1 Then
                            .Cells(PY).FormulaU = GS & arrNewID(CountID - intColumnsCount - 1) & "!PinY-(Sheet." & arrNewID(CountID - intColumnsCount - 1) & "!Height/2)-(Height/2))"
                        Else
                            .Cells(PY).FormulaForceU = GS & arrNewID(CountID - 1) & "!PinY-(Sheet." & arrNewID(CountID - 1) & "!Height/2)-(Height/2))"
                        End If
                        .Cells(WI).FormulaU = GS & arrNewID(0) & "!Width)"
                        If bytInsertType = 2 Or bytInsertType = 4 Then
                            .Cells(HE).Result(64) = sngTH
                        ElseIf bytInsertType = 3 Then
                            .Cells(HE).Result(64) = sngHeightTable / intRowsCount
                        Else
                            .Cells(HE).Result(64) = sngHeightCells
                        End If
                        .Cells(UTN).FormulaU = GS & arrNewID(0) & "!Name(0))"
                        .Cells(UTC).FormulaU = GU & 0 & ")"
                        .Cells(UTR).FormulaForceU = GU & jGT & ")"
                        .Characters.AddCustomFieldU(UTR, 0)
                        .Cells("Fields.Value").FormulaU = "GUARD(" & UTR & ")"
                        .Cells("Char.Color").FormulaForceU = strThGu255
                        .Cells("FillForegnd").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu000 & "," & strThGu255 & "))"
                        .Cells("FillForegndTrans").FormulaForceU = GI & sh & arrNewID(0) & strATC & "0%" & "," & "50%" & "))"
                        .Cells("LineColor").FormulaForceU = GI & sh & arrNewID(0) & strATC & strThGu191 & "," & strThGu255 & "))"
                        .Cells("HideText").FormulaForceU = "=GUARD(NOT(" & sh & arrNewID(0) & "!Actions.Titles.Checked))"
                        .Cells("Comment").FormulaForceU = GI & sh & arrNewID(0) & strACC & """Управляющая ячейка строки""" & "," & """""" & "))"

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
        MsgBox("DrawOfCells" & vbNewLine & Err.Description)
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
                    'AddSectionNum = 240 ' Добавить Action секцию
                    'intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    'arrRowData = {{"SelectWindow", "RUNMACRO(""MainModule.LoadSelectfrmWorks"",""Таблица_V"")", """Выде&лить...""", 2035, 1, 0, "FALSE", "FALSE"}, _
                    '    {"InsertWindow", "RUNMACRO(""MainModule.LoadInsertfrmWorks"",""Таблица_V"")", """Вставить...""", 22, 2, 0, "FALSE", "FALSE"}, _
                    '    {"SizeWindow", "RUNMACRO(""MainModule.LoadSizefrmWorks"",""Таблица_V"")", """Разм&еры...""", 1605, 3, 0, "FALSE", "TRUE"}, _
                    '    {"AutoSizeWindow", "RUNMACRO(""MainModule.LoadAutoSizefrmWorks"",""Таблица_V"")", """Авторазмер...""", 583, 4, 0, "FALSE", "FALSE"}, _
                    '    {"MoreWindow", "RUNMACRO(""MainModule.LoadMorefrmWorks"",""Таблица_V"")", """Дополнительно...""", 586, 5, 0, "FALSE", "FALSE"}, _
                    '    {"GutT", "RUNMACRO(""MainModule.GutT"",""Таблица_V"")", """Вырезать текст из ячеек""", 2046, 6, 0, "FALSE", "TRUE"}, _
                    '    {"CopyT", "RUNMACRO(""MainModule.CopyT"",""Таблица_V"")", """Копировать текст из ячеек""", 2045, 7, 0, "FALSE", "FALSE"}, _
                    '    {"PasteT", "RUNMACRO(""MainModule.PasteT"",""Таблица_V"")", """Вставить текст в ячейки""", 22, 8, 0, "FALSE", "FALSE"}, _
                    '    {"IntCells", "RUNMACRO(""MainModule.IntDeIntCells"",""Таблица_V"")", """Объединить ячейки""", 402, 9, 0, "FALSE", "TRUE"}, _
                    '    {"DeIntCells", "RUNMACRO(""MainModule.IntDeIntCells"",""Таблица_V"")", """Отменить объединение""", 402, 9, 0, "TRUE", "TRUE"}}
                    'AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)
                    shape_ClW = vsoShape

                Case "TvR" ' УЯ строки
                    'AddSectionNum = 240 ' Добавить Action секцию
                    'intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    'arrRowData = {{"AddRowTop", "RUNMACRO(""MainModule.AddRowBefore"",""Таблица_V"")", """&Добавить строку выше""", 296, 1, 0, "FALSE", "FALSE"}, _
                    '    {"AddRowBottom", "RUNMACRO(""MainModule.AddRowAfter"",""Таблица_V"")", """&Добавить строку ниже""", 296, 1, 0, "FALSE", "FALSE"}, _
                    '    {"SelectRow", "RUNMACRO(""MainModule.SelRows"",""Таблица_V"")", """Выде&лить строку""", 801, 2, 0, "FALSE", "FALSE"}, _
                    '    {"AlignOnTextHeight", "RUNMACRO(""MainModule.OnHeight"",""Таблица_V"")", """По высоте &текста""", 541, 3, 0, "FALSE", "FALSE"}, _
                    '    {"DeleteROw", "RUNMACRO(""MainModule.DelRows"",""Таблица_V"")", """&Удалить строку""", 292, 4, 0, "FALSE", "True"}}
                    'AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    AddSectionNum = 9 ' Добавить Control секцию
                    intArrNum = {0, 1, 2, 3, 6, 8} ' Сделать не меньше нуля
                    arrRowData = {{"ControlHeight", "GUARD(Width*0)", "Height*0", "GUARD(Controls.ControlHeight)", "GUARD(Controls.ControlHeight.Y)", "False", """Изменение высоты ячейки"""}}
                    AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

                    .CellsSRC(10, 1, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 2, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    .CellsSRC(10, 5, 1).FormulaU = "GUARD(Controls.ControlHeight.Y)"
                    shape_TvR = vsoShape

                Case "ThC" ' УЯ столбца
                    'AddSectionNum = 240 ' Добавить Action секцию
                    'intArrNum = {3, 0, 15, 16, 4, 7, 8}
                    'arrRowData = {{"AddColumnLeft", "RUNMACRO(""MainModule.AddColBefore"",""Таблица_V"")", """&Добавить столбец слева""", 297, 1, 0, "FALSE", "FALSE"}, _
                    '    {"AddColumnRight", "RUNMACRO(""MainModule.AddColAfter"",""Таблица_V"")", """&Добавить столбец справа""", 297, 1, 0, "FALSE", "FALSE"}, _
                    '    {"SelectColumn", "RUNMACRO(""MainModule.SelCols"",""Таблица_V"")", """Выде&лить столбец""", 802, 2, 0, "FALSE", "FALSE"}, _
                    '    {"AlignOnTextWidth", "RUNMACRO(""MainModule.OnWidth"",""Таблица_V"")", """По ширине &текста""", 542, 3, 0, "FALSE", "FALSE"}, _
                    '    {"DeleteColumn", "RUNMACRO(""MainModule.DelCols"",""Таблица_V"")", """&Удалить столбец""", 294, 4, 0, "FALSE", "True"}}
                    'AddSections(vsoShape, AddSectionNum, arrRowData, intArrNum)

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
                    arrRowData = {{"Titles", "SETF(GetRef(Actions.Titles.Checked),NOT(Actions.Titles.Checked))", """П&оказывать заголовки""", """""", 5, 1, "FALSE", "TRUE"}, _
                        {"Comments", "SETF(GetRef(Actions.Comments.Checked),NOT(Actions.Comments.Checked))", """Показывать коммента&рии""", """""", 6, 1, "FALSE", "FALSE"}, _
                        {"FixingTable", "SETF(GetRef(Actions.FixingTable.Checked),NOT(Actions.FixingTable.Checked))", """Заф&иксировать таблицу""", """""", 7, 0, "FALSE", "FALSE"}}
                    'arrRowData = {{"NewTable", "RUNMACRO(""MainModule.LoadfrmAddTable"",""Таблица_V"")", """Соз&дать таблицу""", 333, 1, 0, "FALSE", "FALSE"}, _
                    '    {"SelectTable", "RUNMACRO(""MainModule.SelTable"",""Таблица_V"")", """Выде&лить таблицу""", 803, 2, 0, "FALSE", "FALSE"}, _
                    '    {"LockPicture", "RUNMACRO(""MainModule.LoadfrmPicture"",""Таблица_V"")", """Закрепить &фигуры""", 954, 3, 0, "FALSE", "FALSE"}, _
                    '    {"LinkData", "RUNMACRO(""MainModule.LoadfrmLinkData"",""Таблица_V"")", """Вне&шние данные""", 4143, 4, 0, "FALSE", "FALSE"}, _
                    '    {"Titles", "SETF(GetRef(Actions.Titles.Checked),NOT(Actions.Titles.Checked))", """П&оказывать заголовки""", """""", 5, 1, "FALSE", "TRUE"}, _
                    '    {"Comments", "SETF(GetRef(Actions.Comments.Checked),NOT(Actions.Comments.Checked))", """Показывать коммента&рии""", """""", 6, 1, "FALSE", "FALSE"}, _
                    '    {"FixingTable", "SETF(GetRef(Actions.FixingTable.Checked),NOT(Actions.FixingTable.Checked))", """Заф&иксировать таблицу""", """""", 7, 0, "FALSE", "FALSE"}, _
                    '    {"DeleteTable", "RUNMACRO(""MainModule.DelTab"",""Таблица_V"")", """&Удалить таблицу""", 2487, 8, 0, "FALSE", "TRUE"}, _
                    '    {"Help", "RUNMACRO(""MainModule.CallHelp"",""Таблица_V"")", """Справка""", 3998, 8, 0, "FALSE", "FALSE"}}
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
        MsgBox("NewShape" & vbNewLine & Err.Description)
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
        MsgBox("AddSections" & vbNewLine & Err.Description)
    End Sub


End Class
