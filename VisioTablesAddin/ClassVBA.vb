Imports System.Runtime.InteropServices

<ComVisible(True)> _
Public Interface IClassVBA

    'Свойства надстройки =========================================================================================

    ' Свойства таблицы
    ReadOnly Property TableName()
    ReadOnly Property TableColumnsCount()
    ReadOnly Property TableRowsCount()
    ReadOnly Property TableWidth()
    ReadOnly Property TableHeight()
    ReadOnly Property TableLeftBorder()
    ReadOnly Property TableTopBorder()
    ReadOnly Property TableRightBorder()
    ReadOnly Property TableBottomBorder()

    'Свойства столбца/строки
    ReadOnly Property ColumnIndex()
    ReadOnly Property RowIndex()
    ReadOnly Property ColumnCellsCount()
    ReadOnly Property RowCellsCount()
    ReadOnly Property ColumnWidth()
    ReadOnly Property RowHeight()

    'Свойства ячейки
    ReadOnly Property CellWidth()
    ReadOnly Property CellHeight()
    ReadOnly Property CellWidthDimension()
    ReadOnly Property CellHeightDimension()

    'Свойства выделенных ячеек
    ReadOnly Property SelectedColumnCount()
    ReadOnly Property SelectedRowCount()
    ReadOnly Property SelectedIsPossibleMerge()


    'Методы надстройки =========================================================================================

    Sub AddTable(a As String, b As Byte, c As Integer, d As Integer, e As Single, f As Single, g As Single, h As Single, i As Boolean, j As Boolean)
    Sub AddColumn(arg As Byte)
    Sub AddRow(arg As Byte)
    Sub MergeCells()
    Sub UnMergeCell()
    Sub SelectCells(c As Integer, r As Integer, c1 As Integer, r1 As Integer, deSel As Boolean)
    Sub SelectCellsExt(arg As String)
    Sub FindText(Oper As String, Patt As String, Act As String)
    Sub ReplaceText(txt As String, txt1 As String, istart As Integer, icount As Integer)
    Sub GutText()
    Sub CopyText()
    Sub PasteText()
    Sub PasteTxtFile(arg As String, RowCount As Integer, rep As Boolean)
    Sub PasteExcelFile(Path As String, Page As String, Address As String)
    Sub LinkToData(a As Integer, b As Boolean, c As String, d As Boolean, e As Boolean, h As Boolean)
    Sub SetCellsText(arg As Object, c As Integer, r As Integer, c1 As Integer, r1 As Integer, byCorR As Byte)
    Sub SetCellsFormula(cell As String, c As Integer, r As Integer, c1 As Integer, r1 As Integer, txt As Object)
    Sub SortTable(NumColumn As Byte, a As Boolean, b As Boolean)
    Sub GetCellsFormula(cell As String, c As Integer, r As Integer, c1 As Integer, r1 As Integer, ByRef arr As Object, res As String)
    Sub GetCellsProp(ByRef arr As Object, c As Integer, r As Integer, c1 As Integer, r1 As Integer, arg As String)
    Sub AutoFit(arg As String)
    Sub AutoFitExt(a As Boolean, b As Boolean, c As Byte, d As Byte, e As Boolean, f As Boolean)
    Sub ResizeCellsOrTable(a As Byte, b As Boolean, c As Single, d As Single, e As Single, f As Single, g As Boolean, h As Boolean)
    Sub CellsDistributeSize(arg)
    Sub CellsBanded(arg As String)
    Sub TextOrientation(ang As Double)
    Sub GroupTable()
    Sub LockShape(hAL As Byte, Val As Byte, shN As Boolean, lF As Boolean, msg As Boolean)
    Sub DeleteColumn()
    Sub DeleteRow()
    Sub DeleteTable(arg As Boolean)

End Interface

<ComVisible(True)> _
<ClassInterface(ClassInterfaceType.None)> _
Public Class ClassVBA
    Implements IClassVBA

#Region "Properties"

    'Свойства  надстройки =========================================================================================

	' Имя таблицы (в главной упр. ячейке)
    Public ReadOnly Property TableName Implements IClassVBA.TableName
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return winObj.Selection.PrimaryItem.Cells("User.TableName").ResultStr("")
        End Get
    End Property

	' Количество столбцов в таблице (без УЯ)
    Public ReadOnly Property TableColumnsCount Implements IClassVBA.TableColumnsCount
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return shpsObj(NT).Cells("User.TableCol").Result("")
        End Get
    End Property

	' Количество строк в таблице (без УЯ)
    Public ReadOnly Property TableRowsCount Implements IClassVBA.TableRowsCount
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return shpsObj(NT).Cells("User.TableRow").Result("")
        End Get
    End Property

	' Ширина таблицы (в милиметрах)
    Public ReadOnly Property TableWidth Implements IClassVBA.TableWidth
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return CDbl(vsoApp.FormatResult(fSTWH(winObj.Selection(1), 1, False), 64, 64, "#0.0000"))
        End Get
    End Property

	' Высота таблицы (в милиметрах)
    Public ReadOnly Property TableHeight Implements IClassVBA.TableHeight
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return CDbl(vsoApp.FormatResult(fSTWH(winObj.Selection(1), 2, False), 64, 64, "#0.0000"))
        End Get
    End Property

	' Левая граница таблицы (в милиметрах)
    Public ReadOnly Property TableLeftBorder Implements IClassVBA.TableLeftBorder
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return PositionBorders("left", ArrShapeID(1, 1))
        End Get
    End Property

	' Верхняя граница таблицы (в милиметрах)
    Public ReadOnly Property TableTopBorder Implements IClassVBA.TableTopBorder
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return PositionBorders("top", ArrShapeID(1, 1))
        End Get
    End Property

	' Правая граница таблицы (в милиметрах)
    Public ReadOnly Property TableRightBorder Implements IClassVBA.TableRightBorder
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return PositionBorders("right", GetShapeId(UBound(ArrShapeID, 1), UBound(ArrShapeID, 2)))

        End Get
    End Property

	' Нижняя граница таблицы (в милиметрах)
    Public ReadOnly Property TableBottomBorder Implements IClassVBA.TableBottomBorder
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return PositionBorders("bottom", GetShapeId(UBound(ArrShapeID, 1), UBound(ArrShapeID, 2)))
        End Get
    End Property

	' Номер активного столбца
    Public ReadOnly Property ColumnIndex() Implements IClassVBA.ColumnIndex
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return winObj.Selection.PrimaryItem.Cells(UTC).Result("")
        End Get
    End Property

	' Номер активной строки
    Public ReadOnly Property RowIndex Implements IClassVBA.RowIndex
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return winObj.Selection.PrimaryItem.Cells(UTR).Result("")
        End Get
    End Property

	' Количество ячеек в активном столбце (без УЯ)
    Public ReadOnly Property ColumnCellsCount Implements IClassVBA.ColumnCellsCount
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)

            Dim iCount As Integer = 0
            Dim iC As Integer = winObj.Selection.PrimaryItem.Cells(UTC).Result("")

			For r = 1 To UBound(ArrShapeID, 2)
				If ArrShapeID(iC, r) <> 0 Then iCount += 1
			Next
				
            Return iCount
        End Get
    End Property

	' Количество ячеек в активной строке (без УЯ)
    Public ReadOnly Property RowCellsCount Implements IClassVBA.RowCellsCount
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)

            Dim iCount As Integer = 0

            Dim iR As Integer = winObj.Selection.PrimaryItem.Cells(UTR).Result("")

			For c = 1 To UBound(ArrShapeID, 1)
				If ArrShapeID(c, iR) <> 0 Then iCount += 1
			Next

            Return iCount
        End Get
    End Property

	' Ширина активного столбца (в милиметрах) 
    Public ReadOnly Property ColumnWidth Implements IClassVBA.ColumnWidth
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return CDbl(vsoApp.FormatResult(shpsObj.ItemFromID(ArrShapeID(winObj.Selection.PrimaryItem.Cells(UTC).Result(""), 0)).Cells("Width").Result(""), "in", 64, "#0.0000"))
        End Get
    End Property

	' Ширина активной строки (в милиметрах) 
    Public ReadOnly Property RowHeight Implements IClassVBA.RowHeight
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Call InitArrShapeID(NT)
            Return CDbl(vsoApp.FormatResult(shpsObj.ItemFromID(ArrShapeID(0, winObj.Selection.PrimaryItem.Cells(UTR).Result(""))).Cells("Height").Result(""), "in", 64, "#0.0000"))
        End Get
    End Property

	' Размерность активной ячейки по ширине
    Public ReadOnly Property CellWidthDimension() Implements IClassVBA.CellWidthDimension
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Dim TestArray() As String = Split(winObj.Selection.PrimaryItem.Cells("Width").FormulaU, ",")
            Return UBound(TestArray) + 1
        End Get
    End Property

	' Размерность активной ячейки по высоте
    Public ReadOnly Property CellHeightDimension() Implements IClassVBA.CellHeightDimension
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Dim TestArray() As String = Split(winObj.Selection.PrimaryItem.Cells("Height").FormulaU, ",")
            Return UBound(TestArray) + 1
        End Get
    End Property

	' Ширина активной ячейки (в милиметрах)
    Public ReadOnly Property CellWidth Implements IClassVBA.CellWidth
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return CDbl(vsoApp.FormatResult(winObj.Selection.PrimaryItem.Cells("Width").Result(""), "", 64, "#0.0000"))
        End Get
    End Property

	' Высота активной ячейки (в милиметрах)
    Public ReadOnly Property CellHeight Implements IClassVBA.CellHeight
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Return CDbl(vsoApp.FormatResult(winObj.Selection.PrimaryItem.Cells("Height").Result(""), "", 64, "#0.0000"))
        End Get
    End Property

	' Количество выделенных столбцов
    Public ReadOnly Property SelectedColumnCount Implements IClassVBA.SelectedColumnCount
        Get
            If Not CheckCells() Then
                Return -1
            End If
            Return SelColRow(1)
        End Get
    End Property

	' Количество выделенных строк
    Public ReadOnly Property SelectedRowCount Implements IClassVBA.SelectedRowCount
        Get
            If Not CheckCells() Then
                Return -1
            End If
            Return SelColRow(2)
        End Get
    End Property

	' Возможно ли объединить выделенные ячейки (true, false)
    Public ReadOnly Property SelectedIsPossibleMerge Implements IClassVBA.SelectedIsPossibleMerge
        Get
            If Not CheckCells() Then
                Return -1
            End If

            Dim booCheck As Boolean = True
            Dim vsoSel As Visio.Selection = winObj.Selection
            Dim cMin, cMax, rMin, rMax As Integer

            Call InitArrShapeID(NT)
            Call GetMinMaxRange(vsoSel, cMin, cMax, rMin, rMax)

            With shpsObj
                For c = cMin To cMax
                    For r = rMin To rMax
                        If InStr(1, .ItemFromID(ArrShapeID(c, r)).Cells("Width").FormulaU, "SUM", 1) <> 0 OrElse _
                                InStr(1, .ItemFromID(ArrShapeID(c, r)).Cells("Height").FormulaU, "Sum", 1) <> 0 OrElse _
                                ArrShapeID(c, r) = 0 OrElse ColumnIndex = 0 OrElse RowIndex = 0 Then
                            booCheck = False
                            Exit For
                        End If
                    Next
                Next
            End With

            Return booCheck
        End Get
    End Property

#End Region

#Region "Methods"

    'Методы надстройки =========================================================================================
	
	' Создание новой таблицы на активном листе
    Public Sub AddTable(a As String, b As Byte, c As Integer, d As Integer, e As Single, f As Single, g As Single, h As Single, i As Boolean, j As Boolean) Implements IClassVBA.AddTable
        Dim NewTable As New VisioTable
        NewTable.CreatTable(a, b, c, d, e, f, g, h, i, j)
        NewTable = Nothing
    End Sub

	' Вставка нового столбца в активную таблицу
    Public Sub AddColumn(arg As Byte) Implements IClassVBA.AddColumn
        If Not CheckCells() Then Exit Sub
        Select Case arg
            Case 0, 1 : Call AddColumns(arg)
        End Select
    End Sub

	' Вставка новой строки в активную таблицу
    Public Sub AddRow(arg As Byte) Implements IClassVBA.AddRow
        If Not CheckCells() Then Exit Sub
        Select Case arg
            Case 0, 1 : Call AddRows(arg)
        End Select
    End Sub

	' Объединение выделенных ячеек в одну
    Public Sub MergeCells() Implements IClassVBA.MergeCells
        Call IntDeIntCells()
    End Sub

	' Разъединение выделенной ячейки
    Public Sub UnMergeCell() Implements IClassVBA.UnMergeCell
        Call IntDeIntCells()
    End Sub

	' Выделение ячеек в активной таблице по номерам столбцов/строк
    Public Sub SelectCells(c As Integer, r As Integer, c1 As Integer, r1 As Integer, deSel As Boolean) Implements IClassVBA.SelectCells
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)
        If deSel Then winObj.DeselectAll()

        If c1 < c Then
            Dim x As Integer = c1 : c1 = c : c = x
        End If
        If r1 < r Then
            Dim y As Integer = r1 : r1 = r : r = y
        End If

        Call SelectCls(c, r, c1, r1)
    End Sub

	' Выделение ячеек в активной таблице по параметрам
    Public Sub SelectCellsExt(arg As String) Implements IClassVBA.SelectCellsExt
        If Not CheckCells() Then Exit Sub

        Select Case StrConv(arg, vbLowerCase)
            Case "all table", "0" : SelCell(1, True)
            Case "table", "1" : SelCell(2, True)
            Case "range", "2" : SelCell(3, True)
            Case "column", "3" : SelCell(4, True)
            Case "row", "4" : SelCell(5, True)
            Case "text", "5" : SelInContent(1)
            Case "value", "6" : SelInContent(2)
            Case "not value", "7" : SelInContent(5)
            Case "date", "8" : SelInContent(3)
            Case "empty", "9" : SelInContent(6)
            Case "not empty", "10" : SelInContent(7)
            Case "invert", "11" : SelInContent(8)
        End Select

    End Sub

	' Поиск текста в ячейках
    Public Sub FindText(Oper As String, Patt As String, Act As String) Implements IClassVBA.FindText
        If Not CheckCells() Then Exit Sub
        Call SearchText(Oper, Patt, Act)
    End Sub

	' Замена текста в ячейках
    Public Sub ReplaceText(txt As String, txt1 As String, istart As Integer, icount As Integer) Implements IClassVBA.ReplaceText
        If Not CheckCells() Then Exit Sub
        Call ReplaceTxt(txt, txt1, istart, icount)
    End Sub

	' Вырезание текста из выделенных ячеек в буфер обмена
    Public Sub GutText() Implements IClassVBA.GutText
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)
        Call GutT()
    End Sub

	' Копирование текста из выделенных ячеек в буфер обмена
    Public Sub CopyText() Implements IClassVBA.CopyText
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)
        Call CopyT()
    End Sub

	' Вставка содержимого буфера обмена в ячейки таблицы
    Public Sub PasteText() Implements IClassVBA.PasteText
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)
        Call PasteT()
    End Sub

	' Заполнение однотипными или просто данными ячейки таблицы из текстового файла
    Public Sub PasteTxtFile(arg As String, RowCount As Integer, rep As Boolean) Implements IClassVBA.PasteTxtFile
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)

        Dim arrRow() As String = IO.File.ReadAllLines(arg, System.Text.Encoding.Default)

        Dim vsoObj As Visio.Selection = winObj.Selection

        Dim iC As Integer, iR As Integer

        If RowCount < 1 Then
            iR = 0
        ElseIf RowCount > UBound(arrRow) + 1 Then
            iR = UBound(arrRow)
        Else
            iR = RowCount - 1
        End If

        Call RecUndo("Вставить из файла")

        On Error Resume Next

        Select Case rep
            Case True
                For iC = 1 To vsoObj.Count
                    vsoObj(iC).Characters.Text = arrRow(iR)

                    If iR = UBound(arrRow) Then
                        iR = 0
                    Else : iR = iR + 1
                    End If

                Next
            Case False
                Dim i As Integer
                For iC = iR To UBound(arrRow)
                    i += 1
                    If i > vsoObj.Count Then GoTo Line1
                    vsoObj(i).Characters.Text = arrRow(iC)
                Next
        End Select

Line1:
        Call RecUndo("0")
    End Sub

	' Заполнение данными ячейки таблицы из файла Excel
    Public Sub PasteExcelFile(Path As String, Page As String, Address As String) Implements IClassVBA.PasteExcelFile
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)

        Dim oExcel As Object, arr As Object
        oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open(Path)

        arr = oExcel.Sheets(Page).Range(Address).Value ' массив таблицы Excel
        Dim ShapeObj As Visio.Shape = winObj.Selection(1)


        Call RecUndo("Вставить из Excel")

        On Error Resume Next

        If IsArray(arr) Then
            Dim iCol As Integer = ShapeObj.Cells(UTC).Result("") - LBound(arr, 1)
            Dim iRow As Integer = ShapeObj.Cells(UTR).Result("") - LBound(arr, 2)

            For i = LBound(arr, 1) To UBound(arr, 1)
                For j = LBound(arr, 2) To UBound(arr, 2)
                    shpsObj.ItemFromID(ArrShapeID(j + iCol, i + iRow)).Characters.Text = arr(i, j).ToString()
                Next
            Next
        Else
            winObj.Selection(1).Characters.Text = arr.ToString()
        End If

        Call RecUndo("0")

        oExcel.Workbooks(1).Close(True)
        oExcel.Quit()
    End Sub

	' Связывание ячеек таблицы с подключенным внешним источником данных
    Public Sub LinkToData(a As Integer, b As Boolean, c As String, d As Boolean, e As Boolean, h As Boolean) Implements IClassVBA.LinkToData
        If Not CheckCells() Then Exit Sub
        Dim lngRowIDs() As Integer = vsoApp.ActiveDocument.DataRecordsets.Item(a).GetDataRowIDs("")
        Dim f As Integer = UBound(lngRowIDs)
        Dim g As Integer = vsoApp.ActiveDocument.DataRecordsets.Item(a).DataColumns.Count
        Call LinkToDataInShapes(a - 1, b, c, d, e, f, g, h)
    End Sub

	' Вставка пользовательских данных в выделенные ячейки таблицы
    Public Sub SetCellsText(arg As Object, c As Integer, r As Integer, c1 As Integer, r1 As Integer, byCorR As Byte) Implements IClassVBA.SetCellsText
        If Not CheckCells() Then Exit Sub
        Call RecUndo("Задать текст")
        Call SetText(arg, c, r, c1, r1, byCorR)
        Call RecUndo("0")
    End Sub

	' Изменение формулы/значения ячеек
    Public Sub SetCellsFormula(cell As String, c As Integer, r As Integer, c1 As Integer, r1 As Integer, txt As Object) Implements IClassVBA.SetCellsFormula
        If Not CheckCells() Then Exit Sub

        Call RecUndo("Задать формулу/значение")
        Call SetFormula(cell, c, r, c1, r1, txt)
        Call RecUndo("0")

    End Sub

	' Сортировка выделенных ячеек (по столбцам)
    Public Sub SortTable(NumColumn As Byte, a As Boolean, b As Boolean) Implements IClassVBA.SortTable
        If Not CheckCells() Then Exit Sub
        Call SortTableData(NumColumn, a, b)
    End Sub

	' Извлечение формулы/значения заданных ячеек из активной таблицы
    Public Sub GetCellsFormula(cell As String, c As Integer, r As Integer, c1 As Integer, r1 As Integer, ByRef arr As Object, res As String) Implements IClassVBA.GetCellsFormula
        If Not CheckCells() Then Exit Sub
        Call GetFormula(cell, c, r, c1, r1, arr, res)
    End Sub

	' Извлечение формулы/значения  заданных ячеек из активной таблицы
    Public Sub GetCellsProp(ByRef arr As Object, c As Integer, r As Integer, c1 As Integer, r1 As Integer, arg As String) Implements IClassVBA.GetCellsProp
        If Not CheckCells() Then Exit Sub
        Call GetCellsProperties(arr, c, r, c1, r1, arg)
    End Sub

	' Подгонка ширины/высоты столбцов/строк по размерам текста всех ячеек находящихся в столбце/строке
    Public Sub AutoFit(arg As String) Implements IClassVBA.AutoFit
        If Not CheckCells() Then Exit Sub
        Call InitArrShapeID(NT)
        Select Case StrConv(arg, vbLowerCase)
            Case "column", "0" : Call AlignOnText(winObj.Selection.PrimaryItem.Cells(UTC).Result(""), 4)
            Case "row", "1" : Call AlignOnText(winObj.Selection.PrimaryItem.Cells(UTR).Result(""), 5)
        End Select
    End Sub

	' Подгонка столбцов/строк/таблиц по размерам текста по заданными параметрам
    Public Sub AutoFitExt(a As Boolean, b As Boolean, c As Byte, d As Byte, e As Boolean, f As Boolean) Implements IClassVBA.AutoFitExt
        If Not CheckCells() Then Exit Sub
        Call AllAlignOnText(a, b, c, d, e, f)
    End Sub

	' установка новых размеров столбцов/строк или всей таблицы
    Public Sub ResizeCellsOrTable(a As Byte, b As Boolean, c As Single, d As Single, e As Single, f As Single, g As Boolean, h As Boolean) Implements IClassVBA.ResizeCellsOrTable
        If Not CheckCells() Then Exit Sub

        With winObj.Page.PageSheet
            If g Then e = .Cells("PageWidth").Result(64) - .Cells("PageRightMargin").Result(64) - .Cells("PageLeftMargin").Result(64)
            If h Then f = .Cells("PageHeight").Result(64) - .Cells("PageTopMargin").Result(64) - .Cells("PageBottomMargin").Result(64)
        End With

        Call ResizeCells(a, b, c, d, e, f, g, h)
    End Sub

	' Установка одинаковыми по размеру выделенные столбцы/строки
    Public Sub CellsDistributeSize(arg) Implements IClassVBA.CellsDistributeSize
        If Not CheckCells() Then Exit Sub
        Select Case StrConv(arg, vbLowerCase)
            Case "columns", "0" : Call AlignOnSize(4)
            Case "rows", "1" : Call AlignOnSize(5)
            Case "columns, rows", "2" : Call AlignOnSize(4) : Call AlignOnSize(5)
        End Select
    End Sub

    ' Создание "полосатой" по столбцам или строкам таблицы 
    Public Sub CellsBanded(arg As String) Implements IClassVBA.CellsBanded
        If Not CheckCells() Then Exit Sub
        Select Case StrConv(arg, vbLowerCase)
            Case "columns", "0" : Call AlternatLines(4)
            Case "rows", "1" : Call AlternatLines(5)
        End Select
    End Sub

	' Разворот текста в выделенных ячейках в соответствии с заданным углом
    Public Sub TextOrientation(ang As Double) Implements IClassVBA.TextOrientation
        If Not CheckCells() Then Exit Sub
        Call AllRotateText(True, ang)
    End Sub

	' Преобразование активной таблицы в единую сгруппированную фигуру
    Public Sub GroupTable() Implements IClassVBA.GroupTable
        If Not CheckCells() Then Exit Sub
        Call ConvertInto1Shape()
    End Sub

    ' Закрепление картинки/фигуры в ячейках таблицы в соответствии c заданными параметрами
    Public Sub LockShape(hAL As Byte, Val As Byte, shN As Boolean, lF As Boolean, msg As Boolean) Implements IClassVBA.LockShape
        winObj = vsoApp.ActiveWindow
        Call LockPicture(hAL, Val, shN, lF, msg)
    End Sub

	' Удаление выделенных столбцов из таблицы на активном листе
    Public Sub DeleteColumn() Implements IClassVBA.DeleteColumn
        If Not CheckCells() Then Exit Sub
        Call DelColRows(0)
    End Sub

	' Удаление выделенных строк из таблицы на активном листе
    Public Sub DeleteRow() Implements IClassVBA.DeleteRow
        If Not CheckCells() Then Exit Sub
        Call DelColRows(1)
    End Sub

	' Удаление активной таблицы на активном листе
    Public Sub DeleteTable(arg As Boolean) Implements IClassVBA.DeleteTable
        If Not CheckCells() Then Exit Sub
        Call DelTab(arg)
    End Sub

#End Region

#Region "Private Sub and Function"

    ' Проверка на отсутствие/некорректном выделении на листе 
    Private Function CheckCells() As Boolean
        winObj = vsoApp.ActiveWindow : shpsObj = winObj.Page.Shapes
        Return CheckSelCells()
    End Function

    ' Определение координат ячеек на листе
    Private Function PositionBorders(arg, ID) As Double
        Dim dblTop, dblBottom, dblLeft, dblRight, dbtmp As Double

        shpsObj.ItemFromID(ID).BoundingBox(1, dblLeft, dblBottom, dblRight, dblTop)

        Select Case arg
            Case "left" : dbtmp = dblLeft
            Case "top" : dbtmp = dblTop
            Case "right" : dbtmp = dblRight
            Case "bottom" : dbtmp = dblBottom
        End Select

        Return CDbl(vsoApp.FormatResult(dbtmp, "in", 64, "#0.0000"))
    End Function

#End Region

End Class