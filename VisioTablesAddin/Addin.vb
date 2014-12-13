Imports System.Drawing
Imports System.Windows.Forms


Partial Public Class Addin
    Public Property Application As Microsoft.Office.Interop.Visio.Application

    Public Sub OnAction(commandId As String, commandTag As String)
        Select Case commandId
            Case "btn_newcolumnbefore", "btn_newcolumnafter" : MessageBox.Show(commandTag) ': AddColumns (Button.Tag)
            Case "btn_newrowbefore", "btn_newrowafter" : MessageBox.Show(commandTag) ': AddRows (Button.Tag)
            Case "btn_onwidth" : MessageBox.Show(commandId) ': OnWidth
            Case "btn_onheight" : MessageBox.Show(commandId) ': OnHeight
            Case "btn_onwidthheight" : MessageBox.Show(commandId) ': AllOnText
            Case "btn_newtable" : MessageBox.Show(commandId) ': LoadfrmAddTable
            Case "btn_seltable", "btn_selrange", "btn_selcolumn", "btn_selrow" : MessageBox.Show(commandTag) ': SelCell (Button.Tag)
            Case "btn_seltxt", "btn_selnum", "btn_selnotnum", "btn_seldate", "btn_selempty", "btn_selnotempty", "btn_selinvert" : MessageBox.Show(commandTag) ': SelInContent (Button.Tag)
            Case "btn_text", "btn_date", "btn_time", "btn_comment", "btn_numcol", "btn_numrow" : MessageBox.Show(commandTag) ': InsertText (Button.Tag)
            Case "btn_intdeint" : MessageBox.Show(commandId) ': IntDeIntCells
            Case "btn_dropdownlist" : MessageBox.Show(commandId) ': LoadfrmSelectFromList
            Case "btn_gut" : MessageBox.Show(commandId) ': GutT
            Case "btn_copy" : MessageBox.Show(commandId) ': CopyT
            Case "btn_paste" : MessageBox.Show(commandId) ': PasteT
            Case "btn_deltable" : MessageBox.Show(commandId) ': DelTab
            Case "btn_delcolumn" : MessageBox.Show(commandId) ': DelCols
            Case "btn_delrow" : MessageBox.Show(commandId) ': DelRows
            Case "btn_intellinput" : MessageBox.Show(commandId) ': LoadfrmIntellInput
            Case "btn_seldialog" : MessageBox.Show(commandId) ': LoadSelectfrmWorks
            Case "btn_size" : MessageBox.Show(commandId) ': LoadSizefrmWorks
            Case "btn_autosize" : MessageBox.Show(commandId) ': LoadAutoSizefrmWorks
            Case "btn_fromfile" : MessageBox.Show(commandId) ': LoadfrmFromFile
            Case "btn_altlinesrow", "btn_altlinescol" : MessageBox.Show(commandTag) ': AlternatLines (Button.Tag)
            Case "btn_moredialog" : MessageBox.Show(commandId) ': LoadMorefrmWorks
            Case "btn_extdata" : MessageBox.Show(commandId) ': LoadfrmLinkData
            Case "btn_convert1Shape" : MessageBox.Show(commandId) ': ConvertInto1Shape
            Case "btn_lockpicture" : MessageBox.Show(commandId) ': LoadfrmPicture
            Case "btn_help" : MessageBox.Show(commandId) ': CallHelp
        End Select
    End Sub

    Public Sub QuickTable(strV)
        Dim nC As Integer = 0, nR As Integer = 0, w As Single = 0, h As Single = 0
        strV = Right(strV, Len(strV) - 1)
        nC = Val(Left(strV, InStr(1, strV, "x", 1) - 1)) : nR = Val(Right(strV, Len(strV) - InStr(1, strV, "x", 1)))
        w = Application.FormatResult(20, 70, 64, "#.0000")
        h = Application.FormatResult(8, 70, 64, "#.0000")
        MessageBox.Show("Таблица: " & nC & " x " & nR)
    End Sub

#Region "RibbonFunctions"

    Public Function IsCommandAltEnabled(commandId As String) As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    Public Function IsCommandEnabled(commandId As String) As Boolean
        If Application.ActiveWindow.Selection.Count > 0 AndAlso _
            Application.ActiveWindow.Selection(1).CellExistsU("User.TableName", 0) Then Return True
        Return False
    End Function

    Sub Startup(app As Object)
        Application = DirectCast(app, Microsoft.Office.Interop.Visio.Application)
        AddHandler Application.SelectionChanged, AddressOf Application_SelectionChanged
        AddHandler Application.DocumentCreated, AddressOf Application_DocumentListChanged
        AddHandler Application.DocumentOpened, AddressOf Application_DocumentListChanged
        AddHandler Application.BeforeDocumentClose, AddressOf Application_DocumentListChanged
    End Sub

    Private Sub Application_DocumentListChanged(ByVal doc As Microsoft.Office.Interop.Visio.Document)
        UpdateUI()
    End Sub

    Private Sub Application_SelectionChanged(ByVal window As Visio.Window)
        UpdateUI()
    End Sub

    Sub UpdateUI()
        UpdateRibbon()
    End Sub

#End Region

End Class