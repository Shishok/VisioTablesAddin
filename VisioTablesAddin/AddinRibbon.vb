Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

<ComVisible(True)> _
Partial Public Class Addin
    Implements IRibbonExtensibility
    Private _ribbon As Microsoft.Office.Core.IRibbonUI

#Region "IRibbonExtensibility Members"

    Public Function GetCustomUI(ribbonId As String) As String Implements IRibbonExtensibility.GetCustomUI
        Return My.Resources.Ribbon
    End Function

#End Region

#Region "Ribbon Callbacks"

    Public Function IsRibbonSplitEnabled(ctrl As Microsoft.Office.Core.IRibbonControl) As Boolean
        Return IsSplitEnabled(ctrl.Id)
    End Function

    Public Function IsRibbonCommandEnabled(ctrl As Microsoft.Office.Core.IRibbonControl) As Boolean
        Return IsCommandEnabled(ctrl.Id)
    End Function

    Public Sub OnRibbonButtonClick(control As Microsoft.Office.Core.IRibbonControl)
        OnAction(control.Id, control.Tag)
    End Sub

    Public Sub OnRibbonGalleryClick(control As Microsoft.Office.Core.IRibbonControl, id As String, index As Integer)
        QuickTable(id)
    End Sub

    Public Sub OnRibbonLoad(ribbonUI As Microsoft.Office.Core.IRibbonUI)
        _ribbon = ribbonUI
    End Sub

#End Region

    Public Sub UpdateRibbon()
        _ribbon.Invalidate()
    End Sub

End Class