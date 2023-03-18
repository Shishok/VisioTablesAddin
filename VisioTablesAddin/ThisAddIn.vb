Public Class ThisAddIn

    Private ReadOnly _addin As Addin = New Addin()

    
    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addin
    End Function
    
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        _addin.Startup(Application)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        '_addin.Shutdown()
    End Sub


    Private accessVBA As ClassVBA

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If accessVBA Is Nothing Then
            accessVBA = New ClassVBA()
        End If
        Return accessVBA
    End Function

End Class
