Partial Public Class PanelProjectBridge
    Private Sub New()
        InitializeComponent()
    End Sub

    Friend Shared Function CreateDesignerOnly() As PanelProjectBridge
        Return New PanelProjectBridge()
    End Function
End Class
