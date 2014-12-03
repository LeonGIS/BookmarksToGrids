Public Class BookmarkToGrid
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
    '
    '  TODO: Sample code showing how to access button host
    '
        My.ArcMap.Application.CurrentTool = Nothing
        Dim Form1 = New frmInput()
        Form1.ShowDialog()

        Form1 = Nothing

  End Sub

  Protected Overrides Sub OnUpdate()
    Enabled = My.ArcMap.Application IsNot Nothing
  End Sub
End Class
