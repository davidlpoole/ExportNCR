Option Explicit

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
' When a hyperlink is clicked in a worksheet which contains
' this code, the following sub-routine is called.

    'only if hyperlink is in Column S
    If Target.Range.Column = "19" Then
        Call ExportNCR
    End If
    
End Sub
