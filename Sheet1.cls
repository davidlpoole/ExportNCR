Option Explicit

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
' When a hyperlink is clicked in a worksheet which contains
' this code, the following sub-routine is called.

    'only if hyperlink is in Column S
    If Target.Range.Column = "19" Then
        Call ExportNCR
    End If
    
End Sub

Private Sub hidesheet()
' create a sheet which can be used as the target
' hyperlinks whicih will be picked up by the
' function within the sheet code.
' then hide it, so excel doesn't switch to it
' when clicked

ActiveWorkbook.Sheets(2).Visible = xlVeryHidden

End Sub
