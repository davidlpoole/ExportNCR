Sub ExportNCR()

'
' David Poole 01/12/2020
'
' TODO (Maybe): show/hide NCR form in master sheet
'
'

    Application.ScreenUpdating = False
   
    'Get NCR number from the selected row and store var for later
    Dim strSelectedNCR As String
    strSelectedNCR = Trim(Cells(Range(Selection.Address).Row, 1).Value)
    
    ' TODO 1 - validate NCR number (check format is "##-###")
    ' if selected row doesn't have an NCR number (is blank)
    If Len(strSelectedNCR) = 0 Then
    MsgBox ("No NCR Number found in cell " & Cells(Range(Selection.Address).Row, 1).Address)
    GoTo Err
    End If
    
    ' TODO 2
    ' if file exists, then open file
    
    'Switch to form and enter the NCR number
    Sheets("NCR Form").Visible = True
    Sheets("NCR Form").Select
    Range("S2:W2").Value = strSelectedNCR
    
    'move form to a new workbook and save (filename = NCR number)
    Sheets("NCR Form").Select
    Application.CutCopyMode = False
    Sheets("NCR Form").Copy
    ActiveWorkbook.SaveAs Filename:="H:\Business Analysis\QA\NCR\" & strSelectedNCR & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    'Remove formulas from sheet (P
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
    'Save file
    ActiveWorkbook.Save
    'ActiveWindow.Close
    
Err:
    Application.ScreenUpdating = True

End Sub
