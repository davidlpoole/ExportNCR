Sub ExportNCR()

'
' David Poole 01/12/2020
' https://github.com/davidlpoole/ExportNCR
'

    Application.ScreenUpdating = False
    On Error GoTo Err 'Make sure screen updating is reverted
   
    Dim cSelectedNCR As Range
    Dim strSelectedNCR As String
    Dim strDir, strFile, strFileExt As String
    Dim wbNCRRegister, wbNCRForm As Workbook
    
    Set wbNCRRegister = ActiveWorkbook
          
    'Get NCR number from the selected row and store var for later
    
    Set cSelectedNCR = Cells(Range(Selection.Address).Row, 1)
    strSelectedNCR = Trim(cSelectedNCR.Value)
    
    'Validate NCR number (check format is "##-###")
    If Len(strSelectedNCR) = 0 Or Not (strSelectedNCR Like "##-###") Then
        MsgBox ("Invalid 'NCR Number' in cell " & cSelectedNCR.Address)
        GoTo Err
    End If
    
    'set the save directory and file extension
    '(can specify a directory, or just use the current directory)
    'strDir = "H:\Business Analysis\QA\NCR"
    strDir = wbNCRRegister.Path 'excludes trailing backslash
    strFileExt = ".xlsx"
    strFile = strDir & "\NCR_" & strSelectedNCR & strFileExt
    
    'if file exists, then open file, else create new
    If Not Dir(strFile, vbDirectory) = vbNullString Then
        Set wbNCRForm = Workbooks.Open(strFile)
     Else
        'Switch to form and enter the NCR number
        Sheets("NCR Form").Visible = True
        Sheets("NCR Form").Select
        Range("S2:W2").Value = strSelectedNCR
        
        'move form to a new workbook and save (filename = NCR number)
        Sheets("NCR Form").Select
        Application.CutCopyMode = False
        Sheets("NCR Form").Copy
        ActiveWorkbook.SaveAs Filename:=strFile, _
            FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Set wbNCRForm = ActiveWorkbook
        
        'Remove formulas from sheet
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Range("A1").Select
        
        'Save file
        wbNCRForm.Save
        
        'hide the NCR form and show the main sheet again
        wbNCRRegister.Sheets("NCR Form").Visible = False
        wbNCRRegister.Sheets("NCR Register 2020").Activate

    End If
    
Err:
    Application.ScreenUpdating = True

End Sub
