Attribute VB_Name = "CellCleaner"

Sub CellCleaner()

'Loops through all sheets and clears the contents of the specified range.

Dim sh As Worksheet
    
    For Each sh In ActiveWorkbook.Worksheets
        sh.Range("C8:L27").ClearContents
        sh.Range("C8:L27").Interior.Color = xlNone ' No Fill
        
    Next sh
    
End Sub

