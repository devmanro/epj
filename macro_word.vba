Sub ConvertAllTablesToText()

    Dim objTable As Table
    Dim tableCount As Long
    
    ' Iterate backwards through the Tables collection for safe deletion/modification
    For tableCount = ActiveDocument.Tables.Count To 1 Step -1
        
        Set objTable = ActiveDocument.Tables(tableCount)
        
        ' Select the entire current table.
        objTable.Select
        
        ' Execute the conversion using the Tab separator.
        objTable.ConvertToText Separator:=wdSeparateByTabs
        
    Next tableCount
    
    MsgBox "Finished converting all tables to tab-separated text.", vbInformation
    
End Sub
