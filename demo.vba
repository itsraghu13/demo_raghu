Sub RearrangeColumns()
    Dim wsBusiness As Worksheet
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    
    ' Set references to the worksheets
    Set wsBusiness = ThisWorkbook.Sheets("BusinessFile") ' Replace with your sheet name
    Set wsConfig = ThisWorkbook.Sheets.Add ' Create a new sheet for the config file format
    
    ' Rename the new sheet to ConfigFile
    wsConfig.Name = "ConfigFile"
    
    ' Find the last row with data in the Business file
    lastRow = wsBusiness.Cells(wsBusiness.Rows.Count, "A").End(xlUp).Row
    
    ' Copy and paste the columns to match the config file format
    wsConfig.Range("A1:A" & lastRow).Value = wsBusiness.Range("D1:D" & lastRow).Value ' Column A from Business Column D
    wsConfig.Range("B1:B" & lastRow).Value = wsBusiness.Range("E1:E" & lastRow).Value ' Column B from Business Column E
    wsConfig.Range("C1:C" & lastRow).Value = wsBusiness.Range("F1:F" & lastRow).Value ' Column C from Business Column F
    wsConfig.Range("D1:D" & lastRow).Value = wsBusiness.Range("G1:G" & lastRow).Value ' Column D from Business Column G
    wsConfig.Range("E1:E" & lastRow).Value = wsBusiness.Range("H1:H" & lastRow).Value ' Column E from Business Column H
    wsConfig.Range("F1:F" & lastRow).Value = wsBusiness.Range("I1:I" & lastRow).Value ' Column F from Business Column I
    
    ' Optional: Set headers for the Config file columns
    wsConfig.Range("A1").Value = "A"
    wsConfig.Range("B1").Value = "B"
    wsConfig.Range("C1").Value = "C"
    wsConfig.Range("D1").Value = "D"
    wsConfig.Range("E1").Value = "E"
    wsConfig.Range("F1").Value = "F"
    
    MsgBox "Config file format created successfully!", vbInformation
End Sub
