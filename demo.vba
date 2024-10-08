def return_schema_evolution(full_table_names: List[strl) -> DataFrame:
Returns the schema evolution of a list of tables passed.
Args:
full_table_names (List[str]): A list of full table names, in format:
'catalog, schema. table'
Returns:
df (DataFrame): A DataFrame containing schema evolution for passed tables.
spark = SparkSession.getActiveSession()
sql_to_run = '
WITH CTE_HIST
AS (
+ ' UNION ALL ' join( [f''SELECT '{table}' AS table_name,
* FROM (DESCRIBE HISTORY {table))™ for table in full_table_names]) + ™™
),
CTE STRUCT
AS (
SELECT *
explode(from_son(operationParameters.columns,
'array<struct<column: struct<name: string>>>'
). column. name) AS column
FROM CTE_ HIST
WHERE operation = 'ADD COLUMNS'
UNION ALL SELECT *
explode(from_json(operationParameters.columns,
'array<string>'
)) AS column
FROM CTE_ HIST
WHERE operation = 'DROP COLUMNS'
SELECT table_name AS table_name
column AS column_name
IF(operation = 'ADD COLUMNS', 'ADDED', 'REMOVED') AS added_or_removed
timestamp AS schema_evolution_date_time version AS changed_by_delta_version userName AS changed_by_user job. jobId AS changed_by_job_id job. jobName AS changed_by_job_name
FROM CTE_STRUCT
110000
df = spark.sql (sql_to_run)
return df

Sub AddBorders(rng As Range)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub









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

















Sub RearrangeColumns()
    Dim wsBusiness As Worksheet
    Dim wsConfig As Worksheet
    Dim lastRow As Long
    Dim currentDate As String
    Dim rng As Range
    
    ' Set references to the worksheets
    Set wsBusiness = ThisWorkbook.Sheets("BusinessFile") ' Replace with your actual sheet name
    
    ' Check if the ConfigFile sheet already exists and delete it if it does
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("ConfigFile").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create a new sheet for the config file format
    Set wsConfig = ThisWorkbook.Sheets.Add
    wsConfig.Name = "ConfigFile"
    
    ' Find the last row with data in the Business file
    lastRow = wsBusiness.Cells(wsBusiness.Rows.Count, 1).End(xlUp).Row
    
    ' Get the current date
    currentDate = Format(Date, "mm/dd/yyyy") ' Format can be changed as needed
    
    ' Define the ranges for the columns to be copied
    wsConfig.Range("A1:A" & lastRow).Value = wsBusiness.Range("D1:D" & lastRow).Value ' Column A from Business Column D
    wsConfig.Range("B1:B" & lastRow).Value = wsBusiness.Range("E1:E" & lastRow).Value ' Column B from Business Column E
    wsConfig.Range("C1:C" & lastRow).Value = wsBusiness.Range("F1:F" & lastRow).Value ' Column C from Business Column F
    wsConfig.Range("D1:D" & lastRow).Value = wsBusiness.Range("G1:G" & lastRow).Value ' Column D from Business Column G
    wsConfig.Range("E1:E" & lastRow).Value = wsBusiness.Range("H1:H" & lastRow).Value ' Column E from Business Column H
    wsConfig.Range("F1:F" & lastRow).Value = wsBusiness.Range("I1:I" & lastRow).Value ' Column F from Business Column I
    
    ' Insert the current date into Column G for each row
    wsConfig.Range("G1:G" & lastRow).Value = currentDate
    
    ' Set headers for the Config file columns
    wsConfig.Range("A1").Value = "A"
    wsConfig.Range("B1").Value = "B"
    wsConfig.Range("C1").Value = "C"
    wsConfig.Range("D1").Value = "D"
    wsConfig.Range("E1").Value = "E"
    wsConfig.Range("F1").Value = "F"
    wsConfig.Range("G1").Value = "Date" ' Header for the date column
    
    ' Add borders to all copied ranges
    Set rng = wsConfig.Range("A1:G" & lastRow)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    MsgBox "Config file format created successfully with borders and the current date!", vbInformation
End Sub




Sub ExportToCSV()
    Dim wsBusiness As Worksheet
    Dim wsTemp As Worksheet
    Dim lastRow As Long
    Dim currentDate As String
    Dim csvFilePath As String
    Dim csvFileName As String
    Dim rng As Range
    Dim i As Long
    Dim cellValue As String
    
    ' Set references to the worksheet
    Set wsBusiness = ThisWorkbook.Sheets("BusinessFile") ' Replace with your actual sheet name
    
    ' Create a temporary worksheet for data manipulation
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "TempSheet"
    
    ' Find the last row with data in the Business file
    lastRow = wsBusiness.Cells(wsBusiness.Rows.Count, 1).End(xlUp).Row
    
    ' Get the current date
    currentDate = Format(Date, "mm/dd/yyyy") ' Format can be changed as needed
    
    ' Copy and paste the columns to match the config file format
    wsTemp.Range("A1:A" & lastRow).Value = wsBusiness.Range("D1:D" & lastRow).Value ' Column A from Business Column D
    wsTemp.Range("B1:B" & lastRow).Value = wsBusiness.Range("E1:E" & lastRow).Value ' Column B from Business Column E
    wsTemp.Range("C1:C" & lastRow).Value = wsBusiness.Range("F1:F" & lastRow).Value ' Column C from Business Column F
    wsTemp.Range("D1:D" & lastRow).Value = wsBusiness.Range("G1:G" & lastRow).Value ' Column D from Business Column G
    wsTemp.Range("E1:E" & lastRow).Value = wsBusiness.Range("H1:H" & lastRow).Value ' Column E from Business Column H
    wsTemp.Range("F1:F" & lastRow).Value = wsBusiness.Range("I1:I" & lastRow).Value ' Column F from Business Column I
    
    ' Insert the current date into Column G for each row
    wsTemp.Range("G1:G" & lastRow).Value = currentDate
    
    ' Extract 'full' from Column K and place it into Column H
    For i = 1 To lastRow
        cellValue = wsBusiness.Cells(i, "K").Value
        If InStr(cellValue, "full load weekly") > 0 Then
            wsTemp.Cells(i, "H").Value = "full"
        End If
    Next i
    
    ' Set headers for the Config file columns
    wsTemp.Range("A1").Value = "A"
    wsTemp.Range("B1").Value = "B"
    wsTemp.Range("C1").Value = "C"
    wsTemp.Range("D1").Value = "D"
    wsTemp.Range("E1").Value = "E"
    wsTemp.Range("F1").Value = "F"
    wsTemp.Range("G1").Value = "Date" ' Header for the date column
    wsTemp.Range("H1").Value = "Extracted" ' Header for the extracted column
    
    ' Define CSV file path and name
    csvFileName = "ConfigFile_" & Format(Date, "yyyymmdd") & ".csv" ' Name with current date
    csvFilePath = ThisWorkbook.Path & "\" & csvFileName
    
    ' Save the temporary worksheet as CSV
    wsTemp.Copy
    With ActiveWorkbook
        .SaveAs Filename:=csvFilePath, FileFormat:=xlCSV, CreateBackup:=False
        .Close False
    End With
    
    ' Delete the temporary worksheet
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
    
    MsgBox "CSV file created successfully at: " & csvFilePath, vbInformation
End Sub
