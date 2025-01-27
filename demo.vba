from pyspark.sql import functions as F
from pyspark.sql.window import Window

def recursive_pipeline_search(df, initial_pipeline_run_id, max_depth=10):
  """
  Performs a recursive search through a DataFrame with pipeline_run_id and invokedById columns.

  Args:
    df: The input DataFrame.
    initial_pipeline_run_id: The starting pipeline_run_id value.
    max_depth: Maximum recursion depth to prevent infinite loops.

  Returns:
    A new DataFrame with columns pipeline_run_id, invokedById, and recursion_level.
  """

  # Create a window specification for ordering
  window = Window.orderBy("recursion_level")

  # Initialize the DataFrame with the initial pipeline_run_id
  result_df = df.filter(F.col("pipeline_run_id") == initial_pipeline_run_id).withColumn("recursion_level", F.lit(0))

  for i in range(1, max_depth + 1):
    # Join the current result with the original DataFrame
    joined_df = result_df.join(
        df,
        F.col("invokedById") == F.col("pipeline_run_id"),
        how="left"
    ).where(
        (F.col("invokedById").isNotNull() & F.col("pipeline_run_id").isNotNull()) | 
        (F.col("invokedById").isNull() & F.col("pipeline_run_id").isNull())
    )


    # Filter for new matches and add recursion level
    new_matches = joined_df.filter(F.col("pipeline_run_id_right").isNotNull()).select(
        F.col("pipeline_run_id_right").alias("pipeline_run_id"),
        F.col("invokedById_right").alias("invokedById"),
        F.lit(i).alias("recursion_level")
    )

    # Union new matches with existing results
    result_df = result_df.union(new_matches)

    # Check if any new matches were found
    if new_matches.count() == 0:
      break

  # Order the final result by recursion level
  return result_df.orderBy("recursion_level")

# Example usage
# Assuming 'df' is your DataFrame with 'pipeline_run_id' and 'invokedById' columns
initial_pipeline_run_id = 1 
result_df = recursive_pipeline_search(df, initial_pipeline_run_id)

# Show the results
result_df.show()























from pyspark.sql import functions as F

# Sample Data
data = [
    ("A", "B"),
    ("B", "C"),
    ("C", "D"),
    ("D", "E"),
    ("E", None)
]
columns = ["id1", "id2"]
df = spark.createDataFrame(data, columns)

# Starting point
start_id = "A"

# Initialize the result DataFrame with the starting id1
df_result = df.filter(df.id1 == start_id).withColumn("recursion_level", F.lit(0))

# Iterate and apply recursion
max_recursion_depth = 10  # A limit to avoid infinite recursion, adjust if needed
recursion_level = 0

while recursion_level < max_recursion_depth:
    # Get the next level of results by joining the current DataFrame with the original DataFrame
    df_next = df_result.alias("r").join(
        df.alias("t"),
        F.col("r.id2") == F.col("t.id1"),
        "inner"
    ).select(
        F.col("t.id1"), F.col("t.id2"), (F.col("r.recursion_level") + 1).alias("recursion_level")
    )
    
    # If no new rows are returned, break the loop
    if df_next.count() == 0:
        break
    
    # Append the results to the result DataFrame
    df_result = df_result.union(df_next)
    
    # Update recursion level
    recursion_level += 1

# Show the final result
df_result.show(truncate=False)






















start_id = "A"  # Initial starting id1 value
recursion_level = 0

# Initialize the first DataFrame with the starting id1
df_result = spark.sql(f"SELECT '{start_id}' as id1, NULL as id2, {recursion_level} as recursion_level")
df_result.createOrReplaceTempView("result_df")

while True:
    # Get the next level of matches
    df_next = spark.sql("""
        SELECT t.id1, t.id2, {recursion_level} + 1 as recursion_level
        FROM example_table t
        INNER JOIN result_df r ON t.id1 = r.id2
    """.format(recursion_level=recursion_level))
    
    # Check if new rows were found
    if df_next.rdd.isEmpty():
        break
    
    # Append the new rows to the result
    df_result = df_result.union(df_next)
    
    # Update the result view for the next iteration
    df_next.createOrReplaceTempView("result_df")
    
    # Increment recursion level
    recursion_level += 1

# Final result
df_result.show()



SELECT 
    CONCAT_WS(' | ', TRANSFORM(map_keys(parameters), k -> CONCAT(k, ':', parameters[k]))) AS parameters_str
FROM your_table






# Define the columns you want to select
select_cols = [
    'pipelineName as pipeline_name',
    'runid as run_id',
    'date_format(runStart, "yyyy-MM-dd HH:mm:ss") as start_time',
    'date_format(runEnd, "yyyy-MM-dd HH:mm:ss") as end_time',
    'message as error',
    'status'
]

# Construct the SQL query
sql_query = f"""
SELECT {', '.join(select_cols)},
       inline(MAP_ENTRIES(parameters)) AS key, value,  -- Convert map to array of structs and inline it
       CONCAT(
           CAST(FLOOR(durationInMs / 60000) AS STRING), ' min ',
           CAST(FLOOR((durationInMs % 60000) / 1000) AS STRING), ' sec'
       ) as duration,
       current_date() as requested_date,
       '{subscription_id}' as subscription_id,
       '{factory_name}' as factory_name,
       '{resource_group_name}' as resource_group_name
FROM tempDF
"""

# Execute the query
df_combined = spark.sql(sql_query)



































# Define the columns you want to select, as well as the necessary transformation and inline logic
select_cols = [
    'pipelineName as pipeline_name',
    'runid as run_id',
    'date_format(runStart, "yyyy-MM-dd HH:mm:ss") as start_time',
    'date_format(runEnd, "yyyy-MM-dd HH:mm:ss") as end_time',
    'message as error',
    'status'
]

# Construct the SQL query
sql_query = f"""
SELECT {', '.join(select_cols)},
       inline(transformed_params) AS key, value,
       CONCAT(
           CAST(FLOOR(durationInMs / 60000) AS STRING), ' min ',
           CAST(FLOOR((durationInMs % 60000) / 1000) AS STRING), ' sec'
       ) as duration,
       current_date() as requested_date,
       '{subscription_id}' as subscription_id,
       '{factory_name}' as factory_name,
       '{resource_group_name}' as resource_group_name
FROM (
    SELECT *,
           TRANSFORM(map_keys(parameters), map_values(parameters), (k, v) -> NAMED_STRUCT('key', k, 'value', v)) AS transformed_params
    FROM tempDF
) AS temp_with_params
"""

# Execute the query
df_combined = spark.sql(sql_query)














































# Define the list of columns you want to select
select_cols = [
    'pipelineName as pipeline_name',
    'runid as run_id',
    'date_format(runStart, "yyyy-MM-dd HH:mm:ss") as start_time',
    'date_format(runEnd, "yyyy-MM-dd HH:mm:ss") as end_time',
    'message as error',
    'status'
]

# Convert map `parameters` to array of structs for inlining
parameter_cols = [
    f"TRANSFORM(map_keys(parameters), map_values(parameters), (k, v) -> NAMED_STRUCT('key', k, 'value', v)) AS parameters"
]

# Build the SQL query
sql_query = f"""
SELECT {', '.join(select_cols)},
       inline({parameter_cols[0]}) AS parameters,  -- Use inline on transformed array of structs
       CONCAT(
           CAST(FLOOR(durationInMs / 60000) AS STRING), ' min ',
           CAST(FLOOR((durationInMs % 60000) / 1000) AS STRING), ' sec'
       ) as duration,
       current_date() as requested_date,
       '{subscription_id}' as subscription_id,
       '{factory_name}' as factory_name,
       '{resource_group_name}' as resource_group_name
FROM tempDF
"""

# Execute the query
df_combined = spark.sql(sql_query)














































from typing import List
from pyspark.sql import SparkSession, DataFrame

def return_schema_evolution(full_table_names: List[str]) -> DataFrame:
    """
    Returns the schema evolution of a list of tables passed.
    
    Args:
        full_table_names (List[str]): A list of full table names, in format:
                                      'catalog.schema.table'
    
    Returns:
        DataFrame: A DataFrame containing schema evolution for passed tables.
    """
    
    spark = SparkSession.getActiveSession()
    
    # Construct SQL query dynamically for each table
    sql_to_run = """
    WITH CTE_HIST AS (
        {}
    ),
    CTE_STRUCT AS (
        SELECT *, 
               explode(from_json(operationParameters.columns, 
                                 'array<struct<column: struct<name: string>>>'
                                ).column.name) AS column
        FROM CTE_HIST
        WHERE operation = 'ADD COLUMNS'
        
        UNION ALL 
        
        SELECT *, 
               explode(from_json(operationParameters.columns, 
                                 'array<string>'
                                )) AS column
        FROM CTE_HIST
        WHERE operation = 'DROP COLUMNS'
    )
    
    SELECT table_name AS table_name,
           column AS column_name,
           IF(operation = 'ADD COLUMNS', 'ADDED', 'REMOVED') AS added_or_removed,
           timestamp AS schema_evolution_date_time,
           version AS changed_by_delta_version,
           userName AS changed_by_user,
           job.jobId AS changed_by_job_id,
           job.jobName AS changed_by_job_name
    FROM CTE_STRUCT
    """.format(
        " UNION ALL ".join([f"SELECT '{table}' AS table_name, * FROM (DESCRIBE HISTORY {table})" for table in full_table_names])
    )
    
    # Execute the SQL query
    df = spark.sql(sql_to_run)
    return df
































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
