Attribute VB_Name = "mod02_Input"
Option Explicit

Sub ImportWeatherCSV()

  '   作成した気象データ（CSV）をシート[weather_data]にセットする。
  '
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得
   
    
    Dim csvPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("weather_data")
    csvPath = OUTPUT_DIR & "\data\weather_" & p.location & "_" & p.yearFrom & "_" & p.yearTo & ".csv"
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 65001 'UTF-8
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    


End Sub

Sub ImportSakuraBloomCSV()

  '   作成した桜開花データ（CSV）をシート[bloom_date]にセットする。
  '
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得

    Dim csvPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("bloom_date")
    csvPath = ThisWorkbook.path & "\data\sakura_bloom_all.csv"
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 65001 'UTF-8
        .Refresh BackgroundQuery:=False
        .Delete
    End With


End Sub

Sub ImportWeatherCurrentCSV()

  '   作成した今年の気象データ（CSV）をシート[weather_data]にセットする。
  '
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得
   
    p.yearFrom = Year(Date)
    
    Dim csvPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("weather_data")
    csvPath = OUTPUT_DIR & "\data\weather_" & p.location_current & "_" & p.yearFrom & ".csv"
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 65001 'UTF-8
        .Refresh BackgroundQuery:=False
        .Delete
    End With
 

End Sub

Sub ImportWeatherForecastCSV()

  '   作成した予報気象データ（CSV）をシート[weather_forecast]にセットする。
  '
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得
   
    p.yearFrom = Year(Date)
    
    Dim csvPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("weather_forecast")
    csvPath = OUTPUT_DIR & "\data\weather_forecast_" & p.location_current & ".csv"
    
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & csvPath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = 65001 'UTF-8
        .Refresh BackgroundQuery:=False
        .Delete
    End With
 

End Sub


