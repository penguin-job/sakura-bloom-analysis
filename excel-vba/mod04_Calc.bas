Attribute VB_Name = "mod04_Calc"
Option Explicit
Sub totalDataCalc()
'
'   気温のデータを開花日まで1件ずつ累積していく。
'
    Dim totalTemp As Long
    Dim rs1 As String
    Dim rs2 As String
    Dim ts As String
    Dim ws As String
   
    Dim idxRow As Long
   
    Dim idxCol As Long
    Dim endFlg As Integer
    Dim widx As Long
     
    Dim workYear As Date
    Dim workTemp As Long
    
     Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得
   
    
    rs1 = "bloom_date"
    rs2 = "weather_data"
    ts = "weather_data_temp"
    ws = "work_weather_bloom"
    
 '   該当期間の開花日取得
   
    Sheets(ws).Select
    Rows("3:100").Select
    Selection.Delete Shift:=xlUp
'
    Sheets(rs1).Select
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$1406").AutoFilter Field:=1, Criteria1:=p.location
    ActiveSheet.Range("$A$1:$E$1406").AutoFilter Field:=2, Criteria1:=">=" & p.yearFrom _
        , Operator:=xlAnd, Criteria2:="<=" & p.yearTo
    Range("A2:G1500").Select
    Selection.Copy
    
    Sheets(ws).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    
    Sheets(rs1).Select

    Selection.AutoFilter
    Range("A1").Select
   
 '
    idxRow = 2
    Sheets(ws).Select
    If Cells(idxRow, 1) = "" Then
       MsgBox "データがありません"
       GoTo sub900
    End If
    
    Do Until endFlg = 9
    
    '    normal date の年を整える
         Cells(idxRow, 4).Value = DateSerial(Cells(idxRow, 2).Value, Month(Cells(idxRow, 4).Value), Day(Cells(idxRow, 4).Value))
    '    normal dateとの差異
         Cells(idxRow, 5).Value = Cells(idxRow, 3).Value - Cells(idxRow, 4).Value
         
    '    起点日
         workYear = DateSerial((Cells(idxRow, 2).Value), 2, 1)
         
    '   対象気温データを抽出する
    
        Sheets(rs2).Select
        Rows("1:1").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$E$1500").AutoFilter Field:=1, Criteria1:= _
            ">=" & workYear, Operator:=xlAnd, Criteria2:="<=" & Sheets(ws).Cells(idxRow, 4).Value
        ActiveWindow.SmallScroll Down:=0
        Range("A1:E1500").Select
        Selection.Copy
        
        Sheets(ts).Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A1").Select
        
        Sheets(rs2).Select
        Selection.AutoFilter
        Range("A1").Select
        
       Sheets(ws).Select
       
        widx = 2
        
       
       '　抽出した気温から積算気温を計算
       Do Until Sheets(ts).Cells(widx, 1) = ""
            Sheets(ws).Cells(idxRow, 6) = Sheets(ws).Cells(idxRow, 6) + Sheets(ts).Cells(widx, 4)
            Sheets(ws).Cells(idxRow, 7) = Sheets(ws).Cells(idxRow, 7) + Sheets(ts).Cells(widx, 3)
            
            widx = widx + 1
       Loop
         
       Cells(idxRow, 8).Value = workYear      '起点日
       Cells(idxRow, 9).Value = Cells(idxRow, 3).Value - workYear    '経過日数
  
        ' うるう年補正（2/29以降のみ）
        Dim stdDays As Long
        Dim d As Date

        d = Cells(idxRow, 4).Value
        stdDays = DateDiff("d", DateSerial(Year(d), 2, 1), d) + 1

        If Month(d) > 2 And Day(DateSerial(Year(d), 2, 29)) = 29 Then
            stdDays = stdDays - 1
        End If
        '-------

        Cells(idxRow, 10).Value = stdDays     '標準経過日数
        Cells(idxRow, 11).Value = Format(Cells(idxRow, 3).Value, "m/d")    'データラベル
        
        idxRow = idxRow + 1
         
        If Cells(idxRow, 1) = "" Then
            endFlg = 9
        End If
         
    Loop
    
sub900:
    

End Sub


Sub totalCalcForecast()
'
'  今年のデータと予想値をまとめてグラフデータを作る
'
    Dim i As Integer
    Dim idxRow As Long
    Dim totalTemp As Long
    Dim rs1 As String
    Dim rs2 As String
    Dim ts As String
    Dim ws As String
    
    rs1 = "weather_data"
    rs2 = "weather_forecast"
    ws = "work_weather_bloom_current"
    ts = "weather_data_temp"
    
 '   いったんtemporaryシートで計算
 
    Sheets(rs1).Select
    idxRow = Cells(1, 1).End(xlDown).Row - 1
    Range("A2:E" & idxRow).Select
    Selection.Copy
    Sheets(ts).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    
    Sheets(rs2).Select
    Range("A2:F100").Select
    Selection.Copy
    
    Sheets(ts).Select
    Range("A" & idxRow + 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    
    idxRow = Cells(1, 1).End(xlDown).Row - 1  '再計算
    Cells(2, 6) = Cells(2, 4)
  '
    For i = 3 To idxRow + 1
        Cells(i, 6) = Cells(i - 1, 6) + Cells(i, 4)
    Next
    
    Range("A2:F" & idxRow).Select
    Selection.Copy
    Sheets(ws).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select


End Sub
