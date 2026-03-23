Attribute VB_Name = "mod03_Process"
Option Explicit

'   指定された期間の気温の過去データ取得
'
Public Sub GetWeatherPastYears()

    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得

    ' 入力チェック
    If p.location = "" Or p.lat = "" Or p.lon = "" Then
        MsgBox "地域名または緯度経度が空です。", vbExclamation
        Exit Sub
    End If
    If p.yearFrom = 0 Or p.yearTo = 0 Then
        MsgBox "年が正しく設定されていません。", vbExclamation
        Exit Sub
    End If

    ' コマンド生成
    Dim cmd As String
    cmd = """" & PYTHON_EXE & """" & " " & _
          """" & SCRIPT_DIR & "\get_weather_past_years.py" & """" & " " & _
          """" & p.location & """" & " " & _
          p.lat & " " & p.lon & " " & _
          p.yearFrom & " " & p.yearTo & " " & _
          """" & OUTPUT_DIR & """"

    ' Python実行
    Dim result As Variant
    result = RunPython(cmd)

    If result(0) = 0 Then
        MsgBox "過去分の気象データ取得が完了しました。" & vbCrLf & result(1), vbInformation
    Else
        MsgBox "Python実行中にエラーが発生しました。" & vbCrLf & result(2), vbExclamation
    End If
End Sub

Sub GetSakuraBloom()
'
'   桜開花日データの取得
'
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得
   
    Dim cmd As String
    Dim wsh As Object
    Dim exec As Object
    Dim stdOut As String
    Dim stdErr As String

    cmd = PYTHON_EXE & " " & OUTPUT_DIR & "\get_sakura_bloom.py"""

    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec(cmd)

    Do While exec.Status = 0
        DoEvents
    Loop

    stdOut = exec.stdOut.ReadAll
    stdErr = exec.stdErr.ReadAll

    MsgBox "stdout:" & vbCrLf & stdOut & vbCrLf & vbCrLf & _
           "stderr:" & vbCrLf & stdErr
           
     ImportSakuraBloomCSV

End Sub

Sub GetWeatherCurrentYear()
'
'   今年の気温データ取得
'
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得

    ' 入力チェック
    If p.location_current = "" Then
        MsgBox "地域名が空です。", vbExclamation
        Exit Sub
    End If

    p.yearFrom = Year(Date)
    p.yearTo = Year(Date)

    ' コマンド生成
    Dim cmd As String
    cmd = """" & PYTHON_EXE & """" & " " & _
          """" & SCRIPT_DIR & "\get_weather_current_year.py" & """" & " " & _
          """" & p.location_current & """" & " " & _
          p.lat_current & " " & p.lon_current & " " & _
          p.yearFrom & " " & p.yearTo & " " & _
          """" & OUTPUT_DIR & """"

    ' Python実行
    Dim result As Variant
    result = RunPython(cmd)

    If result(0) = 0 Then
        MsgBox "今年分の気象データ取得が完了しました。", vbInformation
    Else
        MsgBox "Python実行中にエラーが発生しました。戻り値: " & result(2), vbExclamation
    End If

End Sub
Sub GetWeatherForecast()
'
'   16日先の気温データ取得
'
    Dim p As ControlParams
    p = GetControlParams()  ' Controlシート値を共通取得

    ' 入力チェック
    If p.location_current = "" Then
        MsgBox "地域名が空です。", vbExclamation
        Exit Sub
    End If

    p.yearFrom = Year(Date)
    p.yearTo = Year(Date)


    ' コマンド生成
    Dim cmd As String
    cmd = """" & PYTHON_EXE & """" & " " & _
          """" & SCRIPT_DIR & "\get_weather_forcast.py" & """" & " " & _
          """" & p.location_current & """" & " " & _
          p.lat_current & " " & p.lon_current & " " & _
          """" & OUTPUT_DIR & """"

    ' Python実行
    Dim result As Variant
    result = RunPython(cmd)
    
    MsgBox "16日予報を取得しました"

End Sub
