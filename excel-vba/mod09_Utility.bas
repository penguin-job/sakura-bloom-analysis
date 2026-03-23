Attribute VB_Name = "mod09_Utility"
Option Explicit

' Controlシート値をまとめて取得
Public Type ControlParams
    location As String
    lat As String
    lon As String
    yearFrom As Long
    yearTo As Long
    
    location_current As String
    lat_current As String
    lon_current As String

    
End Type

Public Function GetControlParams() As ControlParams

    Dim p As ControlParams
    
    With ThisWorkbook.Worksheets(CONTROL_SHEET_NAME)
        p.location = Trim(.Range(CELL_LOCATION).Value)
        p.lat = Trim(.Range(CELL_LAT).Value)
        p.lon = Trim(.Range(CELL_LON).Value)
        p.yearFrom = .Range(CELL_YEAR_FROM).Value
        p.yearTo = .Range(CELL_YEAR_TO).Value
        
        p.location_current = Trim(.Range(CELL_LOCATION_CURRENT).Value)
        p.lat_current = Trim(.Range(CELL_LAT_CURRENT).Value)
        p.lon_current = Trim(.Range(CELL_LON_CURRENT).Value)
        
    End With
    GetControlParams = p
    
End Function


' Python実行共通関数
Public Function RunPython(cmd As String) As Variant

    Dim wsh As Object
    Dim exec As Object
    Dim stdOut As String, stdErr As String, exitCode As Long

    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec(cmd)

    Do While exec.Status = 0
        DoEvents
    Loop

    stdOut = exec.stdOut.ReadAll
    stdErr = exec.stdErr.ReadAll
    exitCode = exec.exitCode

    RunPython = Array(exitCode, stdOut, stdErr)
    
End Function


Sub SaveAndCloseThisBook()
'
'   BOOKを保存して終了
'
    ThisWorkbook.Save
    ThisWorkbook.Close

End Sub
