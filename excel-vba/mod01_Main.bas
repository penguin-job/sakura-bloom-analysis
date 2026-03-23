Attribute VB_Name = "mod01_Main"
Option Explicit

Public Sub RunGetWeatherPastYears()

'   過去の気温データと開花日を取得してグラフを作成する。

    Application.ScreenUpdating = False
    
    GetWeatherPastYears
    ImportWeatherCSV
    GetSakuraBloom
    ImportSakuraBloomCSV
    totalDataCalc
    

    Application.ScreenUpdating = True
    
    Sheets("sakura_past_trends").Select
    
    MsgBox ("グラフが作成されました")
    

End Sub


Public Sub RunGetWeatherForcast()

'   今年の気温データと予想値を取得してグラフを作成する。

    Application.ScreenUpdating = False
    
    GetWeatherCurrentYear
    ImportWeatherCurrentCSV
    GetWeatherForecast
    ImportWeatherForecastCSV
    totalCalcForecast
    

    Application.ScreenUpdating = True
    
    Sheets("sakura_current_situation").Select
    
    MsgBox ("グラフが作成されました")
    

End Sub

