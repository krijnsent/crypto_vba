Attribute VB_Name = "ModSrcCoinGecko"
'Two variables for caching, so the formulas don't update every recalculation
Public Const CGCacheSeconds = 6000   'Nr of seconds cache, default >= 60
Public CGDict As New Scripting.Dictionary

Sub TestSrcCoinGecko()

'https://www.coingecko.com/en/api
'https://www.coingecko.com/api/documentations/v3

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModSrcCoinGecko"
' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestPublicCoinGeckoData")

'Test for errors first
TestResult = PublicCoinGeckoData("unknown_command")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":"Incorrect path. Please check https://www.coingecko.com/api/"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404
'Test.IsEqual JsonResult("response_txt")("error"), "Incorrect path. Please check https://www.coingecko.com/api/"

'Simple ping
TestResult = PublicCoinGeckoData("ping")
'{"gecko_says":"(V3) To the Moon!"}
Test.IsOk InStr(TestResult, "gecko") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("gecko_says"), "(V3) To the Moon!"

'Simple price
Dim Params As New Dictionary
Params.Add "ids", "bitcoin"
Params.Add "vs_currencies", "eur"
TestResult = PublicCoinGeckoData("simple/price", Params)
'e.g. {"bitcoin":{"eur":7272.72}}
Test.IsOk InStr(TestResult, "bitcoin") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("bitcoin")("eur") > 0


End Sub

Function PublicCoinGeckoData(Method As String, Optional ParamDict As Dictionary) As String

Dim url As String
Dim TempData As String
Dim Sec As Double

PublicApiSite = "https://api.coingecko.com/api/v3"

MethodParams = ""
If Not ParamDict Is Nothing Then
    'Change the rest of the parameters to JSON
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
End If

urlPath = Method & MethodParams
url = PublicApiSite & "/" & urlPath

GetNewData = False
IsInDict = CGDict.Exists(urlPath)
If IsInDict = True Then
    'In dictionary, check time
    If CGDict(urlPath) + TimeSerial(0, 0, CGCacheSeconds) < Now() Then
        'Has not been updated recently, update now
        CGDict.Remove urlPath
        CGDict.Add urlPath, Now()
        If CGDict.Exists("DATA-" & urlPath) Then CGDict.Remove "DATA-" & urlPath
        GetNewData = True
    End If
Else
    CGDict.Add urlPath, Now()
    GetNewData = True
End If

If GetNewData = True Then
    TempData = WebRequestURL(url, "GET")
    CGDict.Add "DATA-" & urlPath, TempData
Else
    TempData = CGDict("DATA-" & urlPath)
End If

PublicCoinGeckoData = TempData

End Function
