Attribute VB_Name = "ModExchIDEX"
'https://docs.idex.market/#operation/returnCurrencies

Sub TestIDEX()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://docs.idex.market/
'Remember to create a new API key for excel/VBA

Dim Apikey As String

Apikey = "your api key here"

'Remove this lines, unless you define a constant somewhere ( Public Const apikey_idex = "the key to use everywhere" etc )
Apikey = apikey_idex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchIDEX"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase

Set Test = Suite.Test("TestIDEX")
'Error, unknown command
TestResult = PublicIDEX("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":"/AnUnknownCommand does not exist"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, missing parameter
TestResult = PublicIDEX("returnBalances", "GET")
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"error":"Invalid value for parameter: address"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("error"), "Invalid value for parameter: address"

'GET returnTicker
Dim Params As New Dictionary
Params.Add "market", "ETH_SAN"
TestResult = PublicIDEX("returnTicker", "GET", Params)
'{"last":"0.001303371681869683","high":"N/A","low":"N/A","percentChange":"0","baseVolume":"0","quoteVolume":"0","lowestAsk":"0.003560386590568989","highestBid":"0.001259340533790531"}
Test.IsOk InStr(TestResult, "baseVolume") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("highestBid") >= 0

'POST alternative returnTicker
Dim Params1 As New Dictionary
Params1.Add "market", "ETH_GET"
TestResult = PublicIDEX("returnTicker", "POST", Params1)
'{"last":"0.002246960662024021","high":"0.00224741010134807","low":"0.002063826","percentChange":"8.86593061","baseVolume":"18.619876495872901214","quoteVolume":"8922.19057388948559119","lowestAsk":"0.002245747624745226","highestBid":"0.002080526643055107"}
Test.IsOk InStr(TestResult, "quoteVolume") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("lowestAsk") >= 0





End Sub


Function PublicIDEX(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
Dim postdata As String
PublicApiSite = "https://api.idex.market"

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    postdata = DictToString(ParamDict, "JSON")
ElseIf UCase(ReqType) = "GET" Then
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    ApiEndPoint = ApiEndPoint & MethodParams
    postdata = ""
End If

urlPath = "/" & Method & MethodParams
Url = PublicApiSite & urlPath

Dim headerDict As New Dictionary
headerDict.Add "Content-Type", "application/json"

PublicIDEX = WebRequestURL(Url, ReqType, headerDict, postdata)

End Function
Function PrivateIDEX(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

'Work in Progress

End Function
