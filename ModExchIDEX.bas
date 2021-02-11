Attribute VB_Name = "ModExchIDEX"
'https://docs.idex.market/#operation/returnCurrencies

Sub TestIDEX()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://docs.idex.io/
'Remember to create a new API key for excel/VBA

Dim apiKey As String

apiKey = "your api key here"

'Remove this lines, unless you define a constant somewhere ( Public Const apikey_idex = "the key to use everywhere" etc )
apiKey = apikey_idex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey

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
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"code":"ResourceNotFound","message":"/AnUnknownCommand does not exist"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, missing parameter
TestResult = PublicIDEX("candles", "GET")
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"code":"REQUIRED_PARAMETER","message":"parameter \"market\" is required but was not provided"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("code"), "REQUIRED_PARAMETER"

'GET ticker
Dim Params As New Dictionary
Params.Add "market", "ZRX-ETH"
TestResult = PublicIDEX("tickers", "GET", Params)
'[{"market":"ZRX-ETH","time":1612898636288,"open":null,"high":null,"low":null,"close":null,"closeQuantity":null,"baseVolume":"0.00000000","quoteVolume":"0.00000000","percentChange":"0.00","numTrades":0,"ask":"0.00191918","bid":"0.00034900","sequence":null}]
Test.IsOk InStr(TestResult, "baseVolume") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult(1)("time") >= 0

End Sub


Function PublicIDEX(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
Dim postdata As String
PublicApiSite = "https://api-eth.idex.io/v1/"

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    postdata = JsonConverter.ConvertToJson(ParamDict)
ElseIf UCase(ReqType) = "GET" Then
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    ApiEndPoint = ApiEndPoint & MethodParams
    postdata = ""
End If

urlPath = "/" & Method & MethodParams
url = PublicApiSite & urlPath

Dim headerDict As New Dictionary
headerDict.Add "Content-Type", "application/json"

PublicIDEX = WebRequestURL(url, ReqType, headerDict, postdata)

End Function
Function PrivateIDEX(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

'Work in Progress

End Function
