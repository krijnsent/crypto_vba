Attribute VB_Name = "ModExchOkex"
Sub TestOKEx()

'Source: https://github.com/krijnsent/crypto_vba
'https://www.okex.com/docs/en/
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_okex = "the key to use everywhere" etc )
Apikey = apikey_okex
secretKey = secretkey_okex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchOKEx"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestOKExPublic")

'Error, unknown command
TestResult = PublicOKEx("AnUnknownCommand", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, missing parameter
TestResult = PublicOKEx("instruments/EOS-BTC/", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, unknown pair
TestResult = PublicOKEx("instruments/EOS-BLA/ticker", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("code"), 30032
Test.IsEqual JsonResult("response_txt")("message"), "The currency pair does not exist"

TestResult = PublicOKEx("instruments/ticker", "GET")
'[{"best_ask":"0.006388","best_bid":"0.006387","instrument_id":"LTC-BTC","product_id":"LTC-BTC","last":"0.006387","ask":"0.006388","bid":"0.006387","open_24h":"0.006532","high_24h":"0.006727","low_24h":"0.006359","base_volume_24h":"221873.685698","timestamp":"2019-09-04T09:31:50.304Z","quote_volume_24h":"1445.8081"},{"best_ask":"0.01685","best_bid":"0.01684","instrument_id":"ETH-BTC","product_id":"ETH-BTC" etc...
Test.IsOk InStr(TestResult, "best_bid") > 0
Test.IsOk InStr(TestResult, "product_id") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count > 100
Test.IsOk JsonResult(1)("last") * 1 > 0
Test.IsOk Len(JsonResult(2)("instrument_id")) > 0


Dim Params As New Dictionary
Params.Add "granularity", 14400  '14400 seconds = 6 hours
Params.Add "start", "2019-03-18T08%3A28%3A48.899Z"  'ISO 8601
Params.Add "end", "2019-03-19T09%3A28%3A48.899Z"
TestResult = PublicOKEx("instruments/ETH-USDT/candles", "GET", Params)
'Result: TOHLCV
'[["2019-03-19T08:00:00.000Z","137.74","138.69","137.38","137.79","107365.43315"],["2019-03-19T04:00:00.000Z","137.76","138","136.97","137.73","85020.919026"],["2019-03-19T00:00:00.000Z","137.61","139.41","137.31","137.74","94292.72983"],["2019-03-18T20:00:00.000Z","137.44","138.33","137.42","137.63","63587.691327"],["2019-03-18T16:00:00.000Z","137.59","137.91","137.09","137.42","58001.277483"],["2019-03-18T12:00:00.000Z","137.27","138.03","137","137.6","83512.951662"]]
Test.IsOk InStr(TestResult, "2019-03-19") > 0
Test.IsOk InStr(TestResult, "2019-03-18") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)(1), "2019-03-19T08:00:00.000Z"
Test.IsEqual JsonResult(1)(2), "137.74"
Test.IsEqual JsonResult.Count, 6


Set Test = Suite.Test("TestOKExPrivate")

TestResult = PrivateOKEx("accounts", "GET", Cred)
'WORK IN PROGRESS



End Sub

Function PublicOKEx(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://www.okex.com/api/spot/v3"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/" & Method & MethodParams
Url = PublicApiSite & urlPath

PublicOKEx = WebRequestURL(Url, ReqType)

End Function
Function PrivateOKEx(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim Url As String
Dim postdata As String

TradeApiSite = "https://www.okex.com"
ApiEndPoint = "/api/" & Method

'WORK IN PROGRESS

Dim headerDict As New Dictionary
headerDict.Add "Content-Type", "application/json"

PrivateOKEx = WebRequestURL(Url, ReqType, headerDict, postdata)

End Function
