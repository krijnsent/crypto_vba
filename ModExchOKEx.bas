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
passphrase = passphrase_okex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey
Cred.Add "Passphrase", passphrase

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchOKEx"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestOKExPublic")

'Error, unknown command, returns invalid JSON
TestResult = PublicOKEx("AnUnknownCommand", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 403

'Error, missing parameter
TestResult = PublicOKEx("spot/v3/instruments/EOS-BTC/", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, unknown pair
TestResult = PublicOKEx("spot/v3/instruments/EOS-BLA/ticker", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("code"), 30032
Test.IsEqual JsonResult("response_txt")("message"), "The currency pair does not exist"

TestResult = PublicOKEx("spot/v3/instruments/ticker", "GET")
'[{"best_ask":"0.006388","best_bid":"0.006387","instrument_id":"LTC-BTC","product_id":"LTC-BTC","last":"0.006387","ask":"0.006388","bid":"0.006387","open_24h":"0.006532","high_24h":"0.006727","low_24h":"0.006359","base_volume_24h":"221873.685698","timestamp":"2019-09-04T09:31:50.304Z","quote_volume_24h":"1445.8081"},{"best_ask":"0.01685","best_bid":"0.01684","instrument_id":"ETH-BTC","product_id":"ETH-BTC" etc...
Test.IsOk InStr(TestResult, "best_bid") > 0
Test.IsOk InStr(TestResult, "product_id") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count > 100
Test.IsOk JsonResult(1)("last") * 1 > 0
Test.IsOk Len(JsonResult(2)("instrument_id")) > 0


Dim Params As New Dictionary
Params.Add "granularity", 14400  '14400 seconds = 6 hours
Params.Add "start", "2020-12-15T08%3A28%3A48.899Z"  'ISO 8601
Params.Add "end", "2020-12-19T09%3A28%3A48.899Z"
TestResult = PublicOKEx("spot/v3/instruments/ETH-USDT/candles", "GET", Params)
'Result: TOHLCV
'[["2019-03-19T08:00:00.000Z","137.74","138.69","137.38","137.79","107365.43315"],["2019-03-19T04:00:00.000Z","137.76","138","136.97","137.73","85020.919026"],["2019-03-19T00:00:00.000Z","137.61","139.41","137.31","137.74","94292.72983"],["2019-03-18T20:00:00.000Z","137.44","138.33","137.42","137.63","63587.691327"],["2019-03-18T16:00:00.000Z","137.59","137.91","137.09","137.42","58001.277483"],["2019-03-18T12:00:00.000Z","137.27","138.03","137","137.6","83512.951662"]]
Test.IsOk InStr(TestResult, "2020-12-19") > 0
Test.IsOk InStr(TestResult, "2020-12-18") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)(1), "2020-12-19T08:00:00.000Z"
Test.IsEqual JsonResult(1)(2), "648.67"
Test.IsEqual JsonResult.Count, 24

' Create a new test
Set Test = Suite.Test("TestOKExTime")
TestResult = GetOKExTime()
Test.IsOk TestResult > 1500000000#
Test.IsOk TestResult < 1700000000#


Set Test = Suite.Test("TestOKExPrivate")

TestResult = PrivateOKEx("spot/v3/accounts", "GET", Cred)
'[{"frozen":"0","hold":"0","id":"","currency":"BTC","balance":"0","available":"0","holds":"0"},{"frozen":"0","hold":"0","id":"","currency":"XAS","balance":"0.000233","available":"0.000233","holds":"0"}]
Test.IsOk InStr(TestResult, "currency") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count >= 2
Test.IsOk JsonResult(1)("balance") * 1 >= 0
Test.IsOk JsonResult(1)("holds") * 1 >= 0

'Invalid token
TestResult = PrivateOKEx("account/v3/wallet/BLA", "GET", Cred)
'{"error_nr":400,"error_txt":"HTTP-","response_txt":{"code":30031,"message":"BLA is an invalid token"}}
Test.IsOk InStr(TestResult, "error") > 0
Test.IsOk InStr(TestResult, 30031) > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("response_txt")("code"), 30031
Test.IsEqual JsonResult("response_txt")("message"), "BLA is an invalid token"

Set Test = Suite.Test("TestOKExPrivate Orders")
'Create order
'BUY 100 BTC for a price of 1 USDT per BTC
'price hopefully insane enough never to execute
Dim Params2 As New Dictionary
Params2.Add "instrument_id", "BTC-USDT"
Params2.Add "type", "limit"
Params2.Add "side", "buy"
Params2.Add "price", 1
Params2.Add "size", 100
Params2.Add "order_type", 3 '3-Immediate Or Cancel
TestResult = PrivateOKEx("spot/v3/orders", "POST", Cred, Params2)
'e.g. {"client_oid":"","error_code":"33017","error_message":"Greater than the maximum available balance","order_id":"-1","result":false}
Test.IsOk InStr(TestResult, "code") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("error_code") * 1 = 33017
Test.IsOk JsonResult("error_message") = "Greater than the maximum available balance"

Dim Params3 As New Dictionary
Params3.Add "instrument_id", "XMR-BTC"
ClientIdOrderId = "12345"
TestResult = PrivateOKEx("spot/v3/cancel_orders/" & ClientIdOrderId, "POST", Cred, Params3)
'{"client_oid":"","code":"33014","error_code":"33014","error_message":"Order does not exist","message":"Order does not exist","order_id":"12345","result":false}
Test.IsOk InStr(TestResult, "code") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 = 33014
Test.IsOk JsonResult("message") = "Order does not exist"

Dim Params4 As New Dictionary
Params4.Add "instrument_id", "XMR-BTC"
Params4.Add "limit", 4
TestResult = PrivateOKEx("spot/v3/orders_pending", "GET", Cred, Params4)
'[] (no orders), [[{"client_oid":"oktspot86","created_at":"2019-03-20T03:28:14.000Z",etc...
If TestResult = "[]" Then
    'No orders
Else
    Test.IsOk InStr(TestResult, "created_at") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult(1)("instrument_id") = "XMR-BTC"
    Test.IsOk Len(JsonResult(1)("order_id")) > 0
End If

Dim Params5 As New Dictionary
instrument_id = "XAS-BTC"
Params5.Add "instrument_id", instrument_id
JsonResponse = PrivateOKEx("spot/v3/fills", "GET", Cred, Params5)
'e.g. []
If TestResult = "[]" Then
    'No orders
Else
    Test.IsOk InStr(TestResult, "created_at") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult(1)("instrument_id") = "XAS-BTC"
End If


End Sub

Function PublicOKEx(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://www.okex.com/api"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicOKEx = WebRequestURL(url, ReqType)

End Function
Function PrivateOKEx(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim url As String
Dim postdata As String

TradeApiSite = "https://www.okex.com"
ApiEndPoint = "/api/" & Method

'OKEx nonce
NonceUnique = GetOKExTime() & ".00" 'Should be string

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    postdata = JsonConverter.ConvertToJson(ParamDict)
ElseIf UCase(ReqType) = "GET" Then
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    ApiEndPoint = ApiEndPoint & MethodParams
End If

ApiForSign = NonceUnique & UCase(ReqType) & ApiEndPoint & postdata
APIsign = ComputeHash_C("SHA256", ApiForSign, Credentials("secretKey"), "STR64")

url = TradeApiSite & ApiEndPoint

Dim headerDict As New Dictionary
headerDict.Add "OK-ACCESS-KEY", Credentials("apiKey")
headerDict.Add "OK-ACCESS-SIGN", APIsign
headerDict.Add "OK-ACCESS-TIMESTAMP", NonceUnique
headerDict.Add "OK-ACCESS-PASSPHRASE", Credentials("Passphrase")
headerDict.Add "Content-Type", "application/json"

PrivateOKEx = WebRequestURL(url, ReqType, headerDict, postdata)

End Function

Function GetOKExTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicOKEx time
JsonResponse = PublicOKEx("general/v3/time", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
If InStr(Json("epoch"), ".") Then
    GetOKExTime = Left(Json("epoch"), InStr(Json("epoch"), "."))
Else
    GetOKExTime = Json("epoch")
End If
If GetOKExTime = 0 Then
    TimeCorrection = -3600
    GetOKExTime = DateDiff("s", "1/1/1970", Now)
    GetOKExTime = Trim(Str((Val(GetOKExTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set Json = Nothing

End Function
