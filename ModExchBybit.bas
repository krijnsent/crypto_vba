Attribute VB_Name = "ModExchBybit"
Sub TestBybit()

'Source: https://github.com/krijnsent/crypto_vba
'https://doc.Bybit.co.kr/#section/V2-version
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_bybit
secretKey = secretkey_bybit

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBybit"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBybitPublic")

'Error, unknown command
TestResult = PublicBybit("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0, "unknowncommand 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_txt"), "HTTP-Not Found", "unknowncommand 2 failed, result: ${1}"
Test.IsEqual JsonResult("error_nr"), 404, "unknowncommand 3 failed, result: ${1}"

'OK request
TestResult = PublicBybit("time", "GET")
'e.g. {"ret_code":0,"ret_msg":"OK","ext_code":"","ext_info":"","result":{},"time_now":"1572094930.589837"}
Test.IsOk InStr(TestResult, "time_now") > 0, "time 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("ret_msg"), "OK", "time 1 failed, result: ${1}"
Test.IsOk Val(JsonResult("time_now")) > 1500000000#, "time 1 failed, result: ${1}"

'GET with parameter for orderBook
Dim Params1 As New Dictionary
Params1.Add "symbol", "BTCUSD"
TestResult = PublicBybit("orderBook/L2", "GET", Params1)
'e.g {"ret_code":0,"ret_msg":"OK","ext_code":"","ext_info":"","result":[{"symbol":"BTCUSD","price":"9094","size":214217,"side":"Buy"},{"symbol":"BTCUSD","price":"9093","size":208793,"side":"Buy"},{"symbol":"BTCUSD","price":"9092","size":208793,"side":"Buy"},{"symbol":"BTCUSD","price":"9086","size":1,"side":"Buy"},{"symbol":"BTCUSD","price":"9077","size":3855,"side":"Buy"},{"symbol":"BTCUSD","price":"9076","size":2500,"side":"Buy"},{"symbol":"BTCUSD","price":"9075","size":1515,"side":"Buy"},{"symbol":"BTCUSD","price":"9074","size":11419,"side":"Buy"},{"symbol":"BTCUSD","price":"9073","size":500,"side":"Buy"},{"symbol":"BTCUSD","price":"9070.5","size":727,"side":"Buy"},{"symbol":"BTCUSD","price":"9070","size":6786,"side":"Buy"},{"symbol":"BTCUSD","price":"9068","size":10057,"side":"Buy"},{"symbol":"BTCUSD","price":"9067.5","size":5200,"side":"Buy"},{"symbol":"BTCUSD","price":"9067","size":50,"side":"Buy"},{"symbol":"BTCUSD","price":"9066.5","size":433,"side":"Buy"},
Test.IsEqual JsonResult("ret_msg"), "OK", "orderbook 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result").Count > 0, "orderbook 2 failed, result: ${1}"
Test.IsEqual JsonResult("result")(1)("symbol"), "BTCUSD", "orderbook 3 failed, result: ${1}"
Test.Includes Array("Buy", "Sell"), JsonResult("result")(1)("side"), "orderbook 4 failed, result: ${1}"

'GET all tickers -> add a parameter like above to only get one
TestResult = PublicBybit("tickers", "GET")
'e.g. {"ret_code":0,"ret_msg":"OK","ext_code":"","ext_info":"","result":[{"symbol":"BTCUSD","bid_price":"9176","ask_price":"9176.5","last_price":"9176.00","last_tick_direction":"MinusTick","prev_price_24h":"7624.50","price_24h_pcnt":"0.203488","high_price_24h":"10558.00","low_price_24h":"7624.00","prev_price_1h":"9250.50","price_1h_pcnt":"-0.008053","mark_price":"9174.56","index_price":"9174.02","open_interest":98256174,"open_value":"10936.65","total_turnover":"11422803.74","turnover_24h":"476498.44","total_volume":106760806255,"volume_24h":4369471987,"funding_rate":"0.000168","predicted_funding_rate":"0.000352","next_funding_time":"2019-10-26T16:00:00Z","countdown_hour":3},{"symbol":"ETHUSD","bid_price":"180.2","ask_price":"180.25","last_price":"180.20","last_tick_direction":"MinusTick","prev_price_24h":"166.60","price_24h_pcnt":"0.081632","high_price_24h":"199.85","low_price_24h":"166.50","prev_price_1h":"181.65","price_1h_pcnt":"-0.007982","mark_price":"180.49","index_price":"180.48",
Test.IsEqual JsonResult("ret_msg"), "OK", "tickers 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result").Count > 0, "tickers 2 failed, result: ${1}"
Test.IsOk Val(JsonResult("result")(1)("bid_price")) > 0, "tickers 3 failed, result: ${1}"
Test.IsOk Val(JsonResult("result")(1)("prev_price_24h")) > 0, "tickers 4 failed, result: ${1}"

' Create a new test
Set Test = Suite.Test("TestBybitTime")
TestResult = GetBybitTime()
Test.IsOk TestResult > 1500000000000#, "bybit time 1 failed, result: ${1}"
Test.IsOk TestResult < 1600000000000#, "bybit time 2 failed, result: ${1}"


Set Test = Suite.Test("TestBybitPrivate")

'Api key properties
TestResult = PrivateBybit("open-api/api-key", "GET", Cred)
'e.g. {"ret_code":0,"ret_msg":"ok","ext_code":"","result":[{"api_key":"Tc5aI32WaSqSD","user_id":619,"ips":["192.168.1.1"],"note":"ExcelBybit","permissions":["Order","Position"],"created_at":"2019-10-26T10:16:38.000Z","read_only":false}],"ext_info":null,"time_now":"1572103275.354790","rate_limit_status":99,"rate_limit_reset":1572103275}
Test.IsOk InStr(TestResult, "ret_msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("ret_msg"), "ok"
Test.IsOk Len(JsonResult("result")(1)("api_key")) >= 10
Test.IsOk JsonResult("result")(1)("user_id") > 0

'More tests & examples to follow
Dim Params2 As New Dictionary
Params2.Add "symbol", "ETHUSD"
Params2.Add "leverage", 1
TestResult = PrivateBybit("user/leverage/save", "POST", Cred, Params2)
'e.g. {"ret_code":0,"ret_msg":"ok","ext_code":"","result":2,"ext_info":null,"time_now":"1572104006.055933","rate_limit_status":74,"rate_limit_reset":1572104006}
'or {"ret_code":34015,"ret_msg":"cannot set leverage which is same to the old leverage","ext_code":"","result":null,"ext_info":null,"time_now":"1572103987.614015","rate_limit_status":72,"rate_limit_reset":1572103987}
Test.IsOk InStr(TestResult, "ret_msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
If JsonResult("ret_msg") = "ok" Then
    Test.IsEqual JsonResult("ret_msg"), "ok"
    Test.IsEqual JsonResult("result"), 1 'same as input leverage
Else
    'Assume leverage is the same as before
    Test.IsEqual JsonResult("ret_msg"), "cannot set leverage which is same to the old leverage"
    Test.IsUndefined JsonResult("result")
End If



End Sub

Function PublicBybit(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://api.bybit.com/v2/public/"
 
'symbols, orderBook/L2  +symbol , time, tickers (+symbol)

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = Method & MethodParams
Url = PublicApiSite & urlPath

PublicBybit = WebRequestURL(Url, ReqType)

End Function
Function PrivateBybit(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim postdataUrl As String
Dim postdataJSON As String
Dim Url As String

'Get a 10-digit Nonce
NonceUnique = GetBybitTime()
TradeApiSite = "https://api.bybit.com/"

Url = TradeApiSite & Method

Dim PostDict As New Dictionary
PostDict.Add "api_key", Credentials("apiKey")
PostDict.Add "timestamp", NonceUnique
If Not ParamDict Is Nothing Then
    For Each Key In ParamDict.Keys
        PostDict(Key) = ParamDict(Key)
    Next Key
End If
'Sort alphabetically
Call SortDictByKey(PostDict)

'All parameters are in the PostDict dictionary, merge them to a string
MsgToSign = DictToString(PostDict, "URLENC")
APIsign = ComputeHash_C("SHA256", MsgToSign, Credentials("secretKey"), "STRHEX")
PostDict.Add "sign", APIsign

If UCase(ReqType) = "GET" Then
    MethodParams = DictToString(PostDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    contentFormat = "application/x-www-form-urlencoded"
ElseIf UCase(ReqType) = "POST" Then
    postdataJSON = DictToString(PostDict, "JSON")
    contentFormat = "application/json"
    MethodParams = ""
Else
    'Wrong Method, error out
    Exit Function
End If

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", contentFormat

Url = TradeApiSite & Method & MethodParams

PrivateBybit = WebRequestURL(Url, ReqType, headerDict, postdataJSON)
A = 1

End Function


Function GetBybitTime() As Double

Dim JsonResponse As String
Dim json As Object

'PublicBybit time, 10 digit
JsonResponse = PublicBybit("time", "GET")
Set json = JsonConverter.ParseJson(JsonResponse)
GetBybitTime = Left(json("time_now"), InStr(json("time_now"), ".")) & "000"
If GetBybitTime = 0 Then
    TimeCorrection = -3600
    GetBybitTime = DateDiff("s", "1/1/1970", Now)
    GetBybitTime = Trim(Str((Val(GetBybitTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set json = Nothing

End Function
