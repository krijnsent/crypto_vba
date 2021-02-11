Attribute VB_Name = "ModExchBybit"
Sub TestBybit()

'Source: https://github.com/krijnsent/crypto_vba
'https://doc.Bybit.co.kr/#section/V2-version
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_bybit
secretKey = secretkey_bybit

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
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

Dim Params1a As New Dictionary
Dim LimitTime As Double
Dim ResTime As Long
Params1a.Add "symbol", "BTCUSD"
Params1a.Add "interval", 60 'TimeFrame in minutes
Params1a.Add "limit", 2
LimitTime = Round(GetBybitTime() / 1000, 0) - 60 * 60 * 2
'GetByBitTime returns time in ms (microseconds, 13 digits), and this function takes seconds (10 digits)
'In order to get the past 2 hours, deduct that time in seconds: interval*limit*60
Params1a.Add "from", LimitTime
TestResult = PublicBybit("kline/list", "GET", Params1a)
Test.IsEqual JsonResult("ret_msg"), "OK", "kline 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result").Count > 0, "kline 2 failed, result: ${1}"
Test.IsEqual JsonResult("result")(1)("symbol"), "BTCUSD", "kline 3 failed, result: ${1}"
Test.IsOk Val(JsonResult("result")(1)("high")) > 0, "kline 4 failed, result: ${1}"
'ResTime = JsonResult("result")(1)("open_time")
'Debug.Print ResTime, UnixTimeToDate(ResTime)
'ResTime = JsonResult("result")(2)("open_time")
'Debug.Print ResTime, UnixTimeToDate(ResTime)

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
'Debug.Print TestResult
Set JsonResult = JsonConverter.ParseJson(TestResult)

If JsonResult("ret_msg") = "ok" Then
    Test.IsOk Len(JsonResult("result")(1)("api_key")) >= 10
    Test.IsOk JsonResult("result")(1)("user_id") > 0
Else
    'E.g. IP-address block
    Test.IsEqual Left(JsonResult("ret_msg"), 12), "unmatched IP"
    Test.IsUndefined JsonResult("result")
End If

'Example set leverage
Dim Params2 As New Dictionary
Params2.Add "symbol", "ETHUSD"
Params2.Add "leverage", 1
TestResult = PrivateBybit("user/leverage/save", "POST", Cred, Params2)
'Debug.Print TestResult
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

'Example set leverage
Tm = GetBybitTime()
'e.g. 1583401823000 -> milliseconds
Hrs = 24
Dim Params3 As New Dictionary
Params3.Add "symbol", "ETHUSD"
Params3.Add "limit", 1
Params3.Add "start_time", Tm - 3600000 * Hrs
TestResult = PrivateBybit("v2/private/execution/list", "GET", Cred, Params3)
'Debug.Print Tm
'Debug.Print TestResult
'{"ret_code":0,"ret_msg":"OK","ext_code":"","ext_info":"","result":{"order_id":"","trade_list":null},"time_now":"1583400361.063716","rate_limit_status":119,"rate_limit_reset_ms":1583400361061,"rate_limit":120}


'/v2/private/execution/list

End Sub

Function PublicBybit(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.bybit.com/v2/public/"
 
'symbols, orderBook/L2  +symbol , time, tickers (+symbol)

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = Method & MethodParams
url = PublicApiSite & urlPath

PublicBybit = WebRequestURL(url, ReqType)

End Function
Function PrivateBybit(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim postdataUrl As String
Dim postdataJSON As String
Dim url As String

'Get a 10-digit Nonce
NonceUnique = GetBybitTime()
TradeApiSite = "https://api.bybit.com/"

url = TradeApiSite & Method

Dim PostDict As New Dictionary
PostDict.Add "api_key", Credentials("apiKey")
PostDict.Add "timestamp", NonceUnique
If Not ParamDict Is Nothing Then
    For Each key In ParamDict.Keys
        PostDict(key) = ParamDict(key)
    Next key
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
    postdataJSON = JsonConverter.ConvertToJson(ParamDict)
    contentFormat = "application/json"
    MethodParams = ""
Else
    'Wrong Method, error out
    Exit Function
End If

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", contentFormat

url = TradeApiSite & Method & MethodParams

PrivateBybit = WebRequestURL(url, ReqType, headerDict, postdataJSON)

End Function


Function GetBybitTime() As Double

Dim BybitTime As String
Dim ValBybitTime As Double
Dim JsonResponse As String
Dim Json As Object

'PublicBybit time, 13 digit (ms)
JsonResponse = PublicBybit("time", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
BybitTime = Left(Json("time_now"), InStr(Json("time_now"), ".") - 1) & "000"

If Len(BybitTime) = 0 Then
    TimeCorrection = -3600
    ValBybitTime = DateDiff("s", "1/1/1970", Now) + TimeCorrection
    BybitTime = Trim(Str(ValBybitTime) & Right(Int(Timer * 100), 2) & "0")
End If

'Debug.Print BybitTime

GetBybitTime = Val(BybitTime)

Set Json = Nothing

End Function
