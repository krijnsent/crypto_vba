Attribute VB_Name = "ModExchBitfinex"
Sub TestBitfinex()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://docs.bitfinex.com/docs/rest-auth
'Note: there are two versions, v1 and v2, v2 is in Beta and does not have all functions
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_bitfinex = "the key to use everywhere" etc )
Apikey = apikey_bitfinex
secretKey = secretkey_bitfinex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBitfinex"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase

Set Test = Suite.Test("TestBitfinexPublic v1")
'Error, unknown command
TestResult = PublicBitfinex1("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, wrong parameter
TestResult = PublicBitfinex1("ticker/bogus_here", "GET")
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"message":"Unknown symbol"}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("message"), "Unknown symbol"

'OK request
TestResult = PublicBitfinex1("symbols", "GET")
'["btcusd","ltcusd","ltcbtc","ethusd","ethbtc","etcbtc",
Test.IsOk InStr(TestResult, "ethbtc") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1), "btcusd"

'OK request with details
TestResult = PublicBitfinex1("stats/BTCUSD", "GET")
'[{"period":1,"volume":"6815.19360556"},{"period":7,"volume":"98002.43336128"},{"period":30,"volume":"387511.06628926"}]
Test.IsOk InStr(TestResult, "volume") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)("period"), 1
Test.IsOk JsonResult(1)("volume") > 0


Set Test = Suite.Test("TestBitfinexPrivate v1 Balances")
TestResult = PrivateBitfinex1("balances", "POST", Cred)
'[{"type":"exchange","currency":"btc","amount":"5.15334045","available":"5.15334045"},{"type":"exchange","currency":"eos","amount":"15.0","available":"15.0"}]
Test.IsOk InStr(TestResult, "currency") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk InStr(JsonResult(1)("type"), "exchange") + InStr(JsonResult(1)("type"), "margin") + InStr(JsonResult(1)("type"), "funding") > 0
Test.IsOk Len(JsonResult(1)("currency")) >= 3
Test.IsOk Len(JsonResult(1)("amount")) >= 0


Set Test = Suite.Test("TestBitfinexPrivate v1 Orders")
Dim Params1o As New Dictionary
Params1o.Add "symbol", "BTCUSD"
Params1o.Add "amount", "1.33"
Params1o.Add "price", "9"
Params1o.Add "side", "buy"
Params1o.Add "type", "fill-or-kill"
TestResult = PrivateBitfinex1("order/new", "POST", Cred, Params1o)
'e.g. {"error_nr":403,"error_txt":"HTTP-Forbidden","response_txt":{"message":"This API key does not have permission for this action"}}
'or: {"id":448364249,"symbol":"btcusd","exchange":"bitfinex",etc.
If InStr(TestResult, "error") > 0 Then
    Test.IsOk InStr(TestResult, "message") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsEqual JsonResult("error_nr"), 403
Else
    Test.IsOk InStr(TestResult, "symbol") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult("id") > 0
    Test.IsEqual JsonResult("symbol"), "btcusd"
End If


Set Test = Suite.Test("TestBitfinexPublic v2")

'Error, unknown command
TestResult = PublicBitfinex2("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, wrong parameter
TestResult = PublicBitfinex2("ticker/bogus_here", "GET")
'{"error_nr":500,"error_txt":"HTTP-","response_txt":["error",10020,"symbol: invalid"]}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 500
Test.IsEqual JsonResult("response_txt")(2), 10020

'OK request
TestResult = PublicBitfinex2("platform/status", "GET")
'[1] -> 1 = active, 0=maintenance
Test.IsOk InStr(TestResult, "]") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1), 1

'OK request with parameters
Dim Params As New Dictionary
Params.Add "symbols", "tBTCUSD,tNEOETH"
TestResult = PublicBitfinex2("tickers", "GET", Params)
'[["tBTCUSD",3907.1,34.68474518,3907.2,84.93216888,-24.5,-0.0062,3907.2,6790.69338403,3949,3838.89411809],["tNEOETH",0.065716,3437.62864427,0.06589,2087.26914816,0.000835,0.0129,0.065611,4944.19962337,0.068214,0.064699]]
Test.IsOk InStr(TestResult, "tNEOETH") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)(1), "tBTCUSD"
Test.IsOk JsonResult(1)(2) > 100

Set Test = Suite.Test("TestBitfinexPublic v2 POST")
'OK POST request with parameters, no credentials needed
Dim Params2 As New Dictionary
Params2.Add "symbol", "tBTCUSD"
Params2.Add "amount", "-2.5"
TestResult = PublicBitfinex2("calc/trade/avg", "POST", Params2)
'[3905,-2.5]
Test.IsOk InStr(TestResult, "]") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult(1) > 100
Test.IsEqual JsonResult(2), -2.5


Set Test = Suite.Test("TestBitfinexPrivate v2 Wallets")
TestResult = PrivateBitfinex2("auth/r/wallets", "POST", Cred)
'e.g. [["exchange","BTC",5.15334045,0,null],["exchange","EOS",15,0,null]]
Test.IsOk InStr(TestResult, "]]") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
'Test first result for being one of three types exchange, margin, funding
Test.IsOk InStr(JsonResult(1)(1), "exchange") + InStr(JsonResult(1)(1), "margin") + InStr(JsonResult(1)(1), "funding") > 0
Test.IsOk Len(JsonResult(1)(2)) >= 3
Test.IsOk Len(JsonResult(1)(3)) >= 0

Set Test = Suite.Test("TestBitfinexPrivate v2 Trades")
'Unix time period (add 3 zeros for ms):
t1 = DateToUnixTime("1/1/2016") & "000"
t2 = DateToUnixTime("1/1/2018") & "000"

Dim Params3 As New Dictionary
Params3.Add "start", t1
Params3.Add "end", t2
Params3.Add "limit", 25
TestResult = PrivateBitfinex2("auth/r/ledgers/BTC/hist", "POST", Cred, Params3)
'[] for empty or [[ID,CURRENCY,null,TIMESTAMP_MILLI,null,AMOUNT,BALANCE,null,Description]]
Test.IsOk InStr(TestResult, "]") > 0
If Len(TestResult) > 2 Then
    'Results, some more tests
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult(1)(1) > 0
    Test.IsOk Len(JsonResult(1)(2)) >= 3
    Test.IsOk JsonResult(1)(4) > 1400000000000#
End If


End Sub

'Version 2 APIs below
Function PublicBitfinex1(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.bitfinex.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/v1/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBitfinex1 = WebRequestURL(url, ReqType)

End Function
Function PrivateBitfinex1(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

'Thanks to balin77!
Dim NonceUnique As String
Dim TimeCorrection As Long
Dim url As String

NonceUnique = CreateNonce(15)
TradeApiSite = "https://api.bitfinex.com"
ApiPath = "/v1/" & Method

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams

Set PayloadDict = New Dictionary
PayloadDict("request") = ApiPath
PayloadDict("nonce") = NonceUnique
If Not ParamDict Is Nothing Then
    For Each key In ParamDict.Keys
        PayloadDict(key) = ParamDict(key)
    Next key
End If
    
Json = Replace(ConvertToJson(PayloadDict), "/", "\/")
payload = Base64Encode(Json)
APIsign = ComputeHash_C("SHA384", payload, Credentials("secretKey"), "STRHEX")

url = TradeApiSite & ApiPath

Dim UrlHeaders As New Dictionary
UrlHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
UrlHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
UrlHeaders.Add "X-BFX-APIKEY", Credentials("apiKey")
UrlHeaders.Add "X-BFX-PAYLOAD", payload
UrlHeaders.Add "X-BFX-SIGNATURE", APIsign
PrivateBitfinex1 = WebRequestURL(url, ReqType, UrlHeaders)

End Function


Function PublicBitfinex2(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api-pub.bitfinex.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/v2/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBitfinex2 = WebRequestURL(url, ReqType)

End Function
Function PrivateBitfinex2(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim TimeCorrection As Long
Dim url As String

NonceUnique = CreateNonce(15)
TradeApiSite = "https://api.bitfinex.com/"
ApiPath = "v2/" & Method

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams

ToSign = "/api/" & ApiPath & NonceUnique
APIsign = ComputeHash_C("SHA384", ToSign, Credentials("secretKey"), "STRHEX")

url = TradeApiSite & ApiPath & MethodParams

Dim UrlHeaders As New Dictionary
UrlHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
UrlHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
UrlHeaders.Add "bfx-nonce", NonceUnique
UrlHeaders.Add "bfx-apikey", Credentials("apiKey")
UrlHeaders.Add "bfx-signature", APIsign
PrivateBitfinex2 = WebRequestURL(url, ReqType, UrlHeaders)


End Function
